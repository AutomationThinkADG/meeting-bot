# Speaker attribution — `meeting-bot` side

How the bot captures **who is speaking** and ships it to `copilot-bot-api` so
the final transcript carries real names instead of `Speaker 1 / Speaker 2`.

The API-side half (fusion against Azure diarization + roster) is documented in
`copilot-bot-api/docs/SPEAKER_ATTRIBUTION.md`. Read both.

---

## Why this exists

The bot path (Zoom / Meet / Teams-guest / other-tenant) records **audio only**.
Azure Speech batch diarization then labels utterances `Speaker 0/1/2…` with no
identity. The previous attempt tried to recover names by scraping a Teams DOM
"is speaking" attribute that does not exist, and wrote the result to a
container-local file that never reached the API. Net result: every bot meeting
came out as `Speaker N`.

This design fixes three things:

1. **The bot prepares the meeting itself** — it turns on its *own* live captions
   (a per-participant setting, invisible to others, no host action, works for
   anonymous/guest joins), opens the People pane, and forces Gallery view.
2. **It captures a real signal** — live-caption speaker labels (the conferencing
   server's own attribution) plus a voice-level-outline fallback.
3. **The signal actually travels** — it rides the recording-completed webhook to
   the API, which persists it on the meeting row.

---

## Data flow (order of operations)

```
MicrosoftTeamsBot.join()
  └─ joinMeeting()
       ├─ … join lobby, mute, camera off, dismiss dialogs …
       ├─ runTeamsReadiness()                         ← src/lib/meetingReadiness.ts
       │     • open People pane        (roster subtree exists to scrape)
       │     • force Gallery view      (voice rings / tiles render)
       │     • enable live captions    (+ set spoken language)
       │     → MeetingReadinessReport
       └─ recordMeetingPageWithFFmpeg()
             ├─ recorder.start()  →  recordingStartedAt = Date.now()
             ├─ startTeamsSpeakerCapture(page, …)      ← src/lib/speakerCapture.ts
             │     in-page: 1 MutationObserver (flag only) + 500ms parse tick
             │     • caption lines  → { kind:'caption', speaker, text, tSec }
             │     • outline rings   → { kind:'outline', name, tSec }
             │     • People pane     → { kind:'roster', names }
             │     Node accumulates SpeakerSpan[] + CaptionCue[] + roster
             ├─ … record until meeting end / silence / max duration …
             └─ finally:
                   ├─ recorder.stop()
                   ├─ capture.stop() + capture.snapshot()
                   ├─ uploader.setSpeakerTimeline(spans)
                   ├─ uploader.setCaptionTranscript(cues)   (if under size cap)
                   ├─ uploader.setReadinessReport(report)
                   └─ stage recording, close browser

DiskUploader.uploadRecordingToRemoteStorage()
  └─ on success → notifyRecordingCompleted()
        └─ sendWebhook(RecordingCompletedPayload)
              payload.metadata.speakerTimeline    : SpeakerSpan[]
              payload.metadata.captionTranscript  : CaptionCue[]   (optional)
              payload.metadata.readinessReport    : MeetingReadinessReport
                     │
                     ▼  HTTP POST  NOTIFY_WEBHOOK_URL  (= copilot-bot-api /bot-webhook)
```

The webhook is a plain HTTP POST (no size limit); the API stores the payload on
the meeting row and keeps only `{recordingUrl, botId, timestamp}` on its
internal 64 KB storage queue.

---

## The contract (`src/types.ts`)

```ts
interface SpeakerSpan {
  tStartSec: number;
  tEndSec: number | null;          // null => still open at meeting end
  name: string | null;             // display name when the source knows it
  source: 'teams-caption' | 'teams-outline' | 'teams-roster-solo'
        | 'meet-caption' | 'zoom-tile';
  confidence: number;              // 0..1
}

interface CaptionCue {
  speaker: string | null;
  text: string;                    // powers the API's skew-invariant text vote
  tStartSec: number;
  tEndSec: number | null;
}

interface MeetingReadinessReport {
  provider: 'microsoft' | 'google' | 'zoom';
  captionsRequested: boolean;
  captionsEnabled: boolean;
  captionStatus: 'ok' | 'sparse' | 'unavailable' | 'policy-blocked' | 'error';
  captionLanguage?: string;
  captionEventCount: number;
  rosterCaptured: boolean;
  rosterNames: string[];
  peoplePaneOpened: boolean;
  galleryViewForced: boolean;
  participantCountSeen: number | null;
  outlineEventCount: number;
  notes: string[];
}
```

All timestamps are **seconds from recording start** (same clock as ffmpeg and
Azure `offsetInTicks / 1e7`). The bot does not try to be frame-accurate — a
one-off clock-skew correction happens on the API side.

---

## Function reference

| File | Export | Responsibility |
|---|---|---|
| `src/lib/meetingReadiness.ts` | `runTeamsReadiness(ctx)` | Best-effort Playwright UI actions: open People pane, force Gallery view, enable captions + language. Returns `MeetingReadinessReport`. Never throws. |
| `src/lib/speakerCapture.ts` | `startTeamsSpeakerCapture(page, logger, opts)` | Exposes `screenAppSpeakerEvent`, injects the in-page observer, accumulates `SpeakerSpan[]` / `CaptionCue[]` / roster. Returns `{ snapshot(), stop() }`. |
| `src/bots/MicrosoftTeamsBot.ts` | — | Calls the two above; hands the snapshot to the uploader in the recording `finally`. |
| `src/middleware/disk-uploader.ts` | `setSpeakerTimeline` / `setCaptionTranscript` / `setReadinessReport` | Stash on the uploader; injected into `RecordingCompletedPayload.metadata`. `setCaptionTranscript` drops the transcript above `captionTranscriptMaxBytes`. |

### In-page capture internals (`injectedCapture`)

- **Captions** — poll `captionContainerSel` → `captionLineSel` every 500ms.
  Track each line element (WeakMap). A line is *final* when its text has been
  unchanged for `captionFinalizeQuietMs` **or** it has left the DOM. One
  `MutationObserver` (`childList`/`subtree`, no `characterData`) is kept wired to
  notice a container swap; it does no parsing.
- **Outline** — `MutationObserver` on `style`/`class` of every `outlineSel` node;
  count mutations in a 1.5s sliding window; the ring with ≥3 recent mutations is
  the active speaker; name via the enclosing tile (`data-tid` value, then
  `aria-label`).
- **Roster** — scrape People-pane rows every ~15s.
- The bot's own display name is filtered out of every signal.

---

## Configuration (`src/config.ts`, all env-overridable)

| Env | Default | Purpose |
|---|---|---|
| `TEAMS_READINESS_ENABLED` | `true` | Master switch for the readiness routine |
| `TEAMS_ENABLE_CAPTIONS` | `true` | Turn on the bot's own live captions |
| `TEAMS_OPEN_PEOPLE_PANE` | `true` | Open People pane (roster scrape) |
| `TEAMS_FORCE_GALLERY_VIEW` | `true` | Force Gallery view (outline fallback) |
| `TEAMS_CAPTION_LANGUAGE` | `''` | Spoken-language hint, e.g. `en-us`, `es-es` |
| `TEAMS_CAPTION_CONTAINER_SEL` | `[data-tid="closed-caption-v2-window-wrapper"], [data-tid="closed-caption-renderer-wrapper"], …` | Caption container |
| `TEAMS_CAPTION_LINE_SEL` | `.fui-ChatMessageCompact, [data-tid="closed-caption-message"]` | One caption line |
| `TEAMS_CAPTION_AUTHOR_SEL` | `[data-tid="author"]` | Speaker label within a line |
| `TEAMS_CAPTION_TEXT_SEL` | `[data-tid="closed-caption-text"]` | Caption text within a line |
| `TEAMS_SPEAKER_INDICATOR` | `[data-tid="voice-level-stream-outline"]` | Voice-level ring (fallback) |
| `TEAMS_TILE_WRAPPER` | `[data-tid^="calling-participant-stream"], [data-stream-type][data-tid]` | Tile → name |
| `CAPTION_FINALIZE_QUIET_MS` | `1200` | Text-stable window before a caption line is "final" |
| `CAPTION_TRANSCRIPT_MAX_BYTES` | `716800` | Above this, ship the timeline only (no caption text) |

---

## Failure modes & maintenance

| Symptom | Cause | What happens / what to do |
|---|---|---|
| `readinessReport.captionStatus = 'policy-blocked'` | Organizer tenant set `LiveCaptionsEnabledType = Disabled` | Outline + roster fallback still run; API falls back to time-overlap + roster fusion. Nothing to fix. |
| `captionsEnabled: true` but `captionEventCount = 0` | Caption selectors stale (Teams DOM changed) | Report ships `captionStatus: 'sparse'`. **Fix by env override** (`TEAMS_CAPTION_*_SEL`) — no redeploy. Verify new selectors in DevTools against the exact Chrome UA the bot launches. |
| Readiness `notes` full of "not found" | More-menu / submenu labels changed | Update the candidate lists in `meetingReadiness.ts`. Screenshot is uploaded via `uploadDebugImage` on caption-enable failure. |
| Teams meeting UI lives in an iframe for some tenant | `injectedCapture` runs in the main frame | Would miss captions there. Not observed today (join UI is main-frame). If it happens, scope the observer to the iframe document. |
| Wrong caption language | Non-English meeting, no hint | Set `TEAMS_CAPTION_LANGUAGE`, ideally per-meeting from the project/event locale. |

**Selector stability:** `data-tid` is the most durable Teams selector family but
still not a contract. Budget one selector break per quarter. The
`readinessReport` on every meeting makes a fleet-wide break observable (caption
event rate → 0 across meetings).

### Locking in the real selectors from a live run

The shipped selector defaults are a best-guess reconstruction. On the first real
meetings, both modules dump what the DOM actually exposes:

- `runTeamsReadiness` logs `readiness DIAG [toolbar]` (in-call buttons) and
  `readiness DIAG [more-menu-open]` (menu items) — grep the bot logs for
  `readiness DIAG`.
- `startTeamsSpeakerCapture` logs `speakerCapture DIAG` at ~10s / ~40s / then
  every 2 min while `captionsSent === 0`: how many nodes each selector matches
  (`containerHits`, `lineHits`, `outlineHits`, `rosterHits`) plus a sample of
  `data-tid`/class values containing "caption" / "roster".

Take the real `data-tid` / `aria-label` values from those lines and set the
matching `TEAMS_*_SEL` / caption env vars (no redeploy). Set
`TEAMS_READINESS_DIAG=false` once locked in.

**Anonymous "light" meetings.** Guest joins land on
`teams.microsoft.com/light-meetings/launch?...&lightExperience=true` — a reduced
UI. Live captions and the voice-level outline may be absent there entirely (not
just differently named). If the DIAG output confirms nothing caption-like or
outline-like exists, the fallback is roster + diarization + the API's known
attendee list only; the next lever is a `RTCPeerConnection.getStats()` /
`getSynchronizationSources()` audio-level hook (per-SSRC "who is speaking",
mapped to names via roster order) — not yet implemented. Track this in the
API-side roadmap.

---

## Change history

### 2026-09-01 — initial speaker-attribution rework
- **Added** `src/lib/meetingReadiness.ts` — bot enables its own captions, opens
  People pane, forces Gallery view; returns `MeetingReadinessReport`.
- **Added** `src/lib/speakerCapture.ts` — in-page caption + voice-level-outline
  observers, Node-side `SpeakerSpan` / `CaptionCue` accumulator.
- **Added** `SpeakerSpan`, `CaptionCue`, `MeetingReadinessReport`,
  `SpeakerSignalSource`, `MeetingProviderName` to `src/types.ts`.
- **Added** caption/readiness config keys to `src/config.ts`; **fixed** the
  `teamsSpeakerIndicator` default (removed the non-existent
  `[data-is-speaking="true"]`) and widened `teamsTileWrapper`.
- **Added** `setSpeakerTimeline` / `setCaptionTranscript` / `setReadinessReport`
  to `IUploader` + `DiskUploader`; the payload built in
  `uploadRecordingToRemoteStorage` now carries `metadata.speakerTimeline`,
  `metadata.captionTranscript`, `metadata.readinessReport`.
- **Removed** the Node-side `pollActiveSpeaker` loop in `MicrosoftTeamsBot`
  (`this.page.frames()` polling every 1s, wrote `<botId>_speakers.json` to a
  container-local dir that never reached the API).
- **Scope:** Microsoft Teams only. Google Meet / Zoom still record audio-only
  with no speaker capture — see the roadmap in the API-side doc.

### 2026-09-01 (f) — caption enable uses the toggle's own state
Live-DOM check: `#closed-captions-button` carries `data-tid="closed-captions-button-on|off"`,
`aria-checked`, and text "Show/Hide live captions" — a reliable state signal,
readable while the Language-and-speech submenu is open. The old check waited for
the caption *container* to be visible, but that stays empty until the first
spoken word, so in quiet test meetings it flip-flopped and the toggle got
clicked an odd number of times by luck (ending ON) — an even number would have
left captions OFF.
- **`enableLiveCaptions`** now: open submenu → read `captionToggleState()` →
  if `on` leave it alone, if `off` click **once** and poll the state until `on`,
  if `missing` check policy text. No blind retry-toggle.
- Removed `captionsAppearActive()` (container-visibility check).
- Verified against the live client: toggle flips `-off`→`-on` on one click;
  container `[data-tid="closed-caption-v2-window-wrapper"]` is present; roster
  names read from `[data-tid="participant-info-nametag"]`.

### 2026-09-01 (e) — the regression: video gallery has role="menu"
The `menuIsOpen()` guard added in (b) to stop double-click-toggle matched the
Teams **video gallery** (`[data-tid="only-videos-wrapper"]`, which carries
`role="menu"` for roving-tabindex) — so `openMoreMenu` always thought a menu
was already open and **never clicked the More button**. Every readiness run
between (b) and here silently failed to open More → "menu item not found ×3".
- **Fixed**: `realMenuIsOpen()` now excludes `only-videos-wrapper` / roster /
  participant containers and requires actual menu items; it also accepts the
  confirmed More-menu container `[data-tid="callingButtons-showMoreBtn-menu"]`.
  `openMoreMenu` verifies a real dropdown appeared and retries once.
- Validated against the live full client: `realMenuIsOpen` = false with only
  the gallery present; after clicking More it opens
  `[data-tid="callingButtons-showMoreBtn-menu"]` containing
  `#LanguageSpeechMenuControl-id` + `#RecordingMenuControl-id`.
- Whether the anonymous **light** client's More menu also carries "Language
  and speech" is still unconfirmed (could not get a clean anonymous session in
  the inspector) — the widened DIAG on the next real bot run will show it now
  that More actually opens.
- Also confirmed (again) via DB: the recent test runs' empty transcripts are
  Azure returning zero phrases for near-silent recordings, not a code fault.

### 2026-09-01 (d) — light/anonymous client uses different toolbar ids
A real bot run against an anonymous guest join (`light-meetings/launch`)
showed the toolbar has **different button ids than the signed-in full client**
inspected in (c): `custom-view-button` (not `view-mode-button`), `mic-button`
(not `microphone-button`), `screenshare-button` (not `share-button`). Also,
`View` is a **direct toolbar button** here, not nested under More.
- **Fixed**: `forceGalleryView` now tries a direct View click
  (`#custom-view-button`, `#view-mode-button`) before falling back to More →
  View submenu.
- Opening More in this run surfaced only the 3 participant entries as
  `role="menuitem"` — no Language/Settings/Record items were visible to the
  old role-scoped dump. Unknown yet whether they're absent in light mode or
  just not marked `role="menuitem"` there.
- **Widened `dumpUi`**: it now dumps every visible element inside any open
  popup/menu/dialog-like container regardless of ARIA role, plus a page-wide
  keyword search for "caption" / "language and speech" / "transcri" / "record
  and transcribe" — independent of DOM structure. Next real run's
  `readiness DIAG [more-menu-open]` should be conclusive.
- **Confirmed via direct DB query that the pipeline itself is healthy**: this
  test's 0-transcript-line result was Azure Speech genuinely returning zero
  `recognizedPhrases` for a low/no-speech 152s recording — the two immediately
  prior bot runs got 9 and 10 real transcript lines with substantive AI
  summaries. Speaker fusion correctly emitted nothing when given nothing.

### 2026-09-01 (c) — selectors confirmed against a live meeting
Inspected a real Teams meeting DOM directly (joined via the same "Continue on
this browser" → "Continue without audio or video" guest flow the bot uses).
Confirmed and locked in:
- Caption line/author/text selectors were already correct:
  `.fui-ChatMessageCompact` → `span[data-tid="author"]` (e.g. "Michael"),
  `span[data-tid="closed-caption-text"]`, inside
  `[data-tid="closed-caption-v2-window-wrapper"]`.
- The actual menu path: More (`#callingButtons-showMoreBtn`) →
  **`#LanguageSpeechMenuControl-id`** ("Language and speech") →
  **`#closed-captions-button`** (toggles "Show live captions" / "Hide live
  captions" — not a fixed "Turn on" label). Added these exact ids as the
  primary selectors in `meetingReadiness.ts`.
- Gallery path: More → `#view-mode-button` → **`#custom-view-button-MixedGridButton`**.
- Roster rows: **`[data-tid^="participantsInCall-"]`**, itself the row (no
  nested `[role="treeitem"]` needed); `aria-label` = `"Name, Role,
  Muted/Unmuted"` for signed-in/known participants, `"Name Unverified, muted"`
  (no comma before the qualifier) for guests — `cleanRosterName` now strips
  both forms.
- Tile root (present even with camera off — confirms the fallback selector):
  `[data-stream-type="Video"][data-tid]`, `aria-label` starts with the name.
  **`data-tid` on this element is often the participant's email or
  `unknown_email_address@invalid.teams.ms` for guests — never trust it as a
  display name over `aria-label`.** `nameFromTile` now prefers `aria-label`.
  Anonymous joiners are the one case where `data-tid` *is* the typed name
  (e.g. `"Michael"`).
- **Confirmed the camera-off caveat**: `voice-level-stream-outline` render
  count was 0 for an all-camera-off meeting — the outline genuinely does not
  exist without a video tile. Captions are the load-bearing signal for typical
  audio-only meetings; the outline is a real fallback only when someone has
  their camera on.
- Root cause of the "menu item not found ×3" failure in the first live bot run
  was almost certainly the More-menu re-click-toggles-shut bug fixed the same
  day (see below), not wrong selectors — the text-based fallbacks would very
  likely have matched. Exact ids are now primary regardless, for speed and
  i18n-immunity.

### 2026-09-01 (b) — diagnostics after first live run
- First run on an anonymous "light" Teams meeting: readiness could not find the
  More-menu caption/gallery items (`captionStatus: error`), `captionEventCount:
  0`, `rosterNames: []`. Pipeline plumbing confirmed working (roster seeded,
  `maxSpeakers` applied, webhook → row → fusion).
- **Added** `readiness DIAG` (toolbar + open-menu dump) in `meetingReadiness.ts`
  and `speakerCapture DIAG` (selector hit counts + caption/roster-ish samples)
  in `speakerCapture.ts`, gated by `captionsSent === 0` and
  `TEAMS_READINESS_DIAG`.
- **Changed** `openMoreMenu` — detects an already-open `[role="menu"]` instead of
  re-clicking (which toggled it shut); added localised More-button labels.
- **Changed** roster scrape — broader selectors, name cleaning
  (`", Muted, Organizer"` / `"(Guest)"` stripped).
