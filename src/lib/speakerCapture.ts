import { Page } from 'playwright';
import { Logger } from 'winston';
import { CaptionCue, SpeakerSpan } from '../types';

/**
 * In-page "who is speaking" capture for Microsoft Teams (web client).
 *
 * See docs/SPEAKER_ATTRIBUTION.md. Two independent, cheap signals run in the
 * page and stream events back to Node:
 *
 *   1. Live captions  -> `{ kind: 'caption', speaker, text, tSec }`
 *      The conferencing server's own per-stream attribution. Best signal.
 *      One MutationObserver (flag only) + a 500ms parse tick. Teams keeps just
 *      ~2-3 caption lines in the DOM and mutates a line in place until
 *      recognition firms up, so a line is "final" once its text has been
 *      stable for `quietMs` or it has scrolled out of the DOM.
 *
 *   2. Voice-level outline -> `{ kind: 'outline', name, tSec }`
 *      The active-speaker ring animates its inline `style` while a participant
 *      talks (there is no boolean attribute). We count style mutations per ring
 *      in a sliding window; the ring with the most recent activity is the
 *      active speaker, mapped to a name via its tile.
 *
 *   3. Roster -> `{ kind: 'roster', names }`  (People pane rows, slow cadence)
 *
 * Everything is view-only / read-only and invisible to other attendees.
 */

export interface SpeakerCaptureOptions {
  /** `Date.now()` at the moment recording started; caption/outline timestamps are relative to this. */
  recordingStartedAtEpochMs: number;
  /** Bot's own display name, filtered out of every signal. */
  botDisplayName?: string;
  captionContainerSel: string;
  captionLineSel: string;
  captionAuthorSel: string;
  captionTextSel: string;
  outlineSel: string;
  tileSel: string;
  captionFinalizeQuietMs: number;
}

export interface SpeakerCaptureSnapshot {
  speakerTimeline: SpeakerSpan[];
  captionTranscript: CaptionCue[];
  rosterNames: string[];
  captionEventCount: number;
  outlineEventCount: number;
}

export interface SpeakerCaptureHandle {
  snapshot(): SpeakerCaptureSnapshot;
  stop(): Promise<void>;
}

interface InPageEvent {
  kind: 'caption' | 'outline' | 'roster' | 'diag';
  speaker?: string | null;
  name?: string | null;
  text?: string;
  tSec?: number;
  names?: string[];
  diag?: Record<string, unknown>;
}

const EXPOSED_FN = 'screenAppSpeakerEvent';
// A caption cue with no following cue is assumed to last at most this long.
const MAX_OPEN_SPAN_SEC = 20;

export async function startTeamsSpeakerCapture(
  page: Page,
  logger: Logger,
  opts: SpeakerCaptureOptions,
): Promise<SpeakerCaptureHandle> {
  const captionTranscript: CaptionCue[] = [];
  const outlineSpans: SpeakerSpan[] = [];
  const rosterNames = new Set<string>();
  let captionEventCount = 0;
  let outlineEventCount = 0;

  const onEvent = (evt: InPageEvent) => {
    try {
      if (evt.kind === 'caption') {
        captionEventCount++;
        const tStartSec = typeof evt.tSec === 'number' ? evt.tSec : 0;
        const prev = captionTranscript[captionTranscript.length - 1];
        if (prev && prev.tEndSec === null) {
          prev.tEndSec = Math.max(prev.tStartSec, tStartSec);
        }
        captionTranscript.push({
          speaker: evt.speaker ? String(evt.speaker) : null,
          text: String(evt.text || ''),
          tStartSec,
          tEndSec: null,
        });
      } else if (evt.kind === 'outline' && evt.name) {
        outlineEventCount++;
        const tSec = typeof evt.tSec === 'number' ? evt.tSec : 0;
        const prev = outlineSpans[outlineSpans.length - 1];
        if (prev && prev.tEndSec === null) prev.tEndSec = Math.max(prev.tStartSec, tSec);
        outlineSpans.push({
          tStartSec: tSec,
          tEndSec: null,
          name: String(evt.name),
          source: 'teams-outline',
          confidence: 0.5,
        });
      } else if (evt.kind === 'roster' && Array.isArray(evt.names)) {
        for (const n of evt.names) {
          const clean = String(n || '').trim();
          if (clean) rosterNames.add(clean);
        }
      } else if (evt.kind === 'diag') {
        logger.info('speakerCapture DIAG', evt.diag || {});
      }
    } catch (err) {
      logger.warn('speakerCapture: failed to handle in-page event', {
        error: err instanceof Error ? err.message : String(err),
      });
    }
  };

  try {
    await page.exposeFunction(EXPOSED_FN, onEvent);
  } catch (err) {
    // Ignore "already registered" on retry; anything else is fatal for capture.
    const msg = err instanceof Error ? err.message : String(err);
    if (!/has been already registered|already exists/i.test(msg)) {
      logger.error('speakerCapture: could not expose bridge function', { error: msg });
      throw err;
    }
  }

  try {
    await page.evaluate(injectedCapture, {
      startMs: opts.recordingStartedAtEpochMs,
      botName: opts.botDisplayName || '',
      containerSel: opts.captionContainerSel,
      lineSel: opts.captionLineSel,
      authorSel: opts.captionAuthorSel,
      textSel: opts.captionTextSel,
      outlineSel: opts.outlineSel,
      tileSel: opts.tileSel,
      quietMs: opts.captionFinalizeQuietMs,
      fnName: EXPOSED_FN,
    });
    logger.info('speakerCapture: in-page caption + outline observers installed');
  } catch (err) {
    logger.error('speakerCapture: failed to inject in-page observers', {
      error: err instanceof Error ? err.message : String(err),
    });
  }

  const buildTimeline = (): SpeakerSpan[] => {
    const spans: SpeakerSpan[] = [];
    // Caption-derived spans (high confidence).
    for (let i = 0; i < captionTranscript.length; i++) {
      const cue = captionTranscript[i];
      if (!cue.speaker) continue;
      const next = captionTranscript[i + 1];
      const end =
        cue.tEndSec ??
        (next ? next.tStartSec : cue.tStartSec + MAX_OPEN_SPAN_SEC);
      spans.push({
        tStartSec: cue.tStartSec,
        tEndSec: Math.max(cue.tStartSec, end),
        name: cue.speaker,
        source: 'teams-caption',
        confidence: 0.9,
      });
    }
    // Outline-derived spans (fallback), only where captions are silent.
    for (const s of outlineSpans) {
      spans.push({
        ...s,
        tEndSec: s.tEndSec ?? s.tStartSec + MAX_OPEN_SPAN_SEC,
      });
    }
    return mergeAdjacent(spans.sort((a, b) => a.tStartSec - b.tStartSec));
  };

  return {
    snapshot: () => ({
      speakerTimeline: buildTimeline(),
      captionTranscript: captionTranscript.map((c) => ({
        ...c,
        tEndSec: c.tEndSec ?? c.tStartSec + MAX_OPEN_SPAN_SEC,
      })),
      rosterNames: [...rosterNames],
      captionEventCount,
      outlineEventCount,
    }),
    stop: async () => {
      try {
        await page.evaluate((fnName: string) => {
          (window as unknown as Record<string, unknown>)['__' + fnName + 'Stop'] = true;
        }, EXPOSED_FN);
      } catch {
        /* page may already be gone */
      }
    },
  };
}

/** Collapse consecutive same-name spans and drop zero-length noise. */
function mergeAdjacent(spans: SpeakerSpan[]): SpeakerSpan[] {
  const out: SpeakerSpan[] = [];
  for (const s of spans) {
    if (s.tEndSec !== null && s.tEndSec - s.tStartSec < 0.2) continue;
    const last = out[out.length - 1];
    if (
      last &&
      last.name === s.name &&
      last.source === s.source &&
      s.tStartSec - (last.tEndSec ?? last.tStartSec) < 2
    ) {
      last.tEndSec = s.tEndSec;
      last.confidence = Math.max(last.confidence, s.confidence);
    } else {
      out.push({ ...s });
    }
  }
  return out;
}

/**
 * Runs entirely in the page. Must be self-contained (Playwright serializes it).
 */
function injectedCapture(opts: {
  startMs: number;
  botName: string;
  containerSel: string;
  lineSel: string;
  authorSel: string;
  textSel: string;
  outlineSel: string;
  tileSel: string;
  quietMs: number;
  fnName: string;
}): void {
  const w = window as unknown as Record<string, unknown>;
  const stopKey = '__' + opts.fnName + 'Stop';
  const guardKey = '__' + opts.fnName + 'Running';
  if (w[guardKey]) return;
  w[guardKey] = true;

  const norm = (s: string | null | undefined): string =>
    (s || '').replace(/\s+/g, ' ').trim();
  const now = (): number => (Date.now() - opts.startMs) / 1000;
  const send = (e: Record<string, unknown>): void => {
    const fn = w[opts.fnName] as ((e: unknown) => void) | undefined;
    if (typeof fn === 'function') {
      try {
        fn(e);
      } catch {
        /* bridge closed */
      }
    }
  };
  const selAll = (sel: string, root: Document | Element = document): Element[] => {
    const parts = sel
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);
    const out: Element[] = [];
    for (const p of parts) {
      try {
        root.querySelectorAll(p).forEach((el) => out.push(el));
      } catch {
        /* bad selector, skip */
      }
    }
    return out;
  };
  const firstText = (el: Element, sel: string): string => {
    for (const s of sel.split(',').map((x) => x.trim())) {
      const hit = el.querySelector(s);
      if (hit && norm(hit.textContent)) return norm(hit.textContent);
    }
    return '';
  };
  const isBot = (name: string | null): boolean =>
    !!name &&
    !!opts.botName &&
    name.toLowerCase().includes(opts.botName.toLowerCase());

  // ---- roster (People pane) ------------------------------------------------
  // `[data-tid^="participantsInCall-"]` confirmed via live-DOM inspection: each
  // row is the element itself (no nested [role="treeitem"] needed), aria-label
  // = "Name, Role, Muted/Unmuted" or "Name Unverified, muted" for guests.
  const rosterSelectors = [
    '[data-tid^="participantsInCall-"]',
    '[data-tid="roster-title-section-participant"]',
    '[data-tid="participant-roster"] [role="treeitem"]',
    '[data-tid*="roster" i] [role="treeitem"]',
    '[role="tree"][aria-label] [role="treeitem"][aria-label]',
    '[data-tid="calling-roster-cell-name"]',
  ].join(',');
  const cleanRosterName = (raw: string): string => {
    // "Jane Doe, Muted, Organizer" -> "Jane Doe"
    // "Jane Doe Unverified, muted" (guest, no comma before the qualifier) -> "Jane Doe"
    let n = norm(raw.split(/,|\||•/)[0]);
    n = n.replace(/\s+(unverified|guest|external)$/i, '').trim();
    n = n.replace(/\((guest|external|unverified)\)/gi, '').trim();
    return n;
  };
  const scrapeRoster = (): void => {
    const names = new Set<string>();
    for (const el of selAll(rosterSelectors)) {
      const raw =
        el.getAttribute('aria-label') ||
        el.getAttribute('title') ||
        el.textContent ||
        '';
      const n = cleanRosterName(raw);
      if (n && n.length >= 2 && n.length <= 80 && /[a-z]/i.test(n) && !isBot(n)) {
        names.add(n);
      }
    }
    if (names.size) send({ kind: 'roster', names: [...names] });
  };

  // ---- captions -----------------------------------------------------------
  const lineState = new WeakMap<
    Element,
    { speaker: string | null; text: string; tSec: number; lastChange: number }
  >();
  const finalized = new WeakSet<Element>();
  let tracked: Element[] = [];
  let captionCount = 0;

  const finalizeLine = (
    li: Element,
    s: { speaker: string | null; text: string; tSec: number },
  ): void => {
    if (finalized.has(li)) return;
    finalized.add(li);
    if (!s.text || isBot(s.speaker)) return;
    captionCount++;
    send({ kind: 'caption', speaker: s.speaker, text: s.text, tSec: s.tSec });
  };

  const tickCaptions = (): void => {
    const roots = selAll(opts.containerSel);
    if (!roots.length) return;
    const current: Element[] = [];
    for (const root of roots)
      for (const li of selAll(opts.lineSel, root as Element)) current.push(li);
    const nowMs = Date.now();
    for (const li of current) {
      const speaker = firstText(li, opts.authorSel) || null;
      const text = firstText(li, opts.textSel) || norm(li.textContent);
      if (!text) continue;
      let s = lineState.get(li);
      if (!s) {
        s = { speaker, text, tSec: now(), lastChange: nowMs };
        lineState.set(li, s);
      } else if (s.text !== text || s.speaker !== speaker) {
        s.text = text;
        s.speaker = speaker;
        s.lastChange = nowMs;
      }
      if (nowMs - s.lastChange >= opts.quietMs) finalizeLine(li, s);
    }
    for (const li of tracked) {
      if (current.indexOf(li) === -1) {
        const s = lineState.get(li);
        if (s) finalizeLine(li, s);
      }
    }
    tracked = current;
  };

  let captionObserver: MutationObserver | null = null;
  const attachCaptionObserver = (): void => {
    const c = selAll(opts.containerSel)[0] as (Element & { __saCap?: boolean }) | undefined;
    if (!c || c.__saCap) return;
    c.__saCap = true;
    // childList/subtree only — keeps the observer wired so a container swap is
    // noticed. All caption parsing runs on the 500ms tick regardless, so cost
    // is constant no matter the caption token rate.
    captionObserver = new MutationObserver(() => undefined);
    captionObserver.observe(c, { childList: true, subtree: true });
  };

  // ---- voice-level outline ----------------------------------------------
  const outlineHits = new WeakMap<Element, number[]>();
  let lastOutlineName = '';
  const outlineObserver = new MutationObserver((muts) => {
    const nowMs = Date.now();
    for (const m of muts) {
      if (m.type !== 'attributes') continue;
      const el = m.target as Element;
      let arr = outlineHits.get(el);
      if (!arr) {
        arr = [];
        outlineHits.set(el, arr);
      }
      arr.push(nowMs);
    }
  });
  const attachOutlines = (): void => {
    for (const el of selAll(opts.outlineSel) as (Element & { __saOut?: boolean })[]) {
      if (el.__saOut) continue;
      el.__saOut = true;
      outlineObserver.observe(el, { attributes: true, attributeFilter: ['style', 'class'] });
    }
  };
  const nameFromTile = (node: Element): string => {
    let tile: Element | null = null;
    for (const s of opts.tileSel.split(',').map((x) => x.trim())) {
      tile = node.closest(s);
      if (tile) break;
    }
    if (!tile) return '';
    // Prefer aria-label ("Jane Doe, muted, ...") — confirmed via live-DOM
    // inspection. tile data-tid is often the participant's EMAIL or an
    // "unknown_email_address@invalid.teams.ms" placeholder for guests, not a
    // display name; only fall back to it when it looks like an actual name
    // (anonymous joiners: data-tid is literally their typed name, e.g. "Michael").
    const al = tile.getAttribute('aria-label') || '';
    if (al) {
      const first = norm(al.split(',')[0]).replace(/\s+(unverified|guest|external)$/i, '');
      if (first) return first;
    }
    const dt = tile.getAttribute('data-tid') || '';
    if (dt && !/^calling-|stream|participant|@|unknown_email/i.test(dt) && dt.length <= 80) {
      return norm(dt);
    }
    return '';
  };
  const tickOutline = (): void => {
    attachOutlines();
    const nowMs = Date.now();
    let best: Element | null = null;
    let bestCount = 0;
    for (const el of selAll(opts.outlineSel)) {
      const arr = (outlineHits.get(el) || []).filter((t) => nowMs - t < 1500);
      outlineHits.set(el, arr);
      if (arr.length > bestCount) {
        bestCount = arr.length;
        best = el;
      }
    }
    if (best && bestCount >= 3) {
      const name = nameFromTile(best);
      if (name && name !== lastOutlineName && !isBot(name)) {
        lastOutlineName = name;
        send({ kind: 'outline', name, tSec: now() });
      }
    }
  };

  // ---- diagnostics: while nothing is coming through, report what the DOM
  //      actually has so selectors can be fixed from one real run.
  const probe = (): void => {
    const containerHits = selAll(opts.containerSel).length;
    const lineHits = selAll(opts.lineSel).length;
    const captionish = Array.from(
      document.querySelectorAll(
        '[data-tid*="caption" i], [class*="caption" i], [class*="Caption"]',
      ),
    )
      .slice(0, 15)
      .map((el) => ({
        tag: el.tagName.toLowerCase(),
        dataTid: el.getAttribute('data-tid') || undefined,
        cls: (el.getAttribute('class') || '').slice(0, 60) || undefined,
        text: (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 50) || undefined,
      }));
    const outlineHitsCount = selAll(opts.outlineSel).length;
    const rosterHits = selAll(rosterSelectors).length;
    const rosterish = Array.from(
      document.querySelectorAll('[data-tid*="roster" i], [data-tid*="participant" i]'),
    )
      .slice(0, 12)
      .map((el) => ({
        tag: el.tagName.toLowerCase(),
        dataTid: el.getAttribute('data-tid') || undefined,
        role: el.getAttribute('role') || undefined,
        text: (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 50) || undefined,
      }));
    send({
      kind: 'diag',
      diag: {
        captionsSent: captionCount,
        containerSel: opts.containerSel,
        containerHits,
        lineHits,
        outlineSel: opts.outlineSel,
        outlineHits: outlineHitsCount,
        rosterHits,
        captionish,
        rosterish,
      },
    });
  };

  // ---- master loop ------------------------------------------------------
  let ticks = 0;
  const loop = window.setInterval(() => {
    if (w[stopKey]) {
      window.clearInterval(loop);
      try {
        captionObserver?.disconnect();
        outlineObserver.disconnect();
      } catch {
        /* noop */
      }
      return;
    }
    try {
      attachCaptionObserver();
      tickCaptions();
      tickOutline();
      if (ticks % 30 === 0) scrapeRoster(); // ~every 15s
      // probe at 10s, 40s, then every 2min while still empty
      if (
        captionCount === 0 &&
        (ticks === 20 || ticks === 80 || (ticks > 0 && ticks % 240 === 0))
      ) {
        probe();
      }
    } catch {
      /* keep the loop alive through transient DOM errors */
    }
    ticks++;
  }, 500);

  setTimeout(scrapeRoster, 3000);
}
