import { Page } from 'playwright';
import { Logger } from 'winston';
import { MeetingReadinessReport } from '../types';
import config from '../config';
import { uploadDebugImage } from '../services/bugService';

/**
 * Prepare a Microsoft Teams meeting for accurate speaker attribution.
 *
 * See docs/SPEAKER_ATTRIBUTION.md. Everything here is a per-participant or
 * view-only action — it changes nothing for other attendees, fires no consent
 * prompt, and needs no host / organizer action:
 *
 *   - turn on the bot's OWN live captions (+ set the spoken language)
 *   - open the People pane so the roster subtree exists to scrape
 *   - force Gallery view so voice-level rings / tiles render for the fallback
 *
 * Every step is best-effort. A failure is recorded in the report and never
 * aborts the recording. The report ships with the recording so the API knows
 * which fusion strategy to trust and so a fleet-wide DOM regression is visible.
 */

interface ReadinessContext {
  page: Page;
  logger: Logger;
  userId: string;
  botId?: string;
}

const CLICK_TIMEOUT = 2500;

export async function runTeamsReadiness(
  ctx: ReadinessContext,
): Promise<MeetingReadinessReport> {
  const { page, logger } = ctx;
  const report: MeetingReadinessReport = {
    provider: 'microsoft',
    captionsRequested: config.teamsEnableCaptions,
    captionsEnabled: false,
    captionStatus: 'unavailable',
    captionEventCount: 0,
    rosterCaptured: false,
    rosterNames: [],
    peoplePaneOpened: false,
    galleryViewForced: false,
    participantCountSeen: null,
    outlineEventCount: 0,
    notes: [],
  };

  if (!config.teamsReadinessEnabled) {
    report.notes.push('readiness disabled by config');
    return report;
  }

  // One-time snapshot of the in-call toolbar so we can see the real selectors
  // for whatever Teams web variant this meeting is (anonymous "light" vs full).
  await dumpUi(ctx, 'toolbar');

  if (config.teamsOpenPeoplePane) {
    report.peoplePaneOpened = await openPeoplePane(ctx, report);
  }
  if (config.teamsForceGalleryView) {
    report.galleryViewForced = await forceGalleryView(ctx, report);
  }
  if (config.teamsEnableCaptions) {
    const enabled = await enableLiveCaptions(ctx, report);
    report.captionsEnabled = enabled;
    report.captionStatus = enabled ? 'ok' : report.captionStatus;
    if (!enabled) {
      await screenshot(ctx, 'readiness-captions-failed');
    }
  }

  logger.info('Teams meeting readiness complete', {
    captionsEnabled: report.captionsEnabled,
    captionStatus: report.captionStatus,
    peoplePaneOpened: report.peoplePaneOpened,
    galleryViewForced: report.galleryViewForced,
    notes: report.notes,
  });
  return report;
}

/**
 * Log the on-screen controls so a single real run tells us the exact selectors
 * for this Teams web variant. Set TEAMS_READINESS_DIAG=false to silence it once
 * selectors are locked in.
 */
async function dumpUi(ctx: ReadinessContext, phase: string): Promise<void> {
  if (process.env.TEAMS_READINESS_DIAG === 'false') return;
  try {
    const snap = await ctx.page.evaluate(() => {
      const vis = (el: Element) => {
        const r = (el as HTMLElement).getBoundingClientRect();
        return r.width > 0 && r.height > 0;
      };
      const desc = (el: Element) => ({
        tag: el.tagName.toLowerCase(),
        id: (el as HTMLElement).id || undefined,
        dataTid: el.getAttribute('data-tid') || undefined,
        ariaLabel: el.getAttribute('aria-label') || undefined,
        role: el.getAttribute('role') || undefined,
        text: (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 60) || undefined,
      });
      const buttons = Array.from(document.querySelectorAll('button, [role="button"]'))
        .filter(vis)
        .map(desc)
        .filter((d) => d.ariaLabel || d.dataTid || d.text)
        .slice(0, 60);
      const menuItems = Array.from(
        document.querySelectorAll(
          '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"], [role="option"]',
        ),
      )
        .filter(vis)
        .map(desc)
        .slice(0, 40);
      const captionish = Array.from(
        document.querySelectorAll(
          '[data-tid*="caption" i], [class*="caption" i], [aria-label*="caption" i]',
        ),
      )
        .map(desc)
        .slice(0, 20);

      // Role-agnostic: dump EVERY visible clickable-looking descendant inside
      // any currently-open popup/menu/dialog container, regardless of its
      // ARIA role — light/anonymous clients don't always mark items
      // role="menuitem". Also flag which container(s) were found.
      const popupSelectors =
        '[role="menu"], [role="dialog"], [role="listbox"], [data-tid*="flyout" i], [data-tid*="popover" i], [class*="popover" i], [class*="flyout" i], [class*="callout" i]';
      const popups = Array.from(document.querySelectorAll(popupSelectors)).filter(vis);
      const popupContents: Array<Record<string, unknown>> = [];
      for (const popup of popups.slice(0, 3)) {
        const items = Array.from(
          popup.querySelectorAll('button, [role], div[tabindex], span[tabindex]'),
        )
          .filter(vis)
          .filter((el) => (el.textContent || '').trim().length > 0 || el.getAttribute('aria-label'))
          .map(desc)
          .slice(0, 40);
        popupContents.push({ container: desc(popup), items });
      }

      // Whole-page text/attribute search for the words we actually care about,
      // independent of structure — proves whether the feature exists at all
      // in this client variant vs. is just reachable through a different path.
      const keywordHits = Array.from(
        document.querySelectorAll('button, [role], div, span'),
      )
        .filter(vis)
        .filter((el) => {
          const t = ((el.textContent || '') + ' ' + (el.getAttribute('aria-label') || '')).toLowerCase();
          return (
            /\bcaption/.test(t) ||
            /\blanguage and speech\b/.test(t) ||
            /\btranscri/.test(t) ||
            /\brecord and transcribe\b/.test(t)
          );
        })
        .map(desc)
        .slice(0, 20);

      return { buttons, menuItems, captionish, popupCount: popups.length, popupContents, keywordHits };
    });
    ctx.logger.info(`readiness DIAG [${phase}]`, snap);
  } catch (err) {
    ctx.logger.warn(`readiness DIAG [${phase}] failed`, {
      error: err instanceof Error ? err.message : String(err),
    });
  }
}

/** Click the first candidate that becomes visible; returns what worked. */
async function clickFirst(
  page: Page,
  candidates: string[],
  label: string,
  logger: Logger,
): Promise<string | null> {
  for (const selector of candidates) {
    try {
      const loc = page.locator(selector).first();
      if (await loc.isVisible({ timeout: CLICK_TIMEOUT }).catch(() => false)) {
        await loc.click({ timeout: CLICK_TIMEOUT });
        logger.info(`readiness: clicked ${label}`, { selector });
        return selector;
      }
    } catch (err) {
      logger.debug?.(`readiness: ${label} candidate failed`, {
        selector,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }
  return null;
}

/**
 * Is a real dropdown/flyout menu open? Excludes the video gallery
 * (`only-videos-wrapper`) and the roster tree, which also carry role="menu"
 * (roving tabindex) and are always present during a call — the false positive
 * that silently stopped More from ever opening.
 */
async function realMenuIsOpen(page: Page): Promise<boolean> {
  return page
    .locator(
      // The More menu container has a stable data-tid (confirmed via live DOM).
      '[data-tid="callingButtons-showMoreBtn-menu"]:visible, ' +
        '[data-tid$="-menu"][role="menu"]:visible, ' +
        // ...or any real dropdown: a role="menu" that is NOT the video gallery
        // / roster tree and contains actual menu items.
        '[role="menu"]:visible:not([data-tid="only-videos-wrapper"]):not([data-tid*="roster" i]):not([data-tid*="participant" i]) [role="menuitem"], ' +
        '[role="menu"]:visible [role="menuitemcheckbox"], ' +
        '[role="menu"]:visible [role="menuitemradio"]',
    )
    .first()
    .isVisible({ timeout: 300 })
    .catch(() => false);
}

async function openMoreMenu(page: Page, logger: Logger): Promise<boolean> {
  if (await realMenuIsOpen(page)) return true; // already open — don't toggle shut
  const hit = await clickFirst(
    page,
    [
      '#callingButtons-showMoreBtn',
      'button[data-tid="callingButtons-showMoreBtn"]',
      'button[data-tid="more-button"]',
      '#callControlsMoreBtn',
      'button[aria-label="More"]',
      'button[aria-label*="More actions" i]',
      'button[aria-label*="More" i]',
      'button[aria-label*="Weitere" i]',
      'button[aria-label*="Más" i]',
    ],
    'More menu',
    logger,
  );
  if (!hit) return false;
  await page.waitForTimeout(700);
  const opened = await realMenuIsOpen(page);
  if (!opened) {
    // Click landed but no dropdown — try once more (a stray menu may have
    // eaten the first click).
    await clickFirst(page, ['#callingButtons-showMoreBtn', 'button[aria-label="More"]'], 'More menu (retry)', logger);
    await page.waitForTimeout(700);
    return realMenuIsOpen(page);
  }
  return true;
}

/**
 * Read the caption toggle's own state. Confirmed via live-DOM inspection:
 * `#closed-captions-button` carries `data-tid="closed-captions-button-on|off"`,
 * `aria-checked`, and text "Show/Hide live captions". Only readable while the
 * Language-and-speech submenu is open. This is the reliable signal — the caption
 * *container* is empty (and often not "visible") until the first spoken word,
 * which is why the old container check flip-flopped and the toggle got clicked
 * an odd number of times by luck.
 */
async function captionToggleState(
  page: Page,
): Promise<'on' | 'off' | 'missing'> {
  const on = page
    .locator(
      '#closed-captions-button[data-tid$="-on"], ' +
        '#closed-captions-button[aria-checked="true"], ' +
        '[role="menuitemcheckbox"][aria-label*="Hide live captions" i], ' +
        '[role="menuitemcheckbox"]:has-text("Hide live captions")',
    )
    .first();
  if (await on.isVisible({ timeout: 500 }).catch(() => false)) return 'on';

  const off = page
    .locator(
      '#closed-captions-button[data-tid$="-off"], ' +
        '#closed-captions-button[aria-checked="false"], ' +
        '[role="menuitemcheckbox"][aria-label*="Show live captions" i], ' +
        '[role="menuitemcheckbox"]:has-text("Show live captions")',
    )
    .first();
  if (await off.isVisible({ timeout: 500 }).catch(() => false)) return 'off';

  return 'missing';
}

async function enableLiveCaptions(
  ctx: ReadinessContext,
  report: MeetingReadinessReport,
): Promise<boolean> {
  const { page, logger } = ctx;

  for (let attempt = 1; attempt <= 3; attempt++) {
    const opened = await openMoreMenu(page, logger);
    if (!opened) {
      report.notes.push(`captions: More menu not found (attempt ${attempt})`);
      await page.waitForTimeout(1500);
      continue;
    }
    if (attempt === 1) await dumpUi(ctx, 'more-menu-open');

    // "Language and speech" submenu. #LanguageSpeechMenuControl-id confirmed via
    // live-DOM inspection (see docs/SPEAKER_ATTRIBUTION.md); text fallbacks cover
    // localisation / future id changes.
    await clickFirst(
      page,
      [
        '#LanguageSpeechMenuControl-id',
        '[role="menuitem"][aria-label*="Language and speech" i]',
        '[role="menuitem"]:has-text("Language and speech")',
        '[role="menuitem"]:has-text("Language")',
      ],
      'Language and speech submenu',
      logger,
    );
    await page.waitForTimeout(400);

    let state = await captionToggleState(page);
    logger.info('readiness: caption toggle state', { attempt, state });

    if (state === 'on') {
      // Already enabled — do NOT click (that would turn it off).
      if (config.teamsCaptionLanguage) {
        await setCaptionLanguage(ctx).catch(() => undefined);
      }
      await page.keyboard.press('Escape').catch(() => undefined);
      return true;
    }

    if (state === 'missing') {
      const bodyText =
        (await page.evaluate(() => document.body.innerText).catch(() => '')) ||
        '';
      if (/caption.*(turned off|disabled|not allowed|isn't available)/i.test(bodyText)) {
        report.captionStatus = 'policy-blocked';
        report.notes.push(
          'captions: appear disabled by the organiser tenant policy',
        );
        return false;
      }
      report.notes.push(`captions: toggle not found (attempt ${attempt})`);
      await page.keyboard.press('Escape').catch(() => undefined);
      await page.waitForTimeout(1200);
      continue;
    }

    // state === 'off' — click exactly once, then confirm via the toggle state.
    await clickFirst(
      page,
      [
        '#closed-captions-button',
        '[data-tid^="closed-captions-button"]',
        '[role="menuitemcheckbox"][aria-label*="Show live captions" i]',
        '[role="menuitemcheckbox"]:has-text("Show live captions")',
      ],
      'live captions toggle',
      logger,
    );

    for (let i = 0; i < 6 && state !== 'on'; i++) {
      await page.waitForTimeout(400);
      state = await captionToggleState(page);
    }

    if (state === 'on') {
      if (config.teamsCaptionLanguage) {
        await setCaptionLanguage(ctx).catch(() => undefined);
      }
      await page.keyboard.press('Escape').catch(() => undefined);
      return true;
    }

    report.notes.push(
      `captions: clicked toggle but state stayed "${state}" (attempt ${attempt})`,
    );
    await page.keyboard.press('Escape').catch(() => undefined);
    await page.waitForTimeout(1200);
  }

  report.captionStatus =
    report.captionStatus === 'policy-blocked' ? 'policy-blocked' : 'error';
  return false;
}

async function setCaptionLanguage(ctx: ReadinessContext): Promise<void> {
  const { page, logger } = ctx;
  const lang = config.teamsCaptionLanguage;
  await openMoreMenu(page, logger);
  await clickFirst(
    page,
    [
      '#LanguageSpeechMenuControl-id',
      '[role="menuitem"][aria-label*="Language and speech" i]',
      '[role="menuitem"]:has-text("Language and speech")',
    ],
    'Language and speech (for language)',
    logger,
  );
  await clickFirst(
    page,
    [
      '[role="menuitem"][aria-label*="spoken language" i]',
      '[role="menuitem"]:has-text("spoken language")',
      '[role="menuitem"]:has-text("Caption language")',
    ],
    'spoken language',
    logger,
  );
  await page.waitForTimeout(300);
  await clickFirst(
    page,
    [`[role="option"][aria-label*="${lang}" i]`, `[role="menuitemradio"][aria-label*="${lang}" i]`],
    `caption language ${lang}`,
    logger,
  );
  await page.keyboard.press('Escape').catch(() => undefined);
}

async function openPeoplePane(
  ctx: ReadinessContext,
  report: MeetingReadinessReport,
): Promise<boolean> {
  const { page, logger } = ctx;
  const hit = await clickFirst(
    page,
    [
      'button[data-tid="call-roster-button"]',
      'button[data-tid="roster-button"]',
      '#roster-button',
      'button[aria-label="People"]',
      'button[aria-label*="Participants" i]',
      'button[aria-label*="People" i]',
    ],
    'People pane',
    logger,
  );
  if (!hit) {
    report.notes.push('roster: People button not found');
    return false;
  }
  await page.waitForTimeout(800);
  return true;
}

async function forceGalleryView(
  ctx: ReadinessContext,
  report: MeetingReadinessReport,
): Promise<boolean> {
  const { page, logger } = ctx;
  // The View control is a DIRECT toolbar button on some clients (anonymous
  // "light" experience: id="custom-view-button") and only reachable via More
  // on others (id="view-mode-button"). Try direct first — cheaper and doesn't
  // depend on More's contents, which differ by client variant.
  const directView = await clickFirst(
    page,
    ['#custom-view-button', '#view-mode-button', 'button[aria-label="View"]'],
    'View (direct toolbar)',
    logger,
  );
  if (!directView) {
    const opened = await openMoreMenu(page, logger);
    if (!opened) {
      report.notes.push('gallery: neither a direct View button nor More menu found');
      return false;
    }
    await clickFirst(
      page,
      [
        '#custom-view-button',
        '#view-mode-button',
        '[role="menuitem"][aria-label*="View" i]',
        '[role="menuitem"]:has-text("View")',
        '[role="menuitem"]:has-text("Layout")',
      ],
      'View submenu',
      logger,
    );
  }
  await page.waitForTimeout(300);
  // #custom-view-button-MixedGridButton = "Gallery" confirmed via live-DOM
  // inspection. Avoid a bare :has-text("Gallery") first (also matches "Gallery
  // size").
  const hit = await clickFirst(
    page,
    [
      '#custom-view-button-MixedGridButton',
      '[role="menuitemcheckbox"][aria-label*="Gallery" i]:not([aria-label*="size" i])',
      '[role="menuitem"]:has-text("Gallery"):not(:has-text("size"))',
    ],
    'Gallery view',
    logger,
  );
  await page.keyboard.press('Escape').catch(() => undefined);
  if (!hit) {
    report.notes.push('gallery: Gallery option not found');
    return false;
  }
  return true;
}

async function screenshot(ctx: ReadinessContext, label: string): Promise<void> {
  try {
    const buf = await ctx.page.screenshot({ type: 'png', fullPage: false });
    await uploadDebugImage(buf, label, ctx.userId, ctx.logger, ctx.botId);
  } catch (err) {
    ctx.logger.warn('readiness: debug screenshot failed', {
      error: err instanceof Error ? err.message : String(err),
    });
  }
}
