// A bot sent before a meeting's scheduled start may sit alone (or in the
// lobby) until people arrive. These stretch the "nobody else is here" and
// lobby waits to cover that, so an early bot doesn't leave an empty meeting.
// Without a scheduled start, the normal waits apply.

/** Minutes the bot may stay alone before anyone else joins. */
export function aloneGraceMinutes(
  scheduledStart: string | undefined,
  nowMs: number,
  baseMinutes: number,
  afterStartMinutes: number,
): number {
  const start = scheduledStart ? Date.parse(scheduledStart) : NaN;
  if (!Number.isFinite(start)) return baseMinutes;
  const untilMs = start + afterStartMinutes * 60000 - nowMs;
  return Math.max(baseMinutes, untilMs / 60000);
}

/** Milliseconds to wait in the lobby before giving up. */
export function lobbyWaitMs(
  scheduledStart: string | undefined,
  nowMs: number,
  baseMinutes: number,
  afterStartMinutes: number,
): number {
  return aloneGraceMinutes(scheduledStart, nowMs, baseMinutes, afterStartMinutes) * 60000;
}
