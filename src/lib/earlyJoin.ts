// earlyJoin.ts
//
// What it is: how long a bot waits for people when it was sent before the
//   meeting's scheduled start: the "nobody else is here" wait and the lobby wait.
// Requires:   nothing.
// Used by:    bots/MicrosoftTeamsBot.ts (lobby wait, alone and silence checks),
//   test/earlyJoin.test.ts. Setting: ALONE_WAIT_AFTER_START_MINUTES (config.ts).
// How it works: vpm-api sends the bot a few minutes early along with the
//   scheduled start. Both waits are stretched to reach a while past that
//   start, so an early bot doesn't leave an empty meeting before anyone
//   arrives. Without a scheduled start, or when it already passed, the normal
//   waits apply; they are never made shorter.

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
