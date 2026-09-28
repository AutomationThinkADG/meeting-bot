// Lets an outside caller (the API's "Stop bot" action, called from the
// Timeline Dashboard) end a bot's recording on demand — e.g. it's stuck
// waiting for a lone-participant/silence timeout that will never come
// (the meeting actually ended, or two bots are both waiting on each other),
// or the user just wants to cut it short and retry.
//
// Each recording registers a stop function keyed by botId while it's
// running, and unregisters it when it's done — same lifetime as the
// meetingEnded flag it's tied to. Calling stop() reuses the exact same
// "meetingEnded = true" path the silence timer and the on-page ended-text
// detector already use, so it's the same clean stop-record-then-upload
// flow, not a hard kill.
const activeBots = new Map<string, () => void>();

export function registerActiveBot(botId: string, stop: () => void): void {
  activeBots.set(botId, stop);
}

export function unregisterActiveBot(botId: string): void {
  activeBots.delete(botId);
}

/** Returns true if a running bot with this id was found and told to stop. */
export function stopActiveBot(botId: string): boolean {
  const stop = activeBots.get(botId);
  if (!stop) return false;
  stop();
  return true;
}

export function isActiveBot(botId: string): boolean {
  return activeBots.has(botId);
}
