// npm test — the waits for a bot sent before the meeting's scheduled start.
import assert from 'assert';
import { aloneGraceMinutes, lobbyWaitMs } from '../src/lib/earlyJoin';

const now = Date.parse('2030-01-07T15:00:00Z');
const inMin = (m: number) => new Date(now + m * 60000).toISOString();

// No scheduled start: the normal 1 min alone / 10 min lobby.
assert.strictEqual(aloneGraceMinutes(undefined, now, 1, 10), 1);
assert.strictEqual(lobbyWaitMs(undefined, now, 10, 10), 10 * 60000);
assert.strictEqual(aloneGraceMinutes('not a date', now, 1, 10), 1);

// Sent 2 min early: stays until 10 min after the start (12 min from now).
assert.strictEqual(aloneGraceMinutes(inMin(2), now, 1, 10), 12);
assert.strictEqual(lobbyWaitMs(inMin(2), now, 10, 10), 12 * 60000);

// Sent 20 min late: never shorter than the normal waits.
assert.strictEqual(aloneGraceMinutes(inMin(-20), now, 1, 10), 1);
assert.strictEqual(lobbyWaitMs(inMin(-20), now, 10, 10), 10 * 60000);

console.log('earlyJoin: 3 cases passed');
