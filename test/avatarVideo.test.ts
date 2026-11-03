// avatarVideo.test.ts
//
// What it is: checks that the assistant picture becomes a Chrome webcam file.
// Requires:   ffmpeg on the PATH, src/lib/avatarVideo.ts.
// Used by:    npm test (package.json).
// How it works: makes a real picture with ffmpeg, converts it, and checks the
//   Y4M header and size; then checks that bad input returns null instead of
//   throwing or hanging.

import assert from 'assert';
import fs from 'fs';
import { execFileSync } from 'child_process';
import os from 'os';
import path from 'path';
import { makeAvatarVideo, removeAvatarVideo } from '../src/lib/avatarVideo';

/** A broken PNG (bad compressed data): must give null, not hang or throw. */
const BROKEN_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGO4tmYNAATkAj9SpRdxAAAAAElFTkSuQmCC';

/** A real 64x64 teal PNG made with ffmpeg, as a data URI. */
function tealPictureDataUri(): string {
  const file = path.join(os.tmpdir(), 'avatar-test-input.png');
  execFileSync('ffmpeg', ['-y', '-loglevel', 'error', '-f', 'lavfi', '-i', 'color=c=0x38adc2:s=64x64', '-frames:v', '1', file]);
  const dataUri = `data:image/png;base64,${fs.readFileSync(file).toString('base64')}`;
  fs.rmSync(file, { force: true });
  return dataUri;
}

/** Largest acceptable video: 5 frames of 640x360 YUV 4:2:0 plus headers. */
const MAX_VIDEO_BYTES = 2 * 1024 * 1024;

const silentLogger = { info: (): void => undefined, warn: (): void => undefined };

async function main(): Promise<void> {
  // A real picture becomes a Y4M file Chrome can loop.
  const videoPath = await makeAvatarVideo(tealPictureDataUri(), 'test-bot', silentLogger);
  assert.ok(videoPath, 'video created');
  const header = fs.readFileSync(videoPath).subarray(0, 9).toString('ascii');
  assert.strictEqual(header, 'YUV4MPEG2', 'Y4M file');
  assert.ok(fs.statSync(videoPath).size < MAX_VIDEO_BYTES, 'small enough');
  await removeAvatarVideo(videoPath);
  assert.ok(!fs.existsSync(videoPath), 'cleaned up');

  // No picture, or not an image: null, the bot joins with its camera off.
  assert.strictEqual(await makeAvatarVideo(undefined, 'x', silentLogger), null);
  assert.strictEqual(await makeAvatarVideo('https://example.com/a.png', 'x', silentLogger), null);
  assert.strictEqual(await makeAvatarVideo('data:image/png;base64,', 'x', silentLogger), null);
  assert.strictEqual(await makeAvatarVideo(BROKEN_PNG, 'x', silentLogger), null, 'broken picture: camera off');

  console.log('avatarVideo: 5 cases passed');
}

main().catch((error: unknown) => {
  console.error(error);
  process.exit(1);
});
