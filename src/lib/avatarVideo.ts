// avatarVideo.ts
//
// What it is: turns the project's assistant picture into a short looping
//   video file that Chrome plays as the bot's webcam, so the picture shows on
//   the bot's tile in Teams.
// Requires:   ffmpeg on the PATH (installed in every Dockerfile), a writable
//   os.tmpdir().
// Used by:    bots/MicrosoftTeamsBot.ts (before launching the meeting browser)
//   and lib/chromium.ts (passes the file to --use-file-for-fake-video-capture).
// How it works: Teams shows the bot as an anonymous guest, and guests can't
//   have a profile picture. A camera feed is the only way to show one. The
//   picture arrives as a data URI (Virtual PM Settings -> Assistant), is
//   written to a temp file, then ffmpeg letterboxes it onto a 640x360 light
//   background and writes a few identical Y4M frames. Chrome loops the file,
//   so a still frame becomes a steady "camera".

import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import os from 'os';
import path from 'path';

const execFileAsync = promisify(execFile);

/** Video size Teams receives. Small on purpose: a still picture needs no more, and the 1 vCPU container also records. */
const VIDEO_WIDTH = 640;
const VIDEO_HEIGHT = 360;
/** The picture is fitted into a square this tall, centred, like a profile photo. */
const PICTURE_SIZE = 300;
/** Background behind the picture (ADG muted surface, #f5f5f5). */
const BACKGROUND_COLOR = '0xF5F5F5';
/** Frames written and their rate. Chrome loops the file; 5 frames at 5 fps keeps it under 2 MB. */
const FRAME_COUNT = 5;
const FRAMES_PER_SECOND = 5;
/** Largest picture accepted, in bytes of decoded image (the Settings upload is smaller). */
const MAX_PICTURE_BYTES = 5 * 1024 * 1024;
/** ffmpeg should finish in well under a second; give up after this. */
const FFMPEG_TIMEOUT_MS = 20000;
/** Room for ffmpeg's messages (Node's default 1 MB can overflow on odd inputs). */
const FFMPEG_OUTPUT_BYTES = 10 * 1024 * 1024;

const DATA_URI = /^data:image\/(png|jpe?g|gif|webp|bmp);base64,([A-Za-z0-9+/=]+)$/;

/** Logger surface this file uses (the bots' winston logger fits). */
interface IAvatarLogger {
  info: (message: string, meta?: Record<string, unknown>) => void;
  warn: (message: string, meta?: Record<string, unknown>) => void;
}

/**
 * Builds the webcam video for a bot from its picture.
 *
 * @param imageDataUri The assistant picture as `data:image/...;base64,...`, or
 *   undefined / null when the project has none.
 * @param botId Used to name the temp files so parallel bots don't collide.
 * @returns The path of the .y4m file, or null when there is no usable picture
 *   or the conversion failed. Never throws: a missing picture must not stop
 *   the bot from joining, it just joins with its camera off as before.
 */
export async function makeAvatarVideo(
  imageDataUri: string | null | undefined,
  botId: string,
  logger: IAvatarLogger,
): Promise<string | null> {
  // Step 1 of 3: accept only an image data URI of a sane size.
  if (!imageDataUri) return null;
  const match = DATA_URI.exec(imageDataUri.trim());
  if (!match) {
    logger.warn('avatar: picture is not an image data URI; camera stays off');
    return null;
  }
  const pictureBytes = Buffer.from(match[2], 'base64');
  if (pictureBytes.length === 0 || pictureBytes.length > MAX_PICTURE_BYTES) {
    logger.warn('avatar: picture is empty or too large; camera stays off', { bytes: pictureBytes.length });
    return null;
  }

  // Step 2 of 3: write the picture next to where the video will go.
  const safeId = botId.replace(/[^A-Za-z0-9_-]/g, '_') || 'bot';
  const extension = match[1] === 'jpeg' ? 'jpg' : match[1];
  const picturePath = path.join(os.tmpdir(), `avatar-${safeId}.${extension}`);
  const videoPath = path.join(os.tmpdir(), `avatar-${safeId}.y4m`);

  try {
    await fs.writeFile(picturePath, pictureBytes);

    // Step 3 of 3: letterbox the picture and write the looping frames.
    const filter =
      `scale=${PICTURE_SIZE}:${PICTURE_SIZE}:force_original_aspect_ratio=decrease,` +
      `pad=${VIDEO_WIDTH}:${VIDEO_HEIGHT}:(ow-iw)/2:(oh-ih)/2:color=${BACKGROUND_COLOR},` +
      'format=yuv420p';
    await execFileAsync(
      'ffmpeg',
      ['-y', '-nostats', '-loglevel', 'error', '-loop', '1', '-i', picturePath, '-vf', filter,
        '-r', String(FRAMES_PER_SECOND), '-frames:v', String(FRAME_COUNT), videoPath],
      { timeout: FFMPEG_TIMEOUT_MS, maxBuffer: FFMPEG_OUTPUT_BYTES },
    );
    logger.info('avatar: camera video ready', { videoPath });
    return videoPath;
  } catch (error) {
    logger.warn('avatar: could not build the camera video; camera stays off', {
      error: error instanceof Error ? error.message : String(error),
    });
    return null;
  } finally {
    await fs.rm(picturePath, { force: true }).catch(() => undefined);
  }
}

/** Removes a bot's camera video once the meeting is over. Never throws. */
export async function removeAvatarVideo(videoPath: string | null): Promise<void> {
  if (!videoPath) return;
  await fs.rm(videoPath, { force: true }).catch(() => undefined);
}
