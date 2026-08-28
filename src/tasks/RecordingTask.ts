import { Page } from 'playwright';
import { Task } from '../lib/Task';
import config from '../config';
import { Logger } from 'winston';
import { getRecordingMimeTypesForExtension } from '../lib/recording';
import fs from 'fs-extra';
import path from 'path';

export class RecordingTask extends Task<null, void> {
  private userId: string;
  private teamId: string;
  private page: Page;
  private duration: number;
  private inactivityLimit: number;
  private slightlySecretId: string;

  constructor(
    userId: string,
    teamId: string,
    page: Page,
    duration: number,
    slightlySecretId: string,
    logger: Logger,
  ) {
    super(logger);
    this.userId = userId;
    this.teamId = teamId;
    this.duration = duration;
    this.inactivityLimit = config.inactivityLimit * 60 * 1000;
    this.page = page;
    this.slightlySecretId = slightlySecretId;
  }

  protected async execute(): Promise<void> {
    const { mimeTypes } = getRecordingMimeTypesForExtension(
      config.uploaderFileExtension,
    );
    const loneParticipantExitDelayMs =
      config.loneParticipantExitDelaySeconds * 1000;

    // --- STEP 1: EXPOSE WEBHOOK FOR SPEAKER LOGS ---
    const speakerLogsPath = path.join(
      process.cwd(),
      `${this.slightlySecretId}_speakers.json`,
    );
    const speakerLogsArray: Array<{ name: string; timestampSeconds: number }> =
      [];

    // --- FIX 1: FORWARD BROWSER CONSOLE LOGS TO DOCKER LOGS ---
    this.page.on('console', (msg) => {
      this._logger.info(`[Browser Console] ${msg.text()}`);
    });

    await this.page.exposeFunction(
      'screenAppSpeakerLog',
      async (name: string, timestampSeconds: number) => {
        const lastLog = speakerLogsArray[speakerLogsArray.length - 1];

        // Throttle logging: record when the speaker changes or every 2 seconds
        if (
          !lastLog ||
          lastLog.name !== name ||
          timestampSeconds - lastLog.timestampSeconds > 2
        ) {
          speakerLogsArray.push({ name, timestampSeconds });
          this._logger.info(
            `🗣️ Active Speaker: ${name} at [${timestampSeconds.toFixed(1)}s]`,
          );

          // Write to temporary JSON file for post-processing
          await fs.writeJson(speakerLogsPath, speakerLogsArray, { spaces: 2 });
          this._logger.info(`💾 Saved speaker logs to: ${speakerLogsPath}`);
        }
      },
    );

    await this.page.evaluate(
      async ({
        teamId,
        duration,
        inactivityLimit,
        loneParticipantExitDelayMs,
        userId,
        slightlySecretId,
        activateInactivityDetectionAfter,
        activateInactivityDetectionAfterMinutes,
        mimeTypes,
        teamsSpeakerIndicator,
        teamsTileWrapper,
      }: {
        teamId: string;
        duration: number;
        inactivityLimit: number;
        loneParticipantExitDelayMs: number;
        userId: string;
        slightlySecretId: string;
        activateInactivityDetectionAfter: string;
        activateInactivityDetectionAfterMinutes: number;
        mimeTypes: string[];
        teamsSpeakerIndicator: string;
        teamsTileWrapper: string;
      }) => {
        let timeoutId: NodeJS.Timeout;
        let inactivitySilenceDetectionTimeout: NodeJS.Timeout;

        // --- FIX 2: SAFE SELECTOR FALLBACKS ---
        const effectiveIndicator =
          teamsSpeakerIndicator ||
          '[data-tid="voice-level-stream-outline"][data-is-speaking="true"]';
        const effectiveTileWrapper =
          teamsTileWrapper || '[data-tid^="calling-participant-stream"]';

        const sendChunkToServer = async (chunk: ArrayBuffer) => {
          function arrayBufferToBase64(buffer: ArrayBuffer) {
            let binary = '';
            const bytes = new Uint8Array(buffer);
            for (let i = 0; i < bytes.byteLength; i++) {
              binary += String.fromCharCode(bytes[i]);
            }
            return btoa(binary);
          }
          const base64 = arrayBufferToBase64(chunk);
          await (window as any).screenAppSendData(slightlySecretId, base64);
        };

        async function startRecording() {
          console.log('Participant & speaker detection is active...');

          if (
            !navigator.mediaDevices ||
            !navigator.mediaDevices.getDisplayMedia
          ) {
            console.error('MediaDevices or getDisplayMedia not supported.');
            return;
          }

          const stream: MediaStream = await (
            navigator.mediaDevices as any
          ).getDisplayMedia({
            video: true,
            audio: {
              autoGainControl: false,
              channels: 2,
              channelCount: 2,
              echoCancellation: false,
              noiseSuppression: false,
            },
            preferCurrentTab: true,
          });

          const audioOnlyStream = new MediaStream(stream.getAudioTracks());
          const selectedMimeType = 'audio/webm;codecs=opus';

          if (!MediaRecorder.isTypeSupported(selectedMimeType)) {
            throw new Error(
              `MediaRecorder does not support audio codec: ${selectedMimeType}`,
            );
          }

          const mediaRecorder = new MediaRecorder(audioOnlyStream, {
            mimeType: selectedMimeType,
          });
          let chunkUploadChain: Promise<void> = Promise.resolve();
          let isStoppingRecording = false;

          mediaRecorder.ondataavailable = (event: BlobEvent) => {
            if (!event.data.size) return;
            const chunk = event.data;
            chunkUploadChain = chunkUploadChain.then(async () => {
              try {
                const arrayBuffer = await chunk.arrayBuffer();
                await sendChunkToServer(arrayBuffer);
              } catch (error) {
                console.error('Error uploading chunk:', error);
              }
            });
          };

          const chunkDuration = 2000;
          mediaRecorder.start(chunkDuration);
          const recordingStartedAt = Date.now();

          const stopTheRecording = async () => {
            if (isStoppingRecording) return;
            isStoppingRecording = true;
            console.log('-------- TRIGGER stop the recording');
            const recordedDurationSeconds = Math.max(
              1,
              Math.round((Date.now() - recordingStartedAt) / 1000),
            );

            try {
              await new Promise<void>((resolve) => {
                if (mediaRecorder.state === 'inactive') return resolve();
                mediaRecorder.addEventListener('stop', () => resolve(), {
                  once: true,
                });
                mediaRecorder.stop();
              });
              await chunkUploadChain;
            } catch (error) {
              console.error('Error stopping recorder:', error);
            } finally {
              stream.getTracks().forEach((track) => track.stop());
              clearTimeout(timeoutId);
              if (inactivitySilenceDetectionTimeout)
                clearTimeout(inactivitySilenceDetectionTimeout);
              (window as any).screenAppMeetEnd(
                slightlySecretId,
                recordedDurationSeconds,
              );
            }
          };

          // --- DOM SCRAPER ---
          const detectActiveSpeaker = () => {
            let lastSpeaker = '';

            const getMeetingDOMs = () => {
              const doms = [document];
              try {
                const iframes = Array.from(document.querySelectorAll('iframe'));
                iframes.forEach((iframe) => {
                  try {
                    if (iframe.contentDocument) {
                      doms.push(iframe.contentDocument);
                    }
                  } catch (e) {
                    // Ignore cross-origin iframe access errors
                  }
                });
              } catch (e) {
                // Ignore query errors
              }
              return doms;
            };

            let debugLogCounter = 0;

            const speakerCheckTimer = setInterval(() => {
              try {
                if (isStoppingRecording) {
                  clearInterval(speakerCheckTimer);
                  return;
                }

                const activeDoms = getMeetingDOMs();
                let currentSpeakerName = '';

                debugLogCounter++;
                if (debugLogCounter % 20 === 0) {
                  console.log(
                    `🔍 DOM Scraper active: Scanning ${activeDoms.length} DOM context(s)... Target indicator: "${effectiveIndicator}"`,
                  );
                }

                for (const targetDom of activeDoms) {
                  const activeOutlines = Array.from(
                    targetDom.querySelectorAll(effectiveIndicator),
                  );

                  if (activeOutlines.length > 0) {
                    const tile =
                      activeOutlines[0].closest(effectiveTileWrapper);
                    if (tile) {
                      const rawAria = tile.getAttribute('aria-label') || '';
                      currentSpeakerName = rawAria.split(',')[0].trim();
                      if (currentSpeakerName) break;
                    }
                  }

                  if (!currentSpeakerName) {
                    const speakingElements = Array.from(
                      targetDom.querySelectorAll(
                        '[aria-label*="is speaking"], [aria-label*="is talking"]',
                      ),
                    );
                    if (speakingElements.length > 0) {
                      const ariaText =
                        speakingElements[0].getAttribute('aria-label') || '';
                      currentSpeakerName = ariaText
                        .replace('is speaking', '')
                        .replace('is talking', '')
                        .trim();
                      if (currentSpeakerName) break;
                    }
                  }
                }

                if (
                  currentSpeakerName &&
                  currentSpeakerName !== lastSpeaker &&
                  currentSpeakerName.length > 0
                ) {
                  lastSpeaker = currentSpeakerName;
                  const currentVideoOffsetSeconds =
                    (Date.now() - recordingStartedAt) / 1000;

                  console.log(
                    `🎯 DOM Scraper caught active speaker: ${currentSpeakerName}`,
                  );
                  (window as any).screenAppSpeakerLog(
                    currentSpeakerName,
                    currentVideoOffsetSeconds,
                  );
                }
              } catch (err) {
                // Fail silently
              }
            }, 500);
          };

          const detectLoneParticipant = () => {
            /* existing logic */
          };
          const detectIncrediblySilentMeeting = () => {
            /* existing logic */
          };

          detectLoneParticipant();
          detectActiveSpeaker();

          inactivitySilenceDetectionTimeout = setTimeout(
            () => {
              detectIncrediblySilentMeeting();
            },
            activateInactivityDetectionAfterMinutes * 60 * 1000,
          );

          timeoutId = setTimeout(async () => {
            stopTheRecording();
          }, duration);
        }

        await startRecording();
      },
      {
        teamId: this.teamId,
        duration: this.duration,
        inactivityLimit: this.inactivityLimit,
        loneParticipantExitDelayMs,
        userId: this.userId,
        slightlySecretId: this.slightlySecretId,
        activateInactivityDetectionAfterMinutes:
          config.activateInactivityDetectionAfter,
        activateInactivityDetectionAfter: new Date(
          new Date().getTime() +
            config.activateInactivityDetectionAfter * 60 * 1000,
        ).toISOString(),
        mimeTypes,
        teamsSpeakerIndicator: config.teamsSpeakerIndicator,
        teamsTileWrapper: config.teamsTileWrapper,
      },
    );
  }
}
