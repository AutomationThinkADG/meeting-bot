import { Page } from 'playwright';
import { JoinParams } from './AbstractMeetBot';
import { BotStatus } from '../types';
import config from '../config';
import { RecordingUploadFailedError, WaitingAtLobbyRetryError } from '../error';
import { handleWaitingAtLobbyError, MeetBotBase } from './MeetBotBase';
import { v4 } from 'uuid';
import { patchBotStatus } from '../services/botService';
import { IUploader } from '../middleware/disk-uploader';
import { Logger } from 'winston';
import { retryActionWithWait } from '../util/resilience';
import { uploadDebugImage } from '../services/bugService';
import createBrowserContext from '../lib/chromium';
import { browserLogCaptureCallback } from '../util/logger';
import { MICROSOFT_REQUEST_DENIED } from '../constants';
import { FFmpegRecorder } from '../lib/ffmpegRecorder';
import * as path from 'path';
import * as fs from 'fs';
import { exec } from 'child_process';
import { promisify } from 'util';

const execAsync = promisify(exec);

export class MicrosoftTeamsBot extends MeetBotBase {
  private _logger: Logger;
  private _correlationId: string;
  constructor(logger: Logger, correlationId: string) {
    super();
    this.slightlySecretId = v4();
    this._logger = logger;
    this._correlationId = correlationId;
  }
  async join({
    url,
    name,
    bearerToken,
    teamId,
    timezone,
    userId,
    eventId,
    botId,
    uploader,
  }: JoinParams): Promise<void> {
    const _state: BotStatus[] = ['processing'];

    const handleUpload = async () => {
      this._logger.info('Begin recording upload to server', { userId, teamId });
      const uploadResult = await uploader.uploadRecordingToRemoteStorage();
      this._logger.info('Recording upload result', {
        uploadResult,
        userId,
        teamId,
      });
      return uploadResult;
    };

    try {
      const pushState = (st: BotStatus) => _state.push(st);
      await this.joinMeeting({
        url,
        name,
        bearerToken,
        teamId,
        timezone,
        userId,
        eventId,
        botId,
        pushState,
        uploader,
      });

      // Finish the upload from the temp video
      const uploadResult = await handleUpload();

      if (_state.includes('finished') && !uploadResult) {
        _state.splice(_state.indexOf('finished'), 1, 'failed');
        this._logger.error(
          'Recording completed but upload failed; raising non-retryable failure so JobStore does not rejoin the ended meeting',
          { botId, userId, teamId },
        );
        throw new RecordingUploadFailedError(
          'Microsoft Teams recording completed but upload failed',
        );
      } else if (uploadResult) {
        this._logger.info('Recording and upload completed successfully', {
          botId,
          userId,
          teamId,
        });
      }

      await patchBotStatus(
        {
          botId,
          eventId,
          provider: 'microsoft',
          status: _state,
          token: bearerToken,
        },
        this._logger,
      );
    } catch (error) {
      this._logger.error('Error in Microsoft Teams bot join process', {
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
        botId,
        userId,
        teamId,
        currentState: _state,
      });

      if (!_state.includes('finished') && !_state.includes('failed'))
        _state.push('failed');

      await patchBotStatus(
        {
          botId,
          eventId,
          provider: 'microsoft',
          status: _state,
          token: bearerToken,
        },
        this._logger,
      );

      if (error instanceof WaitingAtLobbyRetryError)
        await handleWaitingAtLobbyError(
          { token: bearerToken, botId, eventId, provider: 'microsoft', error },
          this._logger,
        );

      throw error;
    } finally {
      try {
        const browser = this.page?.context().browser();
        if (browser?.isConnected()) {
          await browser.close();
          this._logger.info('Browser closed in join finally');
        } else if (this.page?.context()) {
          await this.page.context().close();
          this._logger.info(
            'Persistent browser context closed in join finally',
          );
        }
      } catch (cleanupErr) {
        this._logger.warn(
          'Browser cleanup in join finally failed (non-fatal)',
          { error: cleanupErr },
        );
      }
    }
  }

  private async joinMeeting({
    url,
    name,
    teamId,
    userId,
    eventId,
    botId,
    pushState,
    uploader,
  }: JoinParams & { pushState(state: BotStatus): void }): Promise<void> {
    const joinButtonSelectors = [
      'button[aria-label="Join meeting from this browser"]',
      'button[aria-label="Continue on this browser"]',
      'button[aria-label="Join on this browser"]',
      'button:has-text("Continue on this browser")',
      'button:has-text("Join from browser")',
    ];

    const clickFirstVisibleSelector = async (
      page: Page,
      selectors: string[],
      timeoutMs: number,
      logPrefix: string,
    ): Promise<boolean> => {
      const startedAt = Date.now();
      for (const selector of selectors) {
        this._logger.info(`${logPrefix}: checking selector`, { selector });
      }

      while (Date.now() - startedAt < timeoutMs) {
        for (const selector of selectors) {
          const button = page.locator(selector).first();
          if (await button.isVisible({ timeout: 500 }).catch(() => false)) {
            this._logger.info(`${logPrefix}: found button`, { selector });
            await button.click({ force: true });
            return true;
          }
        }
        await page.waitForTimeout(500);
      }

      return false;
    };

    if (config.teamsPrewarmEnabled) {
      this._logger.info(
        'Pre-warming: Opening browser to trigger first-run dialogs...',
      );
      let warmupPage: Page | undefined;
      try {
        warmupPage = await createBrowserContext(
          url,
          this._correlationId,
          'microsoft',
        );
        this._logger.info('Pre-warming: Navigating to Teams meeting...');
        await warmupPage.goto(url, { waitUntil: 'domcontentloaded' });
        await clickFirstVisibleSelector(
          warmupPage,
          joinButtonSelectors,
          8000,
          'Pre-warming',
        );
        await warmupPage
          .locator('input[data-tid="prejoin-display-name-input"]')
          .waitFor({ state: 'visible', timeout: 8000 })
          .catch(() => undefined);
        this._logger.info('Pre-warming complete - dialogs triggered');
      } catch (error) {
        this._logger.warn('Pre-warming failed (non-fatal):', error);
      } finally {
        try {
          const browser = warmupPage?.context().browser();
          if (browser?.isConnected()) {
            this._logger.info('Pre-warming: Closing warmup browser...');
            await browser.close();
          }
        } catch (cleanupErr) {
          this._logger.warn(
            'Pre-warming: warmup browser cleanup failed (non-fatal)',
            { error: cleanupErr },
          );
        }
      }
    } else {
      this._logger.info(
        'Teams pre-warming disabled; launching the meeting browser directly',
      );
    }

    this._logger.info('Launching browser for actual meeting...');
    this.page = await createBrowserContext(
      url,
      this._correlationId,
      'microsoft',
    );

    this._logger.info('Navigating to Microsoft Teams Meeting URL...');
    await this.page.goto(url, { waitUntil: 'domcontentloaded' });

    this._logger.info('Waiting for Join meeting from browser button...');
    const buttonClicked = await clickFirstVisibleSelector(
      this.page,
      joinButtonSelectors,
      60000,
      'Join from browser',
    );

    if (!buttonClicked) {
      this._logger.info(
        'Join from browser button not found, proceeding anyway...',
      );
    }

    this._logger.info('Waiting for pre-join screen to load...');

    try {
      this._logger.info('Looking for name input field...');
      const nameInput = this.page.locator(
        'input[data-tid="prejoin-display-name-input"]',
      );
      await nameInput.waitFor({ state: 'visible', timeout: 45000 });
      this._logger.info('Found name input field, filling with bot name...');
      await nameInput.fill(name ? name : 'ScreenApp Notetaker');
    } catch (err) {
      this._logger.info(
        'Name input field not found after 45s, skipping...',
        err instanceof Error ? err.message : String(err),
      );
    }

    const toggleDevices = async () => {
      try {
        this._logger.info(
          'Attempting to turn off camera and mute microphone...',
        );

        // Turn off camera
        try {
          const cameraSelectors = [
            'input[data-tid="toggle-video"][checked]',
            'input[type="checkbox"][title*="Turn camera off" i]',
            'input[role="switch"][data-tid="toggle-video"]',
            'button[aria-label*="Turn camera off" i]',
            'button[aria-label*="Camera off" i]',
          ];

          for (const selector of cameraSelectors) {
            const cameraButton = this.page.locator(selector).first();
            const isVisible = await cameraButton
              .isVisible({ timeout: 1000 })
              .catch(() => false);
            if (isVisible) {
              const label = await cameraButton.getAttribute('aria-label');
              this._logger.info(`Clicking camera toggle: ${label}`);
              await cameraButton.click();
              await this.page.waitForTimeout(250);
              break;
            }
          }
        } catch (err) {}

        // Mute microphone
        try {
          const micSelectors = [
            'input[data-tid="toggle-mute"]:not([checked])',
            'input[type="checkbox"][title*="Mute mic" i]',
            'input[role="switch"][data-tid="toggle-mute"]',
            'button[aria-label*="Mute microphone" i]',
            'button[aria-label*="Mute mic" i]',
          ];

          for (const selector of micSelectors) {
            const micButton = this.page.locator(selector).first();
            const isVisible = await micButton
              .isVisible({ timeout: 1000 })
              .catch(() => false);
            if (isVisible) {
              const label = await micButton.getAttribute('aria-label');
              this._logger.info(`Clicking microphone toggle: ${label}`);
              await micButton.click();
              await this.page.waitForTimeout(250);
              break;
            }
          }
        } catch (err) {}

        this._logger.info('Finished toggling camera and microphone');
      } catch (error) {
        this._logger.warn(
          'Error toggling devices',
          error instanceof Error ? error.message : String(error),
        );
      }
    };

    await toggleDevices();

    this._logger.info('Clicking the join button...');
    await retryActionWithWait(
      'Clicking the join button',
      async () => {
        const possibleTexts = [
          'Join now',
          'Join',
          'Ask to join',
          'Join meeting',
        ];
        let buttonClicked = false;

        for (const text of possibleTexts) {
          try {
            const button = this.page.getByRole('button', {
              name: new RegExp(text, 'i'),
            });
            if (await button.isVisible({ timeout: 1000 }).catch(() => false)) {
              await button.click();
              buttonClicked = true;
              this._logger.info(`Successfully clicked "${text}" button`);
              break;
            }
          } catch (err) {
            this._logger.info(
              `Unable to click "${text}" button, trying next...`,
            );
          }
        }

        if (!buttonClicked) {
          throw new Error('Unable to find any join button variant');
        }
      },
      this._logger,
      3,
      15000,
      async () => {
        await uploadDebugImage(
          await this.page.screenshot({ type: 'png', fullPage: true }),
          'join-button-click',
          userId,
          this._logger,
          botId,
        );
      },
    );

    try {
      const wanderingTime = config.joinWaitTime * 60 * 1000;
      const callButton = this.page.getByRole('button', { name: /Leave/i });
      await callButton.waitFor({ timeout: wanderingTime });
      this._logger.info('Bot is entering the meeting...');
    } catch (error) {
      const bodyText = await this.page.evaluate(() => document.body.innerText);
      const userDenied = (bodyText || '')?.includes(MICROSOFT_REQUEST_DENIED);

      this._logger.error('Cant finish wait at the lobby check', {
        userDenied,
        waitingAtLobbySuccess: false,
        bodyText,
      });

      this._logger.error('Closing the browser on error...', error);
      await this.page.context().browser()?.close();

      throw new WaitingAtLobbyRetryError(
        'Microsoft Teams Meeting bot could not enter the meeting...',
        bodyText ?? '',
        false,
        0,
      );
    }

    pushState('joined');

    const dismissDeviceChecksAndNotifications = async () => {
      const closeSelectors = [
        'button[aria-label=Close]:visible',
        'button[title="Close"]:visible',
      ];
      const startedAt = Date.now();
      let closeButtonsClicked = 0;
      let emptyPasses = 0;

      while (Date.now() - startedAt < 3000) {
        let clickedOnPass = false;
        for (const selector of closeSelectors) {
          const visibleButtons = await this.page.locator(selector).all();
          for (const btn of visibleButtons) {
            try {
              await btn.click({ timeout: 1000 });
              closeButtonsClicked++;
              clickedOnPass = true;
            } catch (err) {}
          }
        }

        if (!clickedOnPass) {
          emptyPasses++;
          if (emptyPasses >= 2) break;
          await this.page.waitForTimeout(250);
        } else {
          emptyPasses = 0;
        }
      }

      this._logger.info('Finished dismissing device checks and notifications', {
        closeButtonsClicked,
      });
    };
    await dismissDeviceChecksAndNotifications();

    if (config.teamsAudioStabilizationMs > 0) {
      this._logger.info(
        'Waiting briefly for audio to stabilize before recording...',
        { ms: config.teamsAudioStabilizationMs },
      );
      await this.page.waitForTimeout(config.teamsAudioStabilizationMs);
    }

    this._logger.info('Begin recording with ffmpeg...');
    await this.recordMeetingPageWithFFmpeg({
      teamId,
      userId,
      eventId,
      botId,
      uploader,
    });

    pushState('finished');
  }

  private async recordMeetingPageWithFFmpeg({
    teamId,
    userId,
    eventId,
    botId,
    uploader,
  }: {
    teamId: string;
    userId: string;
    eventId?: string;
    botId?: string;
    uploader: IUploader;
  }): Promise<void> {
    const duration = config.maxRecordingDuration * 60 * 1000;
    this._logger.info(
      `Recording max duration set to ${duration / 60000} minutes (safety limit only)`,
    );

    const tempFolder = path.join(process.cwd(), 'dist', '_tempvideo');
    const outputPath = path.join(
      tempFolder,
      `recording-${botId || Date.now()}${config.uploaderFileExtension}`,
    );

    this._logger.info('Starting ffmpeg recording...', { outputPath, duration });

    try {
      const { stdout: paStatus } = await execAsync('pactl list sources short');
      this._logger.info(
        'PulseAudio sources available:',
        paStatus.trim() || '(empty - no sources found)',
      );

      if (!paStatus.includes('virtual_output.monitor')) {
        this._logger.error(
          'WARNING: virtual_output.monitor not found in PulseAudio sources!',
        );
        try {
          await execAsync('pulseaudio --kill || true');
          await new Promise((resolve) => setTimeout(resolve, 1000));
          await execAsync('pulseaudio -D --exit-idle-time=-1 --log-level=info');
          await new Promise((resolve) => setTimeout(resolve, 1000));

          await execAsync(
            'pactl load-module module-null-sink sink_name=virtual_output sink_properties=device.description="Virtual_Output"',
          );
          await execAsync('pactl set-default-sink virtual_output');
        } catch (err) {}
      }
    } catch (err) {}

    const recorder = new FFmpegRecorder(outputPath, this._logger);

    let ffmpegFailed = false;
    let ffmpegError: Error | null = null;
    let recordingStartedAt: number | undefined;
    let meetingEnded = false;

    try {
      await recorder.start();
      recordingStartedAt = Date.now();
      const startedAt = recordingStartedAt;
      this._logger.info('FFmpeg recording started successfully');

      recorder.onProcessExit((code) => {
        if (code !== 0 && code !== null) {
          this._logger.error('FFmpeg died unexpectedly during recording', {
            exitCode: code,
          });
          ffmpegFailed = true;
          ffmpegError = new Error(
            `FFmpeg exited with code ${code} during recording`,
          );
        }
      });

      await this.page.exposeFunction('screenAppMeetEnd', () => {
        this._logger.info('Meeting ended signal received from browser');
        meetingEnded = true;
      });

      this.page.on('console', async (msg) => {
        try {
          await browserLogCaptureCallback(this._logger, msg);
        } catch (err) {}
      });
      // --- NEW BULLETPROOF NODE.JS ACTIVE SPEAKER POLLER ---
      const speakerLogsPath = path.join(
        process.cwd(),
        'debug-videos',
        `${botId || 'unknown'}_speakers.json`,
      );
      const speakerLogsArray: Array<{
        name: string;
        timestampSeconds: number;
      }> = [];
      let lastSpeaker = '';
      let pollerHeartbeat = 0;

      const pollActiveSpeaker = async () => {
        if (meetingEnded || ffmpegFailed) return;

        try {
          // 1. THE HEARTBEAT (Log every 10 seconds to prove it is running)
          pollerHeartbeat++;
          if (pollerHeartbeat % 10 === 0) {
            this._logger.info(
              `🔍 Node.js Poller active: Scanning ${this.page.frames().length} browser frames...`,
            );
          }

          let currentSpeakerName = '';
          const indicatorSel =
            (config as any).teamsSpeakerIndicator ||
            '[data-tid="voice-level-stream-outline"][data-is-speaking="true"]';
          const wrapperSel =
            (config as any).teamsTileWrapper ||
            '[data-tid^="calling-participant-stream"]';

          // By querying this.page.frames(), we bypass all cross-origin iframe security limits
          for (const frame of this.page.frames()) {
            try {
              // 2. Primary Strict Selector Search
              const indicators = frame.locator(indicatorSel);
              if ((await indicators.count().catch(() => 0)) > 0) {
                const firstIndicator = indicators.first();
                if (await firstIndicator.isVisible().catch(() => false)) {
                  currentSpeakerName = await firstIndicator
                    .evaluate((el, wSel) => {
                      const tile = el.closest(wSel);
                      return tile
                        ? (tile.getAttribute('aria-label') || '')
                            .split(',')[0]
                            .trim()
                        : '';
                    }, wrapperSel)
                    .catch(() => '');

                  if (currentSpeakerName) break;
                }
              }

              // 3. Brute-Force Broad Fallback Search
              if (!currentSpeakerName) {
                const fallbacks = frame.locator(
                  '[aria-label*="is speaking" i], [aria-label*="is talking" i], [data-is-speaking="true"]',
                );
                const fallbackCount = await fallbacks.count().catch(() => 0);

                if (fallbackCount > 0) {
                  for (let i = 0; i < fallbackCount; i++) {
                    const el = fallbacks.nth(i);
                    if (await el.isVisible().catch(() => false)) {
                      const aria = await el
                        .getAttribute('aria-label')
                        .catch(() => '');
                      if (
                        aria &&
                        (aria.toLowerCase().includes('is speaking') ||
                          aria.toLowerCase().includes('is talking'))
                      ) {
                        currentSpeakerName = aria
                          .replace(/is speaking|is talking/gi, '')
                          .trim();
                        if (currentSpeakerName) break;
                      }
                    }
                  }
                }
              }
            } catch (e) {
              // Ignore destroyed or reloading frames
            }
            if (currentSpeakerName) break;
          }

          // 4. Log and Save
          if (currentSpeakerName && currentSpeakerName !== lastSpeaker) {
            lastSpeaker = currentSpeakerName;
            const ts = (Date.now() - (recordingStartedAt || Date.now())) / 1000;
            speakerLogsArray.push({
              name: currentSpeakerName,
              timestampSeconds: ts,
            });
            this._logger.info(
              `🗣️ Active Speaker: ${currentSpeakerName} at [${ts.toFixed(1)}s]`,
            );
            fs.writeFileSync(
              speakerLogsPath,
              JSON.stringify(speakerLogsArray, null, 2),
            );
          }
        } catch (err) {
          // Keep loop alive if a network exception happens
        }

        if (!meetingEnded && !ffmpegFailed) {
          setTimeout(pollActiveSpeaker, 1000); // Check every 1 second continuously
        }
      };

      // Start the speaker detection loop
      pollActiveSpeaker();
      // --- END NODE.JS SPEAKER POLLER ---
      // --- END NODE.JS SPEAKER POLLER ---

      const inactivityLimitMs = config.inactivityLimit * 60 * 1000;

      const monitorAudioSilence = async () => {
        try {
          let consecutiveSilentChecks = 0;
          const checkIntervalSeconds = 5;
          const checksNeeded = Math.ceil(
            inactivityLimitMs / 1000 / checkIntervalSeconds,
          );

          const checkInterval = setInterval(async () => {
            if (meetingEnded) {
              clearInterval(checkInterval);
              return;
            }
            try {
              const { stdout } = await execAsync(
                'timeout 1 parec --device=virtual_output.monitor --format=s16le --rate=16000 --channels=1 2>/dev/null | ' +
                  'od -An -td2 -v | awk \'BEGIN{max=0} {for(i=1;i<=NF;i++) {val=($i<0)?-$i:$i; if(val>max) max=val}} END{print max}\'',
              );

              const peakLevel = parseInt(stdout.trim()) || 0;
              const silenceThreshold = 200;

              if (peakLevel < silenceThreshold) {
                consecutiveSilentChecks++;

                if (consecutiveSilentChecks >= checksNeeded) {
                  this._logger.warn(
                    'Audio silence threshold reached, ending Microsoft Teams meeting',
                  );
                  clearInterval(checkInterval);
                  meetingEnded = true;
                }
              } else {
                consecutiveSilentChecks = 0;
              }
            } catch (err) {}
          }, 5000);
        } catch (error) {
          this._logger.warn('Will rely on participant detection only');
        }
      };

      setTimeout(
        () => {
          monitorAudioSilence();
        },
        config.activateInactivityDetectionAfter * 60 * 1000,
      );

      // We only inject the participant counter / inactivity detector into the browser now
      await this.page.evaluate(
        ({
          activateAfterMinutes,
          loneParticipantExitDelayMs,
          maxDuration,
        }: {
          activateAfterMinutes: number;
          loneParticipantExitDelayMs: number;
          maxDuration: number;
        }) => {
          setTimeout(() => {
            console.log(
              `Max recording duration (${maxDuration / 60000} minutes) reached, ending meeting`,
            );
            (window as any).screenAppMeetEnd();
          }, maxDuration);
          console.log(
            `Max duration timeout set to ${maxDuration / 60000} minutes (safety limit)`,
          );

          console.log('Activating participant count detection...');

          const recordingStartedAt = Date.now();
          const initialAloneGraceMs = activateAfterMinutes * 60 * 1000;

          let hasSeenOtherParticipant = false;
          let aloneSince: number | null = null;
          let lastParticipantDetectionLogAt = 0;

          const shouldStopForParticipantCount = (participants: number) => {
            const now = Date.now();
            if (participants >= 2) {
              hasSeenOtherParticipant = true;
              aloneSince = null;
              return false;
            }

            if (hasSeenOtherParticipant) {
              if (aloneSince === null) {
                aloneSince = now;
                console.log(
                  'Bot is alone after previously seeing participants; waiting before ending recording.',
                );
              }
              return now - aloneSince >= loneParticipantExitDelayMs;
            }

            return now - recordingStartedAt >= initialAloneGraceMs;
          };

          const normalizeText = (text: string) =>
            text.replace(/\s+/g, ' ').trim();

          const parseParticipantCount = (text: string): number | undefined => {
            const normalized = normalizeText(text);
            const patterns = [
              /\b(?:people|participants?|teilnehm(?:er|ende)?|personen)\D{0,30}(\d{1,3})\b/i,
              /\b(\d{1,3})\D{0,30}(?:people|participants?|teilnehm(?:er|ende)?|personen)\b/i,
            ];

            for (const pattern of patterns) {
              const match = normalized.match(pattern);
              if (match) {
                const value = Number(match[1]);
                if (Number.isFinite(value)) {
                  return value;
                }
              }
            }

            if (/^\D*\d{1,3}\D*$/.test(normalized) && normalized.length <= 16) {
              const match = normalized.match(/\d{1,3}/);
              const value = match ? Number(match[0]) : NaN;
              if (Number.isFinite(value)) {
                return value;
              }
            }

            return undefined;
          };

          const isExplicitEmptyMeetingText = (text: string): boolean => {
            const emptyMeetingPatterns = [
              /\b0\D{0,30}(?:people|participants?|teilnehm(?:er|ende)?|personen)\D{0,80}(?:in|inside|joined|here|meeting|call|besprechung|anruf)\b/i,
              /\b(?:no|zero)\D{0,30}(?:one|one else|people|participants?)\D{0,80}(?:in|inside|joined|here|meeting|call)\b/i,
              /\b(?:keine|niemand)\D{0,80}(?:teilnehm(?:er|ende)?|personen|hier|besprechung|anruf)\b/i,
            ];

            return emptyMeetingPatterns.some((pattern) => pattern.test(text));
          };

          const getTeamsMeetingState = ():
            | 'active'
            | 'alone'
            | 'empty'
            | 'ended' => {
            const bodyText = normalizeText(document.body.innerText || '');

            const endedPhrases = [
              'the meeting has ended',
              'this meeting has ended',
              'meeting has been ended',
              'call ended',
              'you have been removed',
              'you’ve been removed',
              'removed from the meeting',
              'besprechung wurde beendet',
              'anruf beendet',
              'sie wurden entfernt',
              'du wurdest entfernt',
            ];

            if (
              endedPhrases.some((phrase) =>
                bodyText.toLowerCase().includes(phrase),
              )
            ) {
              return 'ended';
            }

            if (isExplicitEmptyMeetingText(bodyText)) {
              return 'empty';
            }

            const alonePhrases = [
              'you\'re the only one here',
              'you’re the only one here',
              'you are the only one here',
              'you\'re the only one in this meeting',
              'you’re the only one in this meeting',
              'you are the only one in this meeting',
              'only one in this meeting',
              'only you are here',
              'no one else is here',
              'waiting for others to join',
              'sie sind der einzige',
              'du bist der einzige',
              'sie sind die einzige person',
              'du bist die einzige person',
              'warten auf andere',
            ];

            return alonePhrases.some((phrase) =>
              bodyText.toLowerCase().includes(phrase),
            )
              ? 'alone'
              : 'active';
          };

          const getParticipantCount = (): {
            count?: number;
            samples: string[];
          } => {
            const selectors = [
              'button[data-tid*="roster" i]',
              '[data-tid*="roster" i]',
              'button[id*="roster" i]',
              '[id*="roster" i]',
              'button[aria-label*="people" i]',
              '[aria-label*="people" i]',
              'button[aria-label*="participant" i]',
              '[aria-label*="participant" i]',
              'button[aria-label*="teilnehm" i]',
              '[aria-label*="teilnehm" i]',
              'button[aria-label*="personen" i]',
              '[aria-label*="personen" i]',
            ];

            const candidates = Array.from(
              document.querySelectorAll(selectors.join(',')),
            );
            const samples: string[] = [];

            for (const element of candidates) {
              const searchRoots = [
                element,
                element.parentElement,
                element.parentElement?.parentElement,
              ].filter(Boolean) as Element[];

              for (const root of searchRoots) {
                const text = normalizeText(
                  [
                    root.getAttribute('aria-label') ?? '',
                    root.getAttribute('title') ?? '',
                    root.getAttribute('data-tid') ?? '',
                    root.textContent ?? '',
                  ].join(' '),
                );

                if (!text) continue;
                if (samples.length < 6) {
                  samples.push(text.slice(0, 140));
                }

                const count = parseParticipantCount(text);
                if (typeof count === 'number') {
                  return { count, samples };
                }
              }
            }

            const bodyLines = (document.body.innerText || '')
              .split(/\n+/)
              .map(normalizeText)
              .filter(
                (text) =>
                  text.length > 0 &&
                  /(?:people|participants?|teilnehm|personen|meeting|call|besprechung|anruf)/i.test(
                    text,
                  ),
              );

            for (const text of bodyLines) {
              if (samples.length < 6) {
                samples.push(text.slice(0, 140));
              }

              if (isExplicitEmptyMeetingText(text)) {
                return { count: 0, samples };
              }

              if (/(?:people|participants?|teilnehm|personen)/i.test(text)) {
                const count = parseParticipantCount(text);
                if (typeof count === 'number') {
                  return { count, samples };
                }
              }
            }

            return { samples };
          };

          const interval = setInterval(() => {
            try {
              const meetingState = getTeamsMeetingState();
              if (meetingState === 'ended') {
                console.log(
                  'Teams meeting ended page state detected, ending recording.',
                );
                clearInterval(interval);
                (window as any).screenAppMeetEnd();
                return;
              }

              const { count, samples } = getParticipantCount();
              let inferredCount = count;
              if (typeof inferredCount !== 'number') {
                if (meetingState === 'empty') {
                  inferredCount = 0;
                } else if (meetingState === 'alone') {
                  inferredCount = 1;
                }
              }

              if (typeof inferredCount !== 'number') {
                const now = Date.now();
                if (now - lastParticipantDetectionLogAt > 30000) {
                  console.log('Teams participant count not detected yet', {
                    samples,
                  });
                  lastParticipantDetectionLogAt = now;
                }
                return;
              }

              if (!shouldStopForParticipantCount(inferredCount)) {
                return;
              }

              console.log('Bot is alone, ending Teams recording', {
                inferredCount,
                meetingState,
              });
              clearInterval(interval);
              (window as any).screenAppMeetEnd();
            } catch (error) {}
          }, 2000);
        },
        {
          activateAfterMinutes: config.activateInactivityDetectionAfter,
          loneParticipantExitDelayMs:
            config.loneParticipantExitDelaySeconds * 1000,
          maxDuration: duration,
        },
      );

      // Wait for either timeout, meeting end, or FFmpeg failure
      while (
        !meetingEnded &&
        !ffmpegFailed &&
        Date.now() - startedAt < duration
      ) {
        await new Promise((resolve) => setTimeout(resolve, 1000));
      }

      this._logger.info('Recording period ended', {
        meetingEnded,
        ffmpegFailed,
        recordedDuration: Math.floor((Date.now() - startedAt) / 1000) + 's',
      });

      if (ffmpegFailed && ffmpegError) {
        throw ffmpegError;
      }
    } catch (error) {
      this._logger.error('Error during recording:', error);
      ffmpegFailed = true;
      ffmpegError = error instanceof Error ? error : new Error(String(error));
      throw error;
    } finally {
      meetingEnded = true;

      this._logger.info('Stopping ffmpeg recording...');
      await recorder.stop();

      this._logger.info('Staging recorded file for upload...', { outputPath });

      let staged = false;
      if (fs.existsSync(outputPath)) {
        if (recordingStartedAt) {
          const recordedDurationSeconds = Math.max(
            1,
            Math.round((Date.now() - recordingStartedAt) / 1000),
          );
          uploader.setRecordingDuration(recordedDurationSeconds);
        }

        const fileBuffer = fs.readFileSync(outputPath);
        await uploader.saveDataToTempFile(fileBuffer);

        fs.unlinkSync(outputPath);
        this._logger.info(
          'Recording staged to temp; ffmpeg output file removed',
        );
        staged = true;
      } else {
        this._logger.error('Recording file not found!', { outputPath });
      }

      this._logger.info('Closing the browser...');
      await this.page.context().browser()?.close();

      if (ffmpegFailed) {
        this._logger.error('Recording failed due to FFmpeg error', {
          botId,
          eventId,
          userId,
          teamId,
        });
      } else if (!staged) {
        this._logger.error('Recording file missing; nothing to upload', {
          botId,
          eventId,
          userId,
          teamId,
        });
      } else {
        this._logger.info(
          'Recording captured and staged; finalizing upload next...',
          { botId, eventId, userId, teamId },
        );
      }
    }
  }
}
