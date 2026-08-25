import { Page } from 'playwright';
import { Task } from '../lib/Task';
import { JoinParams } from '../bots/AbstractMeetBot';
import { WaitPromise } from '../types';
import { IUploader } from '../middleware/disk-uploader';
import { Logger } from 'winston';
import { browserLogCaptureCallback } from '../util/logger';

export class ContextBridgeTask extends Task<null, void> {
  private page: Page;
  private uploader: IUploader;
  private slightlySecretId: string;
  private waitingPromise: WaitPromise;

  constructor(
    page: Page,
    params: JoinParams & { botId: string },
    slightlySecretId: string,
    waitingPromise: WaitPromise,
    uploader: IUploader,
    logger: Logger,
  ) {
    super(logger);
    this.page = page;
    this.slightlySecretId = slightlySecretId;
    this.waitingPromise = waitingPromise;
    this.uploader = uploader;
  }

  protected async execute(input: null): Promise<void> {
    // Capture and send the browser console logs to Node.js context
    this.page?.on('console', async (msg) => {
      try {
        await browserLogCaptureCallback(this._logger, msg);
      } catch (err) {
        this._logger.info('Failed to log browser messages...', err?.message);
      }
    });

    // 1. Setup tracking variables
    const blockIds: string[] = [];
    const blobKey = `meetings/${this.teamId}/${this.userId}-recording.webm`; // Adjust path as needed

    await this.page.exposeFunction(
      'screenAppSendData',
      async (slightlySecretId: string, data: string) => {
        if (slightlySecretId !== this.slightlySecretId) return;

        const buffer = Buffer.from(data, 'base64');

        // 2. Generate exact-length Base64 block ID
        const rawId = crypto.randomUUID().replace(/-/g, '');
        const blockId = Buffer.from(rawId).toString('base64');
        blockIds.push(blockId);

        // 3. Stage the chunk
        await this.uploader.stageChunk(blobKey, blockId, buffer, this._logger);
      },
    );

    // Make sure to add 'async' here!
    await this.page.exposeFunction(
      'screenAppMeetEnd',
      async (slightlySecretId: string, recordedDurationSeconds?: number) => {
        if (slightlySecretId !== this.slightlySecretId) return;
        try {
          if (typeof recordedDurationSeconds === 'number') {
            this.uploader.setRecordingDuration(recordedDurationSeconds);
          }

          // 4. Commit all chunks to Azure before resolving
          this._logger.info('Stitching chunks in Azure...');
          await this.uploader.commitChunks(
            blobKey,
            blockIds,
            'audio/webm;codecs=opus',
            this._logger,
          );

          this._logger.info('Early signal resolve recording');
          this.waitingPromise.resolveEarly();
        } catch (error) {
          this._logger.error('Could not process meeting end event', error);
        }
      },
    );
  }
}
