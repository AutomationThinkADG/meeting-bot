import dotenv from 'dotenv';
import { UploaderType } from './types';
dotenv.config();

const ENVIRONMENTS = [
  'production',
  'staging',
  'development',
  'cli',
  'test',
] as const;

export type Environment = (typeof ENVIRONMENTS)[number];
export const NODE_ENV: Environment = ENVIRONMENTS.includes(
  process.env.NODE_ENV as Environment
)
  ? (process.env.NODE_ENV as Environment)
  : 'staging';

console.log('NODE_ENV', process.env.NODE_ENV);

const requiredSettings = [
  'GCP_DEFAULT_REGION',
  'GCP_MISC_BUCKET',
];
const missingSettings = requiredSettings.filter((s) => !process.env[s]);
if (missingSettings.length > 0) {
  missingSettings.forEach((ms) =>
    console.error(`ENV settings ${ms} is missing.`)
  );
}

const constructRedisUri = () => {
  const host = process.env.REDIS_HOST || 'redis';
  const port = process.env.REDIS_PORT || 6379;
  const username = process.env.REDIS_USERNAME;
  const password = process.env.REDIS_PASSWORD;

  if (username && password) {
    return `redis://${username}:${password}@${host}:${port}`;
  } else if (password) {
    return `redis://:${password}@${host}:${port}`;
  } else {
    return `redis://${host}:${port}`;
  }
};

const normalizeFileExtension = (extension?: string) => {
  if (!extension) return '.webm';
  return extension.startsWith('.') ? extension : `.${extension}`;
};

const parseOptionalNumber = (value?: string) => {
  if (typeof value === 'undefined' || value.trim() === '') return undefined;
  return Number(value);
};

export default {
  port: process.env.PORT || 3000,
  db: {
    host: process.env.DB_HOST || 'localhost',
    user: process,
  },
  authBaseUrlV2: process.env.AUTH_BASE_URL_V2 ?? 'http://localhost:8081/v2',
  // Unset MAX_RECORDING_DURATION_MINUTES to use default upper limit on duration
  maxRecordingDuration: process.env.MAX_RECORDING_DURATION_MINUTES
    ? Number(process.env.MAX_RECORDING_DURATION_MINUTES)
    : 180, // There's an upper limit on meeting duration 3 hours
  chromeExecutablePath: process.env.CHROME_PATH || '/usr/bin/google-chrome', // We use Google Chrome with Playwright for recording
  googleChromeCdpUrl: process.env.GOOGLE_CHROME_CDP_URL,
  googleChromeUserDataDir: process.env.GOOGLE_CHROME_USER_DATA_DIR,
  googleChromeStorageStatePath: process.env.GOOGLE_CHROME_STORAGE_STATE_PATH,
  googleAnonymousJoinRequestAttempts: process.env
    .GOOGLE_ANONYMOUS_JOIN_REQUEST_ATTEMPTS
    ? Number(process.env.GOOGLE_ANONYMOUS_JOIN_REQUEST_ATTEMPTS)
    : 10,
  inactivityLimit: process.env.MEETING_INACTIVITY_MINUTES
    ? Number(process.env.MEETING_INACTIVITY_MINUTES)
    : 1,
  activateInactivityDetectionAfter: process.env
    .INACTIVITY_DETECTION_START_DELAY_MINUTES
    ? Number(process.env.INACTIVITY_DETECTION_START_DELAY_MINUTES)
    : 1,
  loneParticipantExitDelaySeconds: process.env
    .LONE_PARTICIPANT_EXIT_DELAY_SECONDS
    ? Number(process.env.LONE_PARTICIPANT_EXIT_DELAY_SECONDS)
    : 10,
  serviceKey: process.env.SCREENAPP_BACKEND_SERVICE_API_KEY,
  joinWaitTime: process.env.JOIN_WAIT_TIME_MINUTES
    ? Number(process.env.JOIN_WAIT_TIME_MINUTES)
    : 10,
  // Number of retries for transient errors (not applied to WaitingAtLobbyRetryError)
  retryCount: process.env.RETRY_COUNT ? Number(process.env.RETRY_COUNT) : 2,
  teamsPrewarmEnabled: process.env.TEAMS_PREWARM_ENABLED === 'true',
  teamsAudioStabilizationMs: process.env.TEAMS_AUDIO_STABILIZATION_MS
    ? Number(process.env.TEAMS_AUDIO_STABILIZATION_MS)
    : 1000,
  miscStorageBucket: process.env.GCP_MISC_BUCKET,
  miscStorageFolder: process.env.GCP_MISC_BUCKET_FOLDER
    ? process.env.GCP_MISC_BUCKET_FOLDER
    : 'meeting-bot',
  region: process.env.GCP_DEFAULT_REGION,
  accessKey: process.env.GCP_ACCESS_KEY_ID ?? '',
  accessSecret: process.env.GCP_SECRET_ACCESS_KEY ?? '',
  redisQueueName: process.env.REDIS_QUEUE_NAME ?? 'jobs:meetbot:list',
  redisProcessingQueueName:
    process.env.REDIS_PROCESSING_QUEUE_NAME ?? 'jobs:meetbot:processing',
  redisUri: constructRedisUri(),
  // Notification: Webhook (disabled by default)
  notifyWebhookEnabled: process.env.NOTIFY_WEBHOOK_ENABLED === 'true',
  notifyWebhookUrl: process.env.NOTIFY_WEBHOOK_URL,
  // Optional secret to sign payloads (HMAC-SHA256). If set, signature will be sent in X-Webhook-Signature header
  notifyWebhookSecret: process.env.NOTIFY_WEBHOOK_SECRET,
  // Notification: Redis. Explicitly enabled via NOTIFY_REDIS_ENABLED, and enabled
  // automatically for Redis-worker mode so completed jobs are written to result list.
  notifyRedisEnabled:
    process.env.NOTIFY_REDIS_ENABLED === 'true' ||
    process.env.REDIS_CONSUMER_ENABLED === 'true',
  // If not provided, uses redisUri with specified database selection
  notifyRedisUri: process.env.NOTIFY_REDIS_URI, // optional override
  notifyRedisDb: parseOptionalNumber(process.env.NOTIFY_REDIS_DB),
  notifyRedisList: process.env.NOTIFY_REDIS_LIST ?? 'jobs:meetbot:recordings',
  notifyRedisFailureList:
    process.env.NOTIFY_REDIS_FAILURE_LIST ?? 'jobs:meetbot:failures',
  uploaderFileExtension: normalizeFileExtension(
    process.env.UPLOADER_FILE_EXTENSION,
  ),
  isRedisEnabled: process.env.REDIS_CONSUMER_ENABLED === 'true',
  s3CompatibleStorage: {
    endpoint: process.env.S3_ENDPOINT,
    region: process.env.S3_REGION,
    accessKeyId: process.env.S3_ACCESS_KEY_ID,
    secretAccessKey: process.env.S3_SECRET_ACCESS_KEY,
    bucket: process.env.S3_BUCKET_NAME,
    forcePathStyle: process.env.S3_USE_MINIO_COMPATIBILITY === 'true',
  },
  // Object storage provider selection: 's3' (default) or 'azure'
  storageProvider: (process.env.STORAGE_PROVIDER === 'azure'
    ? 'azure'
    : 's3') as 's3' | 'azure',
  azureBlobStorage: {
    // Either provide full connection string OR account + key/SAS OR managed identity
    connectionString: process.env.AZURE_STORAGE_CONNECTION_STRING,
    accountName: process.env.AZURE_STORAGE_ACCOUNT,
    accountKey: process.env.AZURE_STORAGE_ACCOUNT_KEY, // optional when using connection string
    sasToken: process.env.AZURE_STORAGE_SAS_TOKEN, // starts with ?sv=...
    useManagedIdentity: process.env.AZURE_USE_MANAGED_IDENTITY === 'true',
    container: process.env.AZURE_STORAGE_CONTAINER,
    blobPrefix: process.env.AZURE_BLOB_PREFIX || '',
    signedUrlTtlSeconds: process.env.AZURE_SIGNED_URL_TTL_SECONDS
      ? Number(process.env.AZURE_SIGNED_URL_TTL_SECONDS)
      : 3600,
    uploadConcurrency: process.env.AZURE_UPLOAD_CONCURRENCY
      ? Number(process.env.AZURE_UPLOAD_CONCURRENCY)
      : 4,
  },
  uploaderType: process.env.UPLOADER_TYPE
    ? (process.env.UPLOADER_TYPE as UploaderType)
    : ('s3' as UploaderType),
  // inside your config setup
  // Inside your config file
  // --- Speaker attribution (see docs/SPEAKER_ATTRIBUTION.md) ---------------
  //
  // Every selector below is env-overridable on purpose: the Teams/Meet web DOM
  // is not a contract and shifts without notice, so a broken selector must be a
  // hot config patch, never a redeploy. Keep docs/SPEAKER_ATTRIBUTION.md in sync
  // when you change a default here.

  // Active-speaker "voice level" ring. NOTE: there is no `data-is-speaking`
  // attribute on the current client — speaking is an inline-style animation of
  // this node, detected by observing style mutations. Match the node, not a
  // boolean state.
  teamsSpeakerIndicator:
    process.env.TEAMS_SPEAKER_INDICATOR ||
    '[data-tid="voice-level-stream-outline"]',
  // Participant tile wrapper. On the current client the tile's `data-tid` value
  // is the participant display name; we also fall back to aria-label parsing.
  teamsTileWrapper:
    process.env.TEAMS_TILE_WRAPPER ||
    '[data-tid^="calling-participant-stream"], [data-stream-type][data-tid]',

  // Whether the bot actively prepares the meeting for accurate attribution
  // (enable its own live captions, open the People pane, force gallery view).
  // All of these are per-participant / view-only actions — invisible to other
  // attendees, no host action or consent required.
  teamsReadinessEnabled: process.env.TEAMS_READINESS_ENABLED !== 'false',
  teamsEnableCaptions: process.env.TEAMS_ENABLE_CAPTIONS !== 'false',
  teamsOpenPeoplePane: process.env.TEAMS_OPEN_PEOPLE_PANE !== 'false',
  teamsForceGalleryView: process.env.TEAMS_FORCE_GALLERY_VIEW !== 'false',
  // Spoken-language hint for live captions, e.g. "en-us", "es-es". Empty = leave
  // whatever Teams auto-selects.
  teamsCaptionLanguage: process.env.TEAMS_CAPTION_LANGUAGE || '',

  // Teams live-caption DOM. Comma-separated lists are tried in order.
  teamsCaptionContainerSel:
    process.env.TEAMS_CAPTION_CONTAINER_SEL ||
    '[data-tid="closed-caption-v2-window-wrapper"], [data-tid="closed-caption-renderer-wrapper"], [data-tid="closed-captions-renderer"]',
  teamsCaptionLineSel:
    process.env.TEAMS_CAPTION_LINE_SEL ||
    '.fui-ChatMessageCompact, [data-tid="closed-caption-message"]',
  teamsCaptionAuthorSel:
    process.env.TEAMS_CAPTION_AUTHOR_SEL || '[data-tid="author"]',
  teamsCaptionTextSel:
    process.env.TEAMS_CAPTION_TEXT_SEL || '[data-tid="closed-caption-text"]',
  // How long a caption line's text must stay unchanged before we treat it as
  // final (Teams mutates a line in place as recognition firms up).
  captionFinalizeQuietMs: process.env.CAPTION_FINALIZE_QUIET_MS
    ? Number(process.env.CAPTION_FINALIZE_QUIET_MS)
    : 1200,

  // Hard ceiling on the caption transcript we ship inline with the recording
  // webhook. Above this we ship the speaker timeline (names + times) only and
  // the API falls back to time-overlap fusion.
  captionTranscriptMaxBytes: process.env.CAPTION_TRANSCRIPT_MAX_BYTES
    ? Number(process.env.CAPTION_TRANSCRIPT_MAX_BYTES)
    : 700 * 1024,
};
