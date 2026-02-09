/**
 * Shared constants used across the codebase.
 * 
 * Centralising these values makes the code more maintainable and
 * allows for easier configuration changes.
 */

// ─────────────────────────────────────────────────────────────────────────────
// Content Thresholds
// ─────────────────────────────────────────────────────────────────────────────

/** Minimum content length to be considered valid (characters). */
export const MIN_CONTENT_LENGTH = 5;

/** Maximum length for config values in debug output (characters). */
export const MAX_DEBUG_CONFIG_VALUE_LENGTH = 5000;

// ─────────────────────────────────────────────────────────────────────────────
// Pagination Defaults
// ─────────────────────────────────────────────────────────────────────────────

/** Default page size for search results. */
export const DEFAULT_PAGE_SIZE = 25;

/** Maximum page size for search results. */
export const MAX_PAGE_SIZE = 100;

/** Default limit for thread messages. */
export const DEFAULT_THREAD_LIMIT = 50;

/** Maximum limit for thread messages. */
export const MAX_THREAD_LIMIT = 200;

/** Default limit for people search. */
export const DEFAULT_PEOPLE_LIMIT = 10;

/** Maximum limit for people search. */
export const MAX_PEOPLE_LIMIT = 50;

/** Default limit for frequent contacts. */
export const DEFAULT_CONTACTS_LIMIT = 50;

/** Maximum limit for frequent contacts. */
export const MAX_CONTACTS_LIMIT = 500;

/** Default limit for channel search. */
export const DEFAULT_CHANNEL_LIMIT = 10;

/** Maximum limit for channel search. */
export const MAX_CHANNEL_LIMIT = 50;

// ─────────────────────────────────────────────────────────────────────────────
// Timeouts (milliseconds)
// ─────────────────────────────────────────────────────────────────────────────

/** Default timeout for waiting for search results. */
export const SEARCH_RESULT_TIMEOUT_MS = 10000;

/** Default HTTP request timeout. */
export const HTTP_REQUEST_TIMEOUT_MS = 30000;

/** Short delay for UI interactions. */
export const UI_SHORT_DELAY_MS = 300;

/** Medium delay for UI state changes. */
export const UI_MEDIUM_DELAY_MS = 1000;

/** Long delay for API responses to settle. */
export const UI_LONG_DELAY_MS = 2000;

/** Authentication check interval. */
export const AUTH_CHECK_INTERVAL_MS = 2000;

/** Default login timeout (5 minutes). */
export const LOGIN_TIMEOUT_MS = 5 * 60 * 1000;

/** Pause after showing progress overlay step (ms). */
export const OVERLAY_STEP_PAUSE_MS = 1500;

/** Pause after showing final "All done" overlay (ms). */
export const OVERLAY_COMPLETE_PAUSE_MS = 2000;

// ─────────────────────────────────────────────────────────────────────────────
// Session Management
// ─────────────────────────────────────────────────────────────────────────────

/** Session expiry threshold in hours. */
export const SESSION_EXPIRY_HOURS = 12;

// ─────────────────────────────────────────────────────────────────────────────
// Retry Configuration
// ─────────────────────────────────────────────────────────────────────────────

/** Default maximum retry attempts for HTTP requests. */
export const DEFAULT_MAX_RETRIES = 3;

/** Base delay for exponential backoff (milliseconds). */
export const RETRY_BASE_DELAY_MS = 1000;

/** Maximum delay between retries (milliseconds). */
export const RETRY_MAX_DELAY_MS = 10000;

// ─────────────────────────────────────────────────────────────────────────────
// Conversation IDs
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Virtual Conversation IDs.
 * 
 * These special IDs are used with the standard chatsvc messages endpoint
 * (/users/ME/conversations/{id}/messages) to retrieve aggregated views
 * across all conversations. See docs/API-REFERENCE.md for details.
 */

/** Prefix for virtual conversation IDs (48:saved, 48:notifications, etc). */
export const VIRTUAL_CONVERSATION_PREFIX = '48:';

/** Self-chat (notes) conversation ID. */
export const SELF_CHAT_ID = '48:notes';

/** Activity feed (notifications) conversation ID. */
export const NOTIFICATIONS_ID = '48:notifications';

/** Saved messages virtual conversation ID. */
export const SAVED_MESSAGES_ID = '48:saved';

/** Followed threads virtual conversation ID. */
export const FOLLOWED_THREADS_ID = '48:threads';

// ─────────────────────────────────────────────────────────────────────────────
// Activity Feed
// ─────────────────────────────────────────────────────────────────────────────

/** Default limit for activity feed items. */
export const DEFAULT_ACTIVITY_LIMIT = 50;

/** Maximum limit for activity feed items. */
export const MAX_ACTIVITY_LIMIT = 200;

// ─────────────────────────────────────────────────────────────────────────────
// Unread Status
// ─────────────────────────────────────────────────────────────────────────────

/** Maximum conversations to check when aggregating unread status. */
export const MAX_UNREAD_AGGREGATE_CHECK = 20;

// ─────────────────────────────────────────────────────────────────────────────
// Calendar/Meetings
// ─────────────────────────────────────────────────────────────────────────────

/** Default number of days ahead to fetch meetings. */
export const DEFAULT_MEETING_DAYS_AHEAD = 7;

/** Default limit for meeting results. */
export const DEFAULT_MEETING_LIMIT = 50;

/** Maximum limit for meeting results. */
export const MAX_MEETING_LIMIT = 200;

// ─────────────────────────────────────────────────────────────────────────────
// Token Refresh
// ─────────────────────────────────────────────────────────────────────────────

/** Threshold for proactive token refresh (10 minutes before expiry). */
export const TOKEN_REFRESH_THRESHOLD_MS = 10 * 60 * 1000;

// ─────────────────────────────────────────────────────────────────────────────
// User Identity
// ─────────────────────────────────────────────────────────────────────────────

/** MRI type prefix for Teams/AAD users (type 8). */
export const MRI_TYPE_PREFIX = '8:';

/** Identity type prefix for organisation users. */
export const ORGID_PREFIX = 'orgid:';

/** Full MRI prefix for organisation users (8:orgid:). */
export const MRI_ORGID_PREFIX = `${MRI_TYPE_PREFIX}${ORGID_PREFIX}`;

// ─────────────────────────────────────────────────────────────────────────────
// Emoji Data
// ─────────────────────────────────────────────────────────────────────────────

/** Standard Teams emoji entry. */
export interface StandardEmoji {
  key: string;
  description: string;
  category: 'reaction' | 'expression' | 'affection' | 'action' | 'animal' | 'object' | 'other';
}

/**
 * Standard Teams emoji shortcuts (built-in, no API call needed).
 * These are the common emojis available in Teams for reactions and messages.
 */
export const STANDARD_EMOJIS: readonly StandardEmoji[] = [
  // Quick reactions (shown in reaction picker)
  { key: 'like', description: 'Thumbs up 👍', category: 'reaction' },
  { key: 'heart', description: 'Heart ❤️', category: 'reaction' },
  { key: 'laugh', description: 'Laughing 😂', category: 'reaction' },
  { key: 'surprised', description: 'Surprised 😮', category: 'reaction' },
  { key: 'sad', description: 'Sad 😢', category: 'reaction' },
  { key: 'angry', description: 'Angry 😠', category: 'reaction' },
  // Expressions
  { key: 'smile', description: 'Smiley 😊', category: 'expression' },
  { key: 'wink', description: 'Winking 😉', category: 'expression' },
  { key: 'cry', description: 'Crying 😭', category: 'expression' },
  { key: 'cwl', description: 'Crying with laughter 😂', category: 'expression' },
  { key: 'rofl', description: 'Rolling on floor laughing 🤣', category: 'expression' },
  { key: 'blush', description: 'Blushing 😊', category: 'expression' },
  { key: 'speechless', description: 'Speechless 😶', category: 'expression' },
  { key: 'wonder', description: 'Wondering 🤔', category: 'expression' },
  { key: 'sleepy', description: 'Sleepy 😴', category: 'expression' },
  { key: 'yawn', description: 'Yawning 🥱', category: 'expression' },
  { key: 'eyeroll', description: 'Eye roll 🙄', category: 'expression' },
  { key: 'worry', description: 'Worried 😟', category: 'expression' },
  { key: 'puke', description: 'Puking 🤮', category: 'expression' },
  { key: 'giggle', description: 'Giggling 🤭', category: 'expression' },
  { key: 'tongueout', description: 'Tongue out 😛', category: 'expression' },
  // Affection
  { key: 'kiss', description: 'Kiss 😘', category: 'affection' },
  { key: 'inlove', description: 'In love 😍', category: 'affection' },
  { key: 'hug', description: 'Hug 🤗', category: 'affection' },
  { key: 'lips', description: 'Kissing lips 💋', category: 'affection' },
  // Actions
  { key: 'facepalm', description: 'Facepalm 🤦', category: 'action' },
  { key: 'sweat', description: 'Sweating 😓', category: 'action' },
  { key: 'dance', description: 'Dancing 💃', category: 'action' },
  { key: 'bow', description: 'Bowing 🙇', category: 'action' },
  { key: 'headbang', description: 'Banging head on wall', category: 'action' },
  { key: 'wasntme', description: "It wasn't me 🤷", category: 'action' },
  { key: 'hungover', description: 'Hungover', category: 'action' },
  { key: 'shivering', description: 'Shivering 🥶', category: 'action' },
  // Animals
  { key: 'penguin', description: 'Penguin 🐧', category: 'animal' },
  { key: 'cat', description: 'Cat 🐱', category: 'animal' },
  { key: 'monkey', description: 'Monkey 🐵', category: 'animal' },
  { key: 'polarbear', description: 'Polar bear 🐻‍❄️', category: 'animal' },
  { key: 'elephant', description: 'Elephant 🐘', category: 'animal' },
  // Objects
  { key: 'flower', description: 'Flower 🌸', category: 'object' },
  { key: 'sun', description: 'Sun ☀️', category: 'object' },
  { key: 'star', description: 'Star ⭐', category: 'object' },
  { key: 'xmastree', description: 'Christmas tree 🎄', category: 'object' },
  { key: 'cake', description: 'Cake 🎂', category: 'object' },
  { key: 'gift', description: 'Gift 🎁', category: 'object' },
  { key: 'cash', description: 'Cash 💵', category: 'object' },
  { key: 'champagne', description: 'Champagne 🍾', category: 'object' },
  // Other
  { key: 'yes', description: 'Yes/Thumbs up ✅', category: 'other' },
  { key: 'cool', description: 'Cool 😎', category: 'other' },
  { key: 'party', description: 'Party 🎉', category: 'other' },
  { key: 'hi', description: 'Wave/Hello 👋', category: 'other' },
  { key: 'angel', description: 'Angel 😇', category: 'other' },
  { key: 'devil', description: 'Devil 😈', category: 'other' },
  { key: 'holidayspirit', description: 'Holiday spirit 🎅', category: 'other' },
  { key: 'lipssealed', description: 'Lips sealed 🤐', category: 'other' },
  { key: 'makeup', description: 'Make-up 💄', category: 'other' },
  { key: 'snowangel', description: 'Snow angel', category: 'other' },
];
