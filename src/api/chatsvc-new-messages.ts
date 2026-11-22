/**
 * Chat Service API - New messages across conversations.
 *
 * Polls every conversation for messages newer than an opaque cursor and
 * returns compact summaries. One conversations-list call establishes which
 * conversations have new activity; only those are fetched, so an idle poll
 * costs a single request instead of one per chat.
 */

import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok, err } from '../types/result.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { stripHtml } from '../utils/parsers.js';
import { ErrorCode, createError } from '../types/errors.js';
import {
  getConversations,
  type ChatSummary,
  type ChatType,
} from './chatsvc-conversations.js';
import { buildThreadMessagesUrl } from './chatsvc-messaging.js';
import type { RawChatsvcMessage } from '../types/api-responses.js';
import {
  DEFAULT_NEW_MESSAGES_MAX_CHATS,
  DEFAULT_NEW_MESSAGES_PER_CHAT,
  MAX_NEW_MESSAGES_PER_CHAT,
  NEW_MESSAGES_FETCH_CONCURRENCY,
  NEW_MESSAGES_FLOOR_SKEW_MS,
  NEW_MESSAGES_SNIPPET_MAX_LENGTH,
  MAX_THREAD_LIMIT,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Opaque resume state for {@link getNewMessages}.
 *
 * `since` is the poll floor used for conversations the cursor has never seen
 * (including new chats). `chats` holds a high-water message ID per conversation,
 * so each chat resumes exactly where it left off.
 */
export interface NewMessagesCursor {
  version: 1;
  since: number;
  chats: Record<string, string>;
}

/** A compact new message as returned by the feed. */
export interface NewMessageItem {
  conversationId: string;
  /** Channel/group/meeting topic, or empty for 1:1 chats. */
  topic: string;
  chatType: ChatType;
  messageId: string;
  senderName: string;
  sentAt: string;
  /** Plain-text preview, at most {@link NEW_MESSAGES_SNIPPET_MAX_LENGTH} characters. */
  snippet: string;
}

/** A conversation that could not be scanned this poll. */
export interface NewMessagesError {
  conversationId: string;
  code: ErrorCode;
  message: string;
}

/** Result of polling for new messages. */
export interface GetNewMessagesResult {
  /** New messages, oldest first (chronological across conversations). */
  items: NewMessageItem[];
  /** Cursor to pass as `after` on the next poll. */
  nextAfter: string;
  /** True when more conversations or messages remain beyond this page. */
  truncated: boolean;
  /** Per-conversation failures; other conversations are still returned. */
  errors: NewMessagesError[];
}

/** Options for {@link getNewMessages}. */
export interface GetNewMessagesOptions {
  /** `nextAfter` from a previous response. Omit on the first run. */
  after?: string;
  /** Maximum conversations to scan, most recently active first. */
  maxChats?: number;
  /** Maximum messages fetched per conversation. */
  limitPerChat?: number;
  /** Include the caller's own messages (default false). */
  includeOwn?: boolean;
  /** Injectable clock for tests. */
  now?: number;
}

// ─────────────────────────────────────────────────────────────────────────────
// Cursor
// ─────────────────────────────────────────────────────────────────────────────

/** Encodes a cursor as opaque base64url. */
export function encodeNewMessagesCursor(cursor: NewMessagesCursor): string {
  return Buffer.from(JSON.stringify(cursor), 'utf8').toString('base64url');
}

/** Decodes an opaque cursor, returning null when it is not a valid v1 cursor. */
export function decodeNewMessagesCursor(value: string): NewMessagesCursor | null {
  if (!value) return null;

  let parsed: unknown;
  try {
    parsed = JSON.parse(Buffer.from(value, 'base64url').toString('utf8'));
  } catch {
    return null;
  }

  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return null;
  const raw = parsed as Record<string, unknown>;
  if (raw.version !== 1) return null;
  if (typeof raw.since !== 'number' || !Number.isFinite(raw.since) || raw.since < 0) return null;
  if (!raw.chats || typeof raw.chats !== 'object' || Array.isArray(raw.chats)) return null;

  const chats: Record<string, string> = {};
  for (const [conversationId, messageId] of Object.entries(raw.chats as Record<string, unknown>)) {
    if (typeof messageId !== 'string') return null;
    chats[conversationId] = messageId;
  }

  return { version: 1, since: raw.since, chats };
}

// ─────────────────────────────────────────────────────────────────────────────
// Pure helpers
// ─────────────────────────────────────────────────────────────────────────────

/** Truncates text to `max` characters, counting the ellipsis in the limit. */
export function truncateSnippet(text: string, max = NEW_MESSAGES_SNIPPET_MAX_LENGTH): string {
  if (max <= 0) return '';
  if (text.length <= max) return text;
  return `${text.slice(0, max - 1)}…`;
}

/** Parses a numeric chatsvc message ID, or null when it is not finite. */
function parseMessageId(id: string | undefined): number | null {
  if (!id) return null;
  const parsed = Number(id);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
}

/** Highest numeric message ID in a page, including own and system messages. */
export function maxMessageId(messages: readonly RawChatsvcMessage[]): number {
  let max = 0;
  for (const message of messages) {
    const id = parseMessageId(message.id);
    if (id !== null) max = Math.max(max, id);
  }
  return max;
}

/** True for messages that should appear in the feed (not control/system/deleted). */
export function isFeedMessage(message: RawChatsvcMessage): boolean {
  const type = message.messagetype;
  if (!type) return false;
  if (type.startsWith('Control/') || type.startsWith('ThreadActivity/') || type.startsWith('Event/')) {
    return false;
  }
  return !message.properties?.deletetime;
}

/** Derives an ISO timestamp from the message fields, falling back to its numeric ID. */
function messageTimestamp(message: RawChatsvcMessage, id: string): string {
  for (const value of [message.originalarrivaltime, message.composetime]) {
    if (!value) continue;
    const parsed = new Date(value).getTime();
    if (!Number.isNaN(parsed)) return new Date(parsed).toISOString();
  }
  const fromId = parseMessageId(id);
  return fromId !== null ? new Date(fromId).toISOString() : new Date().toISOString();
}

/** Converts a raw chatsvc message into a compact feed item, or null if it is not a message. */
export function parseNewMessage(
  message: RawChatsvcMessage,
  context: { conversationId: string; topic: string; chatType: ChatType }
): NewMessageItem | null {
  const id = message.id;
  if (!id || !isFeedMessage(message)) return null;

  return {
    conversationId: context.conversationId,
    topic: context.topic,
    chatType: context.chatType,
    messageId: id,
    senderName: message.imdisplayname || message.displayName || '',
    sentAt: messageTimestamp(message, id),
    snippet: truncateSnippet(stripHtml(message.content || '')),
  };
}

/**
 * Picks the most recently active conversations that may hold something newer
 * than the cursor. A conversation is skipped when its last message is at or
 * before either the poll floor or its own high-water mark.
 */
export function selectConversationsToScan(
  conversations: readonly ChatSummary[],
  cursor: NewMessagesCursor | null,
  floorMs: number,
  maxChats: number
): { conversations: ChatSummary[]; truncated: boolean } {
  const candidates = conversations
    .map(conversation => ({
      conversation,
      lastMessageTime: new Date(conversation.lastMessageTime).getTime(),
    }))
    .filter(({ conversation, lastMessageTime }) => {
      if (Number.isNaN(lastMessageTime) || lastMessageTime <= floorMs) return false;
      const watermark = cursor?.chats[conversation.conversationId];
      if (!watermark) return true;
      // Message IDs are epoch milliseconds, so the watermark doubles as the
      // arrival time of the newest message already returned for this chat.
      const watermarkMs = parseMessageId(watermark);
      return watermarkMs === null || lastMessageTime > watermarkMs;
    })
    .sort((a, b) => b.lastMessageTime - a.lastMessageTime);

  const limit = Math.max(1, Math.floor(maxChats));
  return {
    conversations: candidates.slice(0, limit).map(candidate => candidate.conversation),
    truncated: candidates.length > limit,
  };
}

/** Orders feed items oldest first, breaking timestamp ties by numeric message ID. */
export function orderNewMessages(items: NewMessageItem[]): NewMessageItem[] {
  return [...items].sort((a, b) => {
    const byTime = new Date(a.sentAt).getTime() - new Date(b.sentAt).getTime();
    if (byTime !== 0) return byTime;
    return (parseMessageId(a.messageId) ?? 0) - (parseMessageId(b.messageId) ?? 0);
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Orchestration
// ─────────────────────────────────────────────────────────────────────────────

/** Runs an async worker over items with a bounded number in flight. */
async function forEachWithConcurrency<T>(
  items: readonly T[],
  limit: number,
  worker: (item: T) => Promise<void>
): Promise<void> {
  let next = 0;
  const runners = Array.from({ length: Math.min(Math.max(1, limit), items.length) }, async () => {
    while (next < items.length) {
      const item = items[next++];
      await worker(item);
    }
  });
  await Promise.all(runners);
}

/** True when a `from` value refers to the signed-in user's MRI. */
function isFromUser(from: string | undefined, userMri: string): boolean {
  if (!from) return false;
  return from === userMri || from.includes(userMri);
}

/**
 * Returns messages newer than the cursor across the user's conversations.
 *
 * Cheap by design: an idle poll costs one conversations-list request because
 * conversations whose last message the cursor already covers are not fetched.
 * Per-conversation watermarks make truncation and per-chat errors resumable:
 * messages beyond `maxChats` or `limitPerChat` are returned on the next call
 * instead of being skipped.
 */
export async function getNewMessages(
  options: GetNewMessagesOptions = {}
): Promise<Result<GetNewMessagesResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;

  const now = options.now ?? Date.now();
  const maxChats = Math.min(
    Math.max(1, Math.floor(options.maxChats ?? DEFAULT_NEW_MESSAGES_MAX_CHATS)),
    MAX_THREAD_LIMIT
  );
  const limitPerChat = Math.min(
    Math.max(1, Math.floor(options.limitPerChat ?? DEFAULT_NEW_MESSAGES_PER_CHAT)),
    MAX_NEW_MESSAGES_PER_CHAT
  );
  const includeOwn = options.includeOwn ?? false;

  let cursor: NewMessagesCursor | null = null;
  if (options.after !== undefined) {
    cursor = decodeNewMessagesCursor(options.after);
    if (!cursor) {
      return err(createError(
        ErrorCode.INVALID_INPUT,
        'Invalid after cursor. Omit after on the first call, then pass nextAfter unchanged from the previous response.',
        { retryable: false }
      ));
    }
  }

  // With no cursor there is no previous poll to resume from: fetch each
  // conversation's most recent page (bounded by limitPerChat) and take it from
  // there. A floor of 0 also drains conversations beyond maxChats across later
  // polls instead of dropping them.
  const floorMs = cursor ? cursor.since : 0;

  const conversationsResult = await getConversations({ limit: MAX_THREAD_LIMIT });
  if (!conversationsResult.ok) {
    return conversationsResult;
  }

  const { conversations, truncated: conversationsTruncated } = selectConversationsToScan(
    conversationsResult.value.conversations,
    cursor,
    floorMs,
    maxChats
  );

  const chats: Record<string, string> = { ...(cursor?.chats ?? {}) };
  const items: NewMessageItem[] = [];
  const errors: NewMessagesError[] = [];
  let messagesTruncated = false;

  await forEachWithConcurrency(conversations, NEW_MESSAGES_FETCH_CONCURRENCY, async conversation => {
    const conversationId = conversation.conversationId;
    const watermark = chats[conversationId];
    const watermarkId = parseMessageId(watermark);
    // Known chats resume after their watermark. Conversations seen for the
    // first time on a resumed poll start at the poll floor so no messages
    // between the floor and the newest page are skipped. On a true first run
    // (floor 0) there is no start bound, so the newest page is returned.
    const startTime = watermarkId ?? (floorMs > 0 ? floorMs : undefined);

    const url = buildThreadMessagesUrl(
      CHATSVC_API.messages(region, conversationId, undefined, baseUrl),
      { limit: limitPerChat, startTime }
    );

    const response = await httpRequest<{ messages?: unknown[] }>(url, {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    });

    if (!response.ok) {
      errors.push({
        conversationId,
        code: response.error.code,
        message: response.error.message,
      });
      return;
    }

    const rawMessages = Array.isArray(response.value.data.messages)
      ? (response.value.data.messages as RawChatsvcMessage[])
      : [];

    // Advance even past own/system messages so an excluded own reply cannot
    // make the same conversation look new forever. Include the conversation's
    // last-message time as well: a channel whose newest activity is a thread
    // reply returns no top-level message, and would otherwise be refetched on
    // every poll. Never move a watermark backwards.
    const pageMax = maxMessageId(rawMessages);
    const conversationLastActivity = new Date(conversation.lastMessageTime).getTime();
    const existingWatermark = parseMessageId(chats[conversationId]);
    const nextWatermark = Math.max(
      existingWatermark ?? 0,
      pageMax,
      Number.isNaN(conversationLastActivity) ? 0 : conversationLastActivity
    );
    if (nextWatermark > 0) {
      chats[conversationId] = String(nextWatermark);
    }

    // A hit page means there may be more messages to resume. On a true first
    // run the page is the newest one, so older history is intentionally out of
    // scope and this is not truncation.
    const resuming = watermarkId !== null || floorMs > 0;
    if (resuming && rawMessages.length >= limitPerChat) {
      messagesTruncated = true;
    }

    const context = {
      conversationId,
      topic: conversation.topic,
      chatType: conversation.chatType,
    };

    for (const raw of rawMessages) {
      const id = parseMessageId(raw.id);
      if (id === null) continue;
      if (startTime !== undefined && id <= startTime) continue;
      if (!includeOwn && isFromUser(raw.from, auth.userMri)) continue;

      const item = parseNewMessage(raw, context);
      if (item) items.push(item);
    }
  });

  // Only advance the poll floor when every active conversation was either
  // scanned or already covered by its own watermark. Otherwise a chat skipped
  // this round could have its backlog silently excluded next round.
  const hasUncoveredFailure = errors.some(error => !(error.conversationId in chats));
  const nextSince = conversationsTruncated || hasUncoveredFailure
    ? floorMs
    : Math.max(0, now - NEW_MESSAGES_FLOOR_SKEW_MS);

  // Keep the cursor bounded to conversations that still exist. Watermarks for
  // chats outside the list are of no use (they cannot be scanned) and would
  // otherwise accumulate forever.
  const liveConversationIds = new Set(
    conversationsResult.value.conversations.map(conversation => conversation.conversationId)
  );
  const nextChats: Record<string, string> = {};
  for (const [conversationId, messageId] of Object.entries(chats)) {
    if (liveConversationIds.has(conversationId)) nextChats[conversationId] = messageId;
  }

  return ok({
    items: orderNewMessages(items),
    nextAfter: encodeNewMessagesCursor({ version: 1, since: nextSince, chats: nextChats }),
    truncated: conversationsTruncated || messagesTruncated,
    errors,
  });
}
