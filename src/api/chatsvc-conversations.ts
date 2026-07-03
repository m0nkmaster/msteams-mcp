/**
 * Chat Service API - Conversation list.
 *
 * Lists the user's recent conversations (1:1 chats, group chats, meetings and
 * channels) with a last-message preview, via the chatsvc conversations endpoint.
 */

import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { stripHtml } from '../utils/parsers.js';
import { DEFAULT_THREAD_LIMIT } from '../constants.js';

/** Maximum length of a last-message preview before truncation. */
const PREVIEW_MAX_LENGTH = 120;

/** A summarised conversation as returned by the chat list. */
export interface ChatSummary {
  conversationId: string;
  /** Channel/group topic, or empty for 1:1 chats. */
  topic: string;
  chatType: 'channel' | 'meeting' | 'group' | 'oneOnOne' | 'chat';
  /** ISO timestamp of the last message, or empty if unknown. */
  lastMessageTime: string;
  /** Display name of the last message's sender. */
  lastMessageFrom: string;
  /** Plain-text preview of the last message, truncated. */
  lastMessagePreview: string;
}

/** Result of listing conversations. */
export interface GetConversationsResult {
  conversations: ChatSummary[];
}

/** Classifies a conversation type from its ID. */
function classifyChatType(id: string): ChatSummary['chatType'] {
  if (id.includes('meeting_')) return 'meeting';
  if (id.includes('@thread.tacv2')) return 'channel';
  if (id.includes('@unq.gbl.spaces')) return 'oneOnOne';
  if (id.includes('@thread.v2')) return 'group';
  return 'chat';
}

/** Truncates a preview to {@link PREVIEW_MAX_LENGTH} characters, adding an ellipsis. */
function truncatePreview(text: string): string {
  if (text.length <= PREVIEW_MAX_LENGTH) return text;
  return `${text.slice(0, PREVIEW_MAX_LENGTH)}…`;
}

/**
 * Parses a raw chatsvc conversation into a {@link ChatSummary}, or null if it
 * has no usable ID. Pure — no IO — so the field mapping and truncation are
 * unit-testable against fixture JSON.
 */
export function parseConversation(raw: Record<string, unknown>): ChatSummary | null {
  const id = typeof raw.id === 'string' ? raw.id : '';
  if (!id) return null;

  const threadProperties = (raw.threadProperties ?? {}) as Record<string, unknown>;
  const topic = typeof threadProperties.topic === 'string' ? threadProperties.topic : '';

  const lastMessage = (raw.lastMessage ?? {}) as Record<string, unknown>;
  const lastMessageTime =
    (typeof lastMessage.originalarrivaltime === 'string' && lastMessage.originalarrivaltime) ||
    (typeof lastMessage.composetime === 'string' && lastMessage.composetime) ||
    '';
  const lastMessageFrom =
    typeof lastMessage.imdisplayname === 'string' ? lastMessage.imdisplayname : '';
  const rawContent = typeof lastMessage.content === 'string' ? lastMessage.content : '';

  return {
    conversationId: id,
    topic,
    chatType: classifyChatType(id),
    lastMessageTime,
    lastMessageFrom,
    lastMessagePreview: truncatePreview(stripHtml(rawContent)),
  };
}

/**
 * Lists the user's recent conversations (chats, group chats, meetings and
 * channels), newest activity first, capped at `limit`.
 */
export async function getConversations(
  options: { limit?: number } = {}
): Promise<Result<GetConversationsResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const limit = options.limit ?? DEFAULT_THREAD_LIMIT;

  const response = await httpRequest<{ conversations?: unknown[] }>(
    CHATSVC_API.conversations(region, baseUrl),
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const raw = Array.isArray(response.value.data.conversations)
    ? response.value.data.conversations
    : [];

  const conversations = raw
    .map(c => parseConversation(c as Record<string, unknown>))
    .filter((c): c is ChatSummary => c !== null)
    .slice(0, limit);

  return ok({ conversations });
}
