/**
 * Chat Service API - Messaging operations.
 * 
 * Send, edit, delete messages. Get thread messages. 1:1 and group chat creation.
 * Conversation properties and participant extraction.
 */

import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getMessagingHeaders, getSkypeAuthHeaders } from '../utils/api-config.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import { getUserDisplayName } from '../auth/token-extractor.js';
import { requireMessageAuth, requireMessageAuthWithConfig, getTeamsBaseUrl, getTenantId } from '../utils/auth-guards.js';
import { stripHtml, extractLinks, buildMessageLink, buildOneOnOneConversationId, extractObjectId, markdownToTeamsHtml, hasMarkdownFormatting, escapeHtmlChars, sanitizeLinkUrl, parseReactions, type ExtractedLink, type Reaction, type ReactionSummary } from '../utils/parsers.js';
import { resolveNames } from './profile-api.js';
import { SELF_CHAT_ID, MRI_ORGID_PREFIX, DEFAULT_WAIT_SECONDS, DEFAULT_WAIT_INTERVAL_SECONDS, MAX_WAIT_SECONDS } from '../constants.js';
import { formatHumanReadableDate } from './chatsvc-common.js';
import type { RawChatsvcMessage, RawConversationResponse, RawCreateThreadResponse } from '../types/api-responses.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Result of sending a message. */
export interface SendMessageResult {
  messageId: string;
  timestamp?: number;
  /** For scheduled messages: the ISO 8601 UTC time the message will be sent. */
  scheduledFor?: string;
}

/** A message from a thread/conversation. */
export interface ThreadMessage {
  id: string;
  content: string;
  contentType: string;
  sender: {
    mri: string;
    displayName?: string;
  };
  timestamp: string;
  /** Human-readable date with day of week, e.g., "Friday, January 30, 2026, 10:45 AM UTC" */
  when?: string;
  conversationId: string;
  clientMessageId?: string;
  isFromMe?: boolean;
  messageLink?: string;
  links?: ExtractedLink[];
  /** For channel messages: the ID of the thread root post (if this is a reply within a thread) */
  threadRootId?: string;
  /** True if this message is a reply within a channel thread (not a top-level post) */
  isThreadReply?: boolean;
  /** Emoji reactions on this message, with reactor identity (names resolved where possible). */
  reactions?: Reaction[];
  /** Reaction counts keyed by emoji (e.g. { like: 3, heart: 1 }). */
  reactionSummary?: ReactionSummary;
}

/** Result of getting thread messages. */
export interface GetThreadResult {
  conversationId: string;
  messages: ThreadMessage[];
}

/** Result of editing a message. */
export interface EditMessageResult {
  messageId: string;
  conversationId: string;
}

/** Result of deleting a message. */
export interface DeleteMessageResult {
  messageId: string;
  conversationId: string;
}

/** A mention to include in a message. */
export interface Mention {
  /** The user's MRI (e.g., '8:orgid:uuid') or tag MRI (e.g., 'tag:abc123'). */
  mri: string;
  /** Display name to show for the mention. */
  displayName: string;
}

/** Returns true if the MRI refers to a channel tag (e.g., 'tag:abc123'). */
function isTagMention(mri: string): boolean {
  return mri.startsWith('tag:');
}

/** Returns the actual MRI to send in the API payload (strips 'tag:' prefix for tags). */
function getActualMri(mri: string): string {
  return isTagMention(mri) ? mri.substring(4) : mri;
}

/** Options for sending a message. */
export interface SendMessageOptions {
  /**
   * Message ID of the thread root to reply to.
   * 
   * When provided, the message is posted as a reply to an existing thread
   * in a channel. The conversationId should be the channel ID, and this
   * should be the ID of the first message in the thread.
   * 
   * For chats (1:1, group, meeting), this is not needed - all messages
   * are part of the same flat conversation.
   */
  replyToMessageId?: string;
  /**
   * How to interpret the content. Defaults to `markdown` (convert markdown +
   * @mentions + links to Teams HTML — the historical behaviour). Use `text` to
   * send verbatim with no reinterpretation, `html` for caller-supplied Teams
   * HTML, or `auto` to convert only when markdown/mention syntax is present.
   */
  contentType?: ContentType;
  /**
   * For channel posts: a thread subject/title. Starts a new titled thread.
   * Ignored for 1:1/group chats (they have no per-message subject).
   */
  subject?: string;
  /**
   * Schedule the message for future delivery. ISO 8601 (e.g.
   * "2026-04-11T09:00:00Z"); a timezone-less value is treated as UTC. Cannot be
   * combined with mentions or a thread reply.
   */
  scheduleAt?: string;
}

/** Result of getting a 1:1 conversation. */
export interface GetOneOnOneChatResult {
  conversationId: string;
  otherUserId: string;
  currentUserId: string;
}

/** Result of creating a group chat. */
export interface CreateGroupChatResult {
  /** The newly created conversation ID. */
  conversationId: string;
  /** The member MRIs included in the chat. */
  members: string[];
  /** Optional topic/name set for the chat. */
  topic?: string;
  /** Optional note about the result, e.g. if the ID could not be retrieved. */
  note?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Message Sending
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Generates a Teams client message ID.
 *
 * Teams requires this to be "a number in string format" — a UUID is rejected
 * with `StoreInvalidInput - ClientMessageId must be a number in string format`.
 * This mirrors the Teams web client's large random integer: current time in
 * milliseconds plus 6 random digits, keeping IDs numeric and unique even for
 * rapid successive sends.
 */
export function generateClientMessageId(): string {
  const random = Math.floor(Math.random() * 1_000_000).toString().padStart(6, '0');
  return `${Date.now()}${random}`;
}

/**
 * Sends a message to a Teams conversation.
 * 
 * For channels, you can either:
 * - Post a new top-level message: just provide the channel's conversationId
 * - Reply to a thread: provide the channel's conversationId AND replyToMessageId
 * 
 * For chats (1:1, group, meeting), all messages go to the same conversation
 * without threading - just provide the conversationId.
 */
export async function sendMessage(
  conversationId: string,
  content: string,
  options: SendMessageOptions = {}
): Promise<Result<SendMessageResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;

  const { replyToMessageId, contentType = 'markdown', subject, scheduleAt } = options;
  const displayName = getUserDisplayName() || 'User';

  const clientMessageId = generateClientMessageId();

  // Resolve content per the requested content type (markdown by default).
  const resolved = resolveMessageContent(content, contentType);

  // Scheduled send goes through the chatsvc drafts API (same as the Teams web
  // client). Drafts do not support mentions or thread replies.
  if (scheduleAt) {
    if (resolved.mentions.length > 0 || replyToMessageId) {
      return err(createError(
        ErrorCode.INVALID_INPUT,
        'Scheduled messages cannot include @mentions or be thread replies. Remove scheduleAt, or the mention/replyToMessageId.'
      ));
    }
    const parsedTime = parseScheduleTime(scheduleAt);
    if (!parsedTime.ok) {
      return parsedTime;
    }
    const draftBody = buildScheduledDraftBody({
      conversationId,
      content: resolved.content,
      messagetype: resolved.messagetype,
      epochMs: parsedTime.value.epochMs,
      userMri: auth.userMri,
      isoNow: new Date().toISOString(),
      conversationLink: CHATSVC_API.conversation(region, conversationId, baseUrl),
      clientMessageId,
    });
    const draftResponse = await httpRequest<unknown>(
      CHATSVC_API.drafts(region, baseUrl),
      {
        method: 'POST',
        headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
        body: JSON.stringify(draftBody),
      }
    );
    if (!draftResponse.ok) {
      return draftResponse;
    }
    return ok({ messageId: clientMessageId, scheduledFor: parsedTime.value.iso });
  }

  // Build the message body
  const body: Record<string, unknown> = {
    content: resolved.content,
    messagetype: resolved.messagetype,
    contenttype: 'text',
    imdisplayname: displayName,
    clientmessageid: clientMessageId,
  };

  // Message properties: mentions and/or a channel thread subject.
  const properties: Record<string, unknown> = {};
  if (resolved.mentions.length > 0) {
    properties.mentions = buildMentionsProperty(resolved.mentions);
  }
  if (subject && subject.trim()) {
    properties.subject = subject.trim();
  }
  if (Object.keys(properties).length > 0) {
    body.properties = properties;
  }

  const url = CHATSVC_API.messages(region, conversationId, replyToMessageId, baseUrl);
  const requestInit = {
    method: 'POST' as const,
    headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
    body: JSON.stringify(body),
  };

  let response = await httpRequest<{ OriginalArrivalTime?: number }>(url, requestInit);

  // A 1:1 chat that has never been messaged has no chatsvc thread yet, so the
  // deterministic @unq.gbl.spaces id 404s. Create the unique-roster thread (what
  // the Teams client does when starting a new chat) and retry once.
  if (!response.ok && response.error.code === ErrorCode.NOT_FOUND) {
    const memberIds = extractOneOnOneMemberIds(conversationId);
    if (memberIds) {
      const created = await createOneOnOneThread(region, baseUrl, auth.skypeToken, auth.authToken, memberIds);
      if (created.ok) {
        response = await httpRequest<{ OriginalArrivalTime?: number }>(url, requestInit);
      }
    }
  }

  if (!response.ok) {
    return response;
  }

  return ok({
    messageId: clientMessageId,
    timestamp: response.value.data.OriginalArrivalTime,
  });
}

/**
 * Creates the unique-roster 1:1 thread for a `19:<a>_<b>@unq.gbl.spaces`
 * conversation so the first message to a never-used 1:1 chat can be delivered.
 */
async function createOneOnOneThread(
  region: string,
  baseUrl: string,
  skypeToken: string,
  authToken: string,
  memberIds: string[]
): Promise<Result<void>> {
  const response = await httpRequest<unknown>(
    CHATSVC_API.createThread(region, baseUrl),
    {
      method: 'POST',
      headers: {
        ...getMessagingHeaders(skypeToken, authToken, baseUrl),
        'Accept': 'application/json',
      },
      body: JSON.stringify(buildOneOnOneThreadBody(memberIds)),
    }
  );
  if (!response.ok) {
    return response;
  }
  return ok(undefined);
}

/**
 * Sends a message to your own notes/self-chat.
 */
export async function sendNoteToSelf(content: string): Promise<Result<SendMessageResult>> {
  return sendMessage(SELF_CHAT_ID, content);
}

// ─────────────────────────────────────────────────────────────────────────────
// Single Message
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Fetches a single message by ID from a conversation.
 * 
 * Works for messages of any age - no retention limit observed.
 * Useful for resolving saved message stubs, search result snippets,
 * or any case where you have a conversationId + messageId.
 */
export async function getMessage(
  conversationId: string,
  messageId: string
): Promise<Result<ThreadMessage>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;

  const url = CHATSVC_API.singleMessage(region, conversationId, messageId, baseUrl);

  const response = await httpRequest<RawChatsvcMessage>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const msg = response.value.data;

  const id = msg.id || msg.originalarrivaltime;
  if (!id) {
    return err(createError(
      ErrorCode.API_ERROR,
      'Message response missing ID'
    ));
  }

  const content = msg.content || '';
  const contentType = msg.messagetype || 'Text';
  const fromMri = msg.from || '';
  const displayName = msg.imdisplayname || msg.displayName;

  const timestamp = msg.originalarrivaltime ||
    msg.composetime ||
    timestampFromIdOrNow(id);

  const rootMessageId = msg.rootMessageId;
  const isThreadReply = !!rootMessageId && rootMessageId !== id;

  const messageLink = /^\d+$/.test(id)
    ? buildMessageLink({
        conversationId,
        messageId: id,
        tenantId: getTenantId() ?? undefined,
        parentMessageId: isThreadReply ? rootMessageId : undefined,
        teamsBaseUrl: getTeamsBaseUrl(),
      })
    : undefined;

  const links = extractLinks(content);
  const when = formatHumanReadableDate(timestamp);

  const { reactions, reactionSummary } = parseReactions(msg);

  // Enrich reactor names via fetchShortProfile batch API
  await enrichReactionNamesFromProfiles(reactions);

  return ok({
    id,
    content: stripHtml(content),
    contentType,
    sender: {
      mri: fromMri,
      displayName,
    },
    timestamp,
    when: when || undefined,
    conversationId,
    clientMessageId: msg.clientmessageid,
    isFromMe: fromMri === auth.userMri,
    messageLink,
    links: links.length > 0 ? links : undefined,
    threadRootId: isThreadReply ? rootMessageId : undefined,
    isThreadReply: isThreadReply || undefined,
    reactions,
    reactionSummary,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Thread Messages
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets messages from a Teams conversation/thread.
 * 
 * @param conversationId - The conversation ID to fetch messages from
 * @param options.limit - Maximum messages to return (default 50)
 * @param options.startTime - Fetch messages from this timestamp onwards
 * @param options.order - Sort order: 'desc' (newest-first, default) or 'asc' (oldest-first)
 * @param options.replyToMessageId - For channels: scope to replies of a specific top-level post
 */
export async function getThreadMessages(
  conversationId: string,
  options: { limit?: number; startTime?: number; order?: 'asc' | 'desc'; replyToMessageId?: string } = {}
): Promise<Result<GetThreadResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const limit = options.limit ?? 50;

  let url = CHATSVC_API.messages(region, conversationId, options.replyToMessageId, baseUrl);
  url += `?view=msnp24Equivalent|supportsMessageProperties&pageSize=${limit}`;

  if (options.startTime) {
    url += `&startTime=${options.startTime}`;
  }

  const response = await httpRequest<{ messages?: unknown[] }>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const rawMessages = response.value.data.messages;
  if (!Array.isArray(rawMessages)) {
    return ok({
      conversationId,
      messages: [],
    });
  }

  const messages: ThreadMessage[] = [];

  for (const raw of rawMessages) {
    const msg = raw as RawChatsvcMessage;

    // Skip non-message types
    const messageType = msg.messagetype;
    if (!messageType || messageType.startsWith('Control/') || messageType === 'ThreadActivity/AddMember') {
      continue;
    }

    // Skip deleted messages (they have empty content and a deletetime property)
    if (msg.properties?.deletetime) {
      continue;
    }

    const id = msg.id || msg.originalarrivaltime;
    if (!id) continue;

    const content = msg.content || '';
    const contentType = msg.messagetype || 'Text';

    const fromMri = msg.from || '';
    const displayName = msg.imdisplayname || msg.displayName;

    const timestamp = msg.originalarrivaltime ||
      msg.composetime ||
      timestampFromIdOrNow(id);

    // Extract thread root ID for channel messages
    // When rootMessageId differs from id, this message is a reply within a thread
    const rootMessageId = msg.rootMessageId;
    const isThreadReply = !!rootMessageId && rootMessageId !== id;

    // Build message link with tenant context for reliable deep links
    const messageLink = /^\d+$/.test(id)
      ? buildMessageLink({
          conversationId,
          messageId: id,
          tenantId: getTenantId() ?? undefined,
          parentMessageId: isThreadReply ? rootMessageId : undefined,
          teamsBaseUrl: getTeamsBaseUrl(),
        })
      : undefined;

    // Extract links before stripping HTML
    const links = extractLinks(content);

    // Format human-readable date with day of week to help LLMs
    const when = formatHumanReadableDate(timestamp);

    const { reactions, reactionSummary } = parseReactions(msg);

    messages.push({
      id,
      content: stripHtml(content),
      contentType,
      sender: {
        mri: fromMri,
        displayName,
      },
      timestamp,
      when: when || undefined,
      conversationId,
      clientMessageId: msg.clientmessageid,
      isFromMe: fromMri === auth.userMri,
      messageLink,
      links: links.length > 0 ? links : undefined,
      threadRootId: isThreadReply ? rootMessageId : undefined,
      isThreadReply: isThreadReply || undefined,
      reactions,
      reactionSummary,
    });
  }

  // Enrich reactor names via fetchShortProfile batch API.
  // Collects all unresolved reactor MRIs across all messages and resolves
  // them in a single batch call.
  const allReactions = messages.flatMap(m => m.reactions ?? []);
  await enrichReactionNamesFromProfiles(allReactions);

  // Sort by timestamp - default to newest-first (desc) for "what's latest" use cases
  const order = options.order ?? 'desc';
  if (order === 'asc') {
    // Oldest first (chronological reading order)
    messages.sort((a, b) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime());
  } else {
    // Newest first (default - latest activity at top)
    messages.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
  }

  return ok({
    conversationId,
    messages,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Edit / Delete
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Edits an existing message.
 *
 * Uses the same content pipeline as {@link sendMessage} (markdown, @[mentions], links).
 *
 * Note: You can only edit your own messages. The API will reject
 * attempts to edit messages from other users.
 *
 * @param conversationId - The conversation containing the message
 * @param messageId - The ID of the message to edit
 * @param newContent - New content (markdown by default; supports mentions like sendMessage)
 * @param contentType - How to interpret the content (default `markdown`)
 */
export async function editMessage(
  conversationId: string,
  messageId: string,
  newContent: string,
  contentType: ContentType = 'markdown'
): Promise<Result<EditMessageResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const displayName = getUserDisplayName() || 'User';

  // Same pipeline as sendMessage: resolve content per the requested type.
  const resolved = resolveMessageContent(newContent, contentType);

  const body: Record<string, unknown> = {
    id: messageId,
    type: 'Message',
    conversationid: conversationId,
    content: resolved.content,
    messagetype: resolved.messagetype,
    contenttype: 'text',
    imdisplayname: displayName,
  };

  if (resolved.mentions.length > 0) {
    body.properties = {
      mentions: buildMentionsProperty(resolved.mentions),
    };
  }

  const url = CHATSVC_API.editMessage(region, conversationId, messageId, baseUrl);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'PUT',
      headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    messageId,
    conversationId,
  });
}

/**
 * Deletes a message (soft delete).
 * 
 * Note: You can only delete your own messages, unless you are a
 * channel owner/moderator. The API will reject unauthorised attempts.
 * 
 * @param conversationId - The conversation containing the message
 * @param messageId - The ID of the message to delete
 */
export async function deleteMessage(
  conversationId: string,
  messageId: string
): Promise<Result<DeleteMessageResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const url = CHATSVC_API.deleteMessage(region, conversationId, messageId, baseUrl);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'DELETE',
      headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    messageId,
    conversationId,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Conversation Properties
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets properties for a single conversation.
 */
export async function getConversationProperties(
  conversationId: string
): Promise<Result<{ displayName?: string; conversationType?: string }>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const url = CHATSVC_API.conversation(region, conversationId, baseUrl) + '?view=msnp24Equivalent';

  const response = await httpRequest<RawConversationResponse>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  const data = response.value.data;
  const threadProps = data.threadProperties;
  const productType = threadProps?.productThreadType;

  // Try to get display name from various sources
  let displayName: string | undefined;

  if (threadProps?.topicThreadTopic) {
    displayName = threadProps.topicThreadTopic as string;
  }

  if (!displayName && threadProps?.topic) {
    displayName = threadProps.topic as string;
  }

  if (!displayName && threadProps?.spaceThreadTopic) {
    displayName = threadProps.spaceThreadTopic as string;
  }

  if (!displayName && threadProps?.threadtopic) {
    displayName = threadProps.threadtopic as string;
  }

  // For chats without a topic: build from members
  if (!displayName) {
    const members = data.members as Array<Record<string, unknown>> | undefined;
    if (members && members.length > 0) {
      const otherMembers = members
        .filter(m => m.mri !== auth.userMri && m.id !== auth.userMri)
        .map(m => (m.friendlyName || m.displayName || m.name) as string | undefined)
        .filter((name): name is string => !!name);

      if (otherMembers.length > 0) {
        displayName = otherMembers.length <= 3
          ? otherMembers.join(', ')
          : `${otherMembers.slice(0, 3).join(', ')} + ${otherMembers.length - 3} more`;
      }
    }
  }

  // Determine conversation type
  let conversationType: string | undefined;

  if (productType) {
    if (productType === 'Meeting') {
      conversationType = 'Meeting';
    } else if (productType.includes('Channel') || productType === 'TeamsTeam') {
      conversationType = 'Channel';
    } else if (productType === 'Chat' || productType === 'OneOnOne') {
      conversationType = 'Chat';
    }
  }

  // Fallback to ID pattern detection
  if (!conversationType) {
    if (conversationId.includes('meeting_')) {
      conversationType = 'Meeting';
    } else if (threadProps?.groupId) {
      conversationType = 'Channel';
    } else if (conversationId.includes('@thread.tacv2') || conversationId.includes('@thread.v2')) {
      conversationType = 'Chat';
    } else if (conversationId.startsWith('8:')) {
      conversationType = 'Chat';
    }
  }

  return ok({ displayName, conversationType });
}

/**
 * Extracts unique participant names from recent messages.
 */
export async function extractParticipantNames(
  conversationId: string
): Promise<Result<string | undefined>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return ok(undefined); // Non-critical: just return undefined if not authenticated
  }
  const { auth, region, baseUrl } = authResult.value;
  let url = CHATSVC_API.messages(region, conversationId, undefined, baseUrl);
  url += '?view=msnp24Equivalent&pageSize=10';

  const response = await httpRequest<{ messages?: unknown[] }>(
    url,
    {
      method: 'GET',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
    }
  );

  if (!response.ok) {
    return ok(undefined);
  }

  const messages = response.value.data.messages;
  if (!messages || messages.length === 0) {
    return ok(undefined);
  }

  const senderNames = new Set<string>();
  for (const msg of messages) {
    const m = msg as RawChatsvcMessage;
    const fromMri = m.from || '';
    const displayName = m.imdisplayname;

    if (fromMri === auth.userMri || !displayName) {
      continue;
    }

    senderNames.add(displayName);
  }

  if (senderNames.size === 0) {
    return ok(undefined);
  }

  const names = Array.from(senderNames);
  const result = names.length <= 3
    ? names.join(', ')
    : `${names.slice(0, 3).join(', ')} + ${names.length - 3} more`;

  return ok(result);
}

// ─────────────────────────────────────────────────────────────────────────────
// 1:1 Chat
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets the conversation ID for a 1:1 chat with another user.
 * 
 * Constructs the predictable format: `19:{id1}_{id2}@unq.gbl.spaces`
 * where IDs are sorted lexicographically. The conversation is created
 * implicitly when the first message is sent.
 */
export function getOneOnOneChatId(
  otherUserIdentifier: string
): Result<GetOneOnOneChatResult> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  // Extract the current user's object ID from their MRI
  const currentUserId = extractObjectId(auth.userMri);
  if (!currentUserId) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not extract user ID from session. Please try logging in again.'
    ));
  }

  // Extract the other user's object ID
  const otherUserId = extractObjectId(otherUserIdentifier);
  if (!otherUserId) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      `Invalid user identifier: ${otherUserIdentifier}. Expected MRI (8:orgid:guid), ID with tenant (guid@tenant), or raw GUID.`
    ));
  }

  const conversationId = buildOneOnOneConversationId(currentUserId, otherUserId);

  if (!conversationId) {
    // This shouldn't happen if both IDs were validated above, but handle it anyway
    return err(createError(
      ErrorCode.UNKNOWN,
      'Failed to construct conversation ID.'
    ));
  }

  return ok({
    conversationId,
    otherUserId,
    currentUserId,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Group Chat
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Creates a new group chat with multiple members.
 * 
 * Unlike 1:1 chats which have a predictable ID format, group chats require
 * an API call to create. The conversation ID is returned by the server.
 * 
 * @param memberIdentifiers - Array of user identifiers (MRI, object ID, or GUID)
 * @param topic - Optional chat topic/name
 * 
 * @example
 * ```typescript
 * // Create a group chat with 2 other people
 * const result = await createGroupChat(
 *   ['8:orgid:abc123...', '8:orgid:def456...'],
 *   'Project Discussion'
 * );
 * if (result.ok) {
 *   console.log('Created chat:', result.value.conversationId);
 * }
 * ```
 */
export async function createGroupChat(
  memberIdentifiers: string[],
  topic?: string
): Promise<Result<CreateGroupChatResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;

  // Validate we have at least 2 other members (plus current user = 3+ total)
  if (!memberIdentifiers || memberIdentifiers.length < 2) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      'Group chat requires at least 2 other members. For 1:1 chats, use teams_get_chat instead.'
    ));
  }

  // Check for duplicate members
  const uniqueIdentifiers = new Set(memberIdentifiers);
  if (uniqueIdentifiers.size !== memberIdentifiers.length) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      'Duplicate members detected in group chat request.'
    ));
  }

  // Teams group chats have a 250-member limit
  if (memberIdentifiers.length > 250) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      'Group chat cannot have more than 250 members.'
    ));
  }

  // Build MRI list for all members, including current user
  const memberMris: string[] = [auth.userMri];

  for (const identifier of memberIdentifiers) {
    // Extract object ID and convert to MRI format
    const objectId = extractObjectId(identifier);
    if (!objectId) {
      return err(createError(
        ErrorCode.INVALID_INPUT,
        `Invalid user identifier: ${identifier}. Expected MRI (8:orgid:guid), ID with tenant (guid@tenant), or raw GUID.`
      ));
    }
    
    // Convert to MRI format if not already
    const mri = identifier.startsWith(MRI_ORGID_PREFIX) 
      ? identifier 
      : `${MRI_ORGID_PREFIX}${objectId}`;
    memberMris.push(mri);
  }

  // Build the request body
  // Format discovered via API research: POST /threads with members having "Admin" role
  const body: Record<string, unknown> = {
    members: memberMris.map(mri => ({
      id: mri,
      role: 'Admin',
    })),
    properties: {
      threadType: 'chat',
    },
  };

  // Add topic if provided
  if (topic) {
    (body.properties as Record<string, unknown>).topic = topic;
  }

  const url = CHATSVC_API.createThread(region, baseUrl);

  const response = await httpRequest<RawCreateThreadResponse>(
    url,
    {
      method: 'POST',
      headers: {
        ...getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
        'Accept': 'application/json',
      },
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return response;
  }

  // The response returns threadResource.id with the new conversation ID
  // Note: Sometimes the API returns an empty body {} but includes the ID in the Location header
  const responseData = response.value.data;
  let conversationId = responseData.threadResource?.id
    || responseData.id
    || responseData.threadId;
  
  // Fallback: extract from Location header (format: .../threads/19:xxx@thread.v2)
  if (!conversationId) {
    const locationHeader = response.value.headers.get('location');
    if (locationHeader) {
      const match = locationHeader.match(/threads\/(19:[^/]+)/);
      if (match) {
        conversationId = match[1];
      }
    }
  }
  
  if (!conversationId) {
    // Chat was likely created (201 status) but we couldn't find the ID
    return ok({
      conversationId: '(created - check Teams for conversation ID)',
      members: memberMris,
      topic,
      note: 'Group chat created successfully but API did not return the conversation ID. Check Teams to find the new chat.',
    });
  }

  return ok({
    conversationId,
    members: memberMris,
    topic,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Internal Helpers (mention/link parsing)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Derives an ISO timestamp from a numeric message ID, or returns the current time.
 */
function timestampFromIdOrNow(id: string): string {
  const parsed = parseInt(id, 10);
  return !isNaN(parsed) && parsed > 0 ? new Date(parsed).toISOString() : new Date().toISOString();
}

/**
 * Builds the HTML for a single mention.
 * 
 * Tag mentions use a span-only format; person mentions include a readonly wrapper.
 */
export function buildMentionHtml(displayName: string, itemId: number, mri: string): string {
  const spanHtml = `<span itemtype="http://schema.skype.com/Mention" itemscope itemid="${itemId}">${escapeHtmlChars(displayName)}</span>`;
  if (isTagMention(mri)) {
    return spanHtml;
  }
  return `<readonly class="skipProofing" itemtype="http://schema.skype.com/Mention" contenteditable="false" spellcheck="false">${spanHtml}</readonly>`;
}

/**
 * Builds the mentions property array for the API request.
 * 
 * Tag mentions use mentionType 'tag' with the raw tag ID (without 'tag:' prefix).
 */
export function buildMentionsProperty(mentions: Mention[]): string {
  const mentionObjects = mentions.map((mention, index) => ({
    '@type': 'http://schema.skype.com/Mention',
    'itemid': String(index),
    'mri': getActualMri(mention.mri),
    'mentionType': isTagMention(mention.mri) ? 'tag' : 'person',
    'displayName': mention.displayName,
  }));
  return JSON.stringify(mentionObjects);
}

/**
 * Parses content for both mentions @[Name](mri) and links [text](url).
 * Processes them in a single pass to avoid escaping conflicts.
 */
export function parseContentWithMentionsAndLinks(content: string): { html: string; mentions: Mention[] } {
  // Patterns for mentions and links
  // Note: Link pattern uses [^)\s] to reject URLs with spaces
  const mentionPattern = /@\[([^\]]+)\]\(([^)]+)\)/g;
  const linkPattern = /(?<!@)\[([^\]]+)\]\((https?:\/\/[^)\s]+|mailto:[^)\s]+)\)/g;
  
  // Match type for tracking positions
  interface Match {
    index: number;
    length: number;
    type: 'mention' | 'link';
    text: string;
    target: string; // mri for mentions, url for links
  }
  
  // Helper to find all matches for a pattern
  const findAll = (pattern: RegExp, type: 'mention' | 'link'): Match[] => {
    const results: Match[] = [];
    let match;
    while ((match = pattern.exec(content)) !== null) {
      results.push({
        index: match.index,
        length: match[0].length,
        type,
        text: match[1],
        target: match[2],
      });
    }
    return results;
  };
  
  // Find all mentions and links, then sort by position
  const matches = [
    ...findAll(mentionPattern, 'mention'),
    ...findAll(linkPattern, 'link'),
  ].sort((a, b) => a.index - b.index);
  
  // No mentions or links - use full markdown conversion
  if (matches.length === 0) {
    return { html: markdownToTeamsHtml(content), mentions: [] };
  }
  
  // Strategy: replace mentions/links with unique placeholders, run the whole
  // content through markdownToTeamsHtml (so links stay inline within their
  // paragraph), then substitute placeholders back with actual HTML.
  const mentions: Mention[] = [];
  const placeholders: Map<string, string> = new Map();
  let mentionId = 0;
  
  // Build content with placeholders (process in reverse to preserve indices).
  // Uses Unicode Private Use Area characters (U+E000, U+E001) as delimiters —
  // these never appear in normal text, so they're safe as unique markers.
  let placeholderContent = content;
  for (let i = matches.length - 1; i >= 0; i--) {
    const m = matches[i];
    const placeholder = `\uE000MCP_PH_${i}\uE001`;
    
    let html: string;
    if (m.type === 'mention') {
      mentions.push({ mri: m.target, displayName: m.text });
      html = buildMentionHtml(m.text, mentionId, m.target);
      mentionId++;
    } else {
      const safeText = escapeHtmlChars(m.text);
      const safeUrl = sanitizeLinkUrl(m.target).replace(/"/g, '&quot;');
      html = `<a href="${safeUrl}">${safeText}</a>`;
    }
    placeholders.set(placeholder, html);
    
    placeholderContent = placeholderContent.substring(0, m.index) + placeholder + placeholderContent.substring(m.index + m.length);
  }
  
  // Reverse mentions array since we processed in reverse order
  mentions.reverse();
  
  // Convert the whole content (with placeholders) through markdown pipeline
  let result = markdownToTeamsHtml(placeholderContent);
  
  // Substitute placeholders back with actual HTML
  for (const [placeholder, html] of placeholders) {
    result = result.replace(placeholder, html);
  }
  
  return { html: result, mentions };
}

/** How message content should be interpreted before sending. */
export type ContentType = 'auto' | 'text' | 'html' | 'markdown';

/** Content resolved into the shape the chatsvc/Graph message body expects. */
export interface ResolvedContent {
  /** The body content to send. */
  content: string;
  /** The wire message type: plain `Text` or `RichText/Html`. */
  messagetype: 'Text' | 'RichText/Html';
  /** Any @mentions extracted from the content (empty for text/html modes). */
  mentions: Mention[];
}

/** Detects whether content contains @[mention](mri) or [text](url) markdown syntax. */
function hasMentionOrLink(content: string): boolean {
  return (
    /@\[[^\]]+\]\([^)]+\)/.test(content) ||
    /(?<!@)\[[^\]]+\]\((?:https?:\/\/|mailto:)[^)\s]+\)/.test(content)
  );
}

/**
 * Resolves outbound message content according to the requested content type.
 *
 * - `markdown` (default): convert markdown + @mentions + links to Teams HTML.
 *   This preserves the historical behaviour where all content was converted.
 * - `text`: send verbatim as plain text — no markdown, mention or link
 *   processing. The escape hatch for content that should not be reinterpreted
 *   (e.g. `5*3*2`, `a_b_c`).
 * - `html`: caller-supplied Teams HTML, sent as `RichText/Html` unchanged.
 * - `auto`: convert only when the content actually contains markdown formatting
 *   or mention/link syntax; otherwise send plain text.
 */
export function resolveMessageContent(
  content: string,
  contentType: ContentType = 'markdown'
): ResolvedContent {
  switch (contentType) {
    case 'text':
      return { content, messagetype: 'Text', mentions: [] };
    case 'html':
      return { content, messagetype: 'RichText/Html', mentions: [] };
    case 'auto':
      if (!hasMarkdownFormatting(content) && !hasMentionOrLink(content)) {
        return { content, messagetype: 'Text', mentions: [] };
      }
    // falls through to markdown conversion when formatting is detected
    // eslint-disable-next-line no-fallthrough
    case 'markdown':
    default: {
      const parsed = parseContentWithMentionsAndLinks(content);
      return { content: parsed.html, messagetype: 'RichText/Html', mentions: parsed.mentions };
    }
  }
}

/**
 * Parses and validates a schedule datetime for a scheduled message.
 *
 * Accepts ISO 8601 (with or without a timezone) and a space-separated
 * `YYYY-MM-DD HH:mm[:ss]` form. A timezone-less value is treated as UTC (so
 * results are deterministic regardless of the host clock). Rejects unparseable
 * or past times. `now` is injectable for testing.
 */
export function parseScheduleTime(
  input: string,
  now: number = Date.now()
): Result<{ iso: string; epochMs: number }> {
  const trimmed = input.trim();
  if (!trimmed) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      'Schedule time cannot be empty. Use ISO 8601, e.g. "2026-04-11T09:00:00Z".'
    ));
  }
  // Allow a space separator and treat a timezone-less datetime as UTC.
  let normalized = trimmed.replace(' ', 'T');
  const hasTimezone = /(?:Z|[+-]\d{2}:?\d{2})$/.test(normalized);
  if (!hasTimezone) normalized += 'Z';

  const epochMs = Date.parse(normalized);
  if (Number.isNaN(epochMs)) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      `Invalid schedule time: "${input}". Use ISO 8601, e.g. "2026-04-11T09:00:00Z".`
    ));
  }
  if (epochMs <= now) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      `Schedule time "${input}" is in the past.`
    ));
  }
  return ok({ iso: new Date(epochMs).toISOString(), epochMs });
}

/**
 * Builds the chatsvc `/drafts` request body for a scheduled message — the same
 * mechanism the Teams web client uses. `sendAt` must be epoch milliseconds.
 */
export function buildScheduledDraftBody(opts: {
  conversationId: string;
  content: string;
  messagetype: 'Text' | 'RichText/Html';
  epochMs: number;
  userMri: string;
  isoNow: string;
  conversationLink: string;
  clientMessageId: string;
}): Record<string, unknown> {
  return {
    draftDetails: { sendAt: opts.epochMs },
    draftType: 'ScheduledDraft',
    innerThreadId: opts.conversationId,
    message: {
      id: '-1',
      type: 'Message',
      conversationid: opts.conversationId,
      conversationLink: opts.conversationLink,
      from: opts.userMri,
      fromUserId: opts.userMri,
      composetime: opts.isoNow,
      originalarrivaltime: opts.isoNow,
      content: opts.content,
      messagetype: opts.messagetype,
      contenttype: 'Text',
      imdisplayname: '',
      clientmessageid: opts.clientMessageId,
      properties: {
        importance: '',
        subject: '',
        title: '',
        cards: '[]',
        links: '[]',
        mentions: '[]',
        onbehalfof: null,
        files: '[]',
        policyViolation: null,
        draftId: opts.clientMessageId,
      },
      draftDetails: { sendAt: opts.epochMs },
      threadtype: 'streamofdrafts',
      innerThreadId: opts.conversationId,
    },
  };
}

/**
 * Extracts the two member object IDs from a 1:1 conversation ID of the form
 * `19:<a>_<b>@unq.gbl.spaces`. Returns null for any other conversation type or
 * a malformed ID (so the caller only attempts thread creation for real 1:1s).
 */
export function extractOneOnOneMemberIds(conversationId: string): string[] | null {
  const match = /^19:(.+)@unq\.gbl\.spaces$/.exec(conversationId);
  if (!match) return null;
  const parts = match[1].split('_');
  if (parts.length !== 2 || parts.some(p => p.length === 0)) return null;
  return parts;
}

/**
 * Builds the `/v1/threads` body that creates the unique-roster 1:1 thread — what
 * the Teams client does the first time you message someone. Both members are
 * added as Admins with the fixed unique-roster properties.
 */
export function buildOneOnOneThreadBody(memberIds: string[]): Record<string, unknown> {
  return {
    members: memberIds.map(id => ({ id: `${MRI_ORGID_PREFIX}${id}`, role: 'Admin' })),
    properties: {
      threadType: 'chat',
      fixedRoster: 'true',
      uniquerosterthread: 'true',
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Wait for reply
// ─────────────────────────────────────────────────────────────────────────────

/** Result of waiting for a new message in a conversation. */
export interface WaitForReplyResult {
  /** True if the wait timed out before any new message arrived. */
  timedOut: boolean;
  /** The baseline message ID the wait started from. */
  after: string;
  /** Resume cursor: pass as `after` to a follow-up wait to continue from here. */
  nextAfter: string;
  /** The new messages that arrived (oldest first), or empty on timeout. */
  newMessages: ThreadMessage[];
  /** How many poll iterations ran. */
  polls: number;
  /** How many seconds the wait blocked for. */
  waitedSeconds: number;
}

/** Options for {@link waitForReply}. */
export interface WaitForReplyOptions {
  /** Only return messages newer than this message ID. Defaults to the caller's most recent message. */
  after?: string;
  /** Maximum seconds to block (clamped to MAX_WAIT_SECONDS). */
  maxWaitSeconds?: number;
  /** Seconds between polls. */
  intervalSeconds?: number;
  /** Include the caller's own messages in the result (default false). */
  includeSelf?: boolean;
  /** Messages fetched per poll (default 50). */
  limit?: number;
}

/** Clamps the wait timing knobs into their sane range (pure, for testability). */
export function clampWaitParams(
  maxWaitSeconds: number,
  intervalSeconds: number
): { maxWait: number; interval: number } {
  const maxWait = Math.min(Math.max(Math.floor(maxWaitSeconds), 1), MAX_WAIT_SECONDS);
  const interval = Math.min(Math.max(Math.floor(intervalSeconds), 1), maxWait);
  return { maxWait, interval };
}

/** Parses a numeric Teams message ID, or null if it is not a positive integer string. */
function parseMessageId(id: string): number | null {
  const n = Number(id);
  return Number.isFinite(n) ? n : null;
}

/**
 * Selects the messages strictly newer than `afterId`, oldest first. Excludes the
 * caller's own messages unless `includeSelf` is set; ignores non-numeric IDs.
 * Pure so the filter/sort semantics are unit-testable.
 */
export function selectNewMessages<T extends { id: string; isFromMe?: boolean }>(
  messages: readonly T[],
  afterId: number,
  includeSelf: boolean
): T[] {
  return messages
    .filter(m => {
      const id = parseMessageId(m.id);
      if (id === null || id <= afterId) return false;
      if (!includeSelf && m.isFromMe) return false;
      return true;
    })
    .sort((a, b) => (parseMessageId(a.id) ?? 0) - (parseMessageId(b.id) ?? 0));
}

/**
 * Resolves the baseline message ID for a wait with no explicit `after`: the
 * caller's own most recent message, else the most recent message overall, else
 * 0 for an empty thread. Anchoring to the caller's own last message makes a wait
 * restarted after a send idempotent without remembering the sent ID. Pure.
 */
export function resolveWaitBaseline(
  messages: readonly { id: string; isFromMe?: boolean }[]
): number {
  let ownMax = 0;
  let anyMax = 0;
  for (const m of messages) {
    const id = parseMessageId(m.id);
    if (id === null) continue;
    anyMax = Math.max(anyMax, id);
    if (m.isFromMe) ownMax = Math.max(ownMax, id);
  }
  return ownMax > 0 ? ownMax : anyMax;
}

const sleep = (ms: number): Promise<void> => new Promise(resolve => setTimeout(resolve, ms));

/**
 * Blocks until a new message appears in a conversation, then returns only the
 * new messages. Polls inside the call so a waiting agent makes a single tool
 * call instead of a poll loop. Read-only and idempotent: it never posts, and the
 * `after` baseline makes a restart resume from the same point without replaying.
 *
 * The block is capped at MAX_WAIT_SECONDS so an MCP client timeout never kills
 * the call; on timeout it returns `timedOut: true` with `nextAfter` so the
 * caller can simply call again to keep waiting.
 */
export async function waitForReply(
  conversationId: string,
  options: WaitForReplyOptions = {}
): Promise<Result<WaitForReplyResult>> {
  const { after, includeSelf = false, limit = 50 } = options;
  const { maxWait, interval } = clampWaitParams(
    options.maxWaitSeconds ?? DEFAULT_WAIT_SECONDS,
    options.intervalSeconds ?? DEFAULT_WAIT_INTERVAL_SECONDS
  );

  // Resolve the idempotent baseline: an explicit `after` wins, otherwise anchor
  // to the caller's own last message (fetched once).
  let afterId: number;
  if (after !== undefined) {
    const parsed = parseMessageId(after);
    if (parsed === null) {
      return err(createError(
        ErrorCode.INVALID_INPUT,
        `'after' must be a numeric Teams message ID (got "${after}"). Use the id of a message from teams_get_thread.`
      ));
    }
    afterId = parsed;
  } else {
    const initial = await getThreadMessages(conversationId, { limit, order: 'desc' });
    if (!initial.ok) return initial;
    afterId = resolveWaitBaseline(initial.value.messages);
  }

  const start = Date.now();
  const deadline = start + maxWait * 1000;
  let polls = 0;
  let newMessages: ThreadMessage[] = [];

  for (;;) {
    polls++;
    const page = await getThreadMessages(conversationId, { limit, order: 'asc' });
    if (!page.ok) return page;
    newMessages = selectNewMessages(page.value.messages, afterId, includeSelf);
    if (newMessages.length > 0) break;
    const remaining = deadline - Date.now();
    if (remaining <= 0) break;
    await sleep(Math.min(interval * 1000, remaining));
  }

  // ponytail: advance the cursor to the last returned id. If more than `limit`
  // messages arrive between two polls the surplus is re-fetched next call by id;
  // a burst larger than `limit` could strand the oldest — upgrade to a
  // backfill scan (like the Rust CLI) only if that proves to matter.
  const nextAfter = newMessages.length > 0
    ? newMessages[newMessages.length - 1].id
    : String(afterId);
  const waitedSeconds = Math.round((Date.now() - start) / 1000);

  return ok({
    timedOut: newMessages.length === 0,
    after: String(afterId),
    nextAfter,
    newMessages,
    polls,
    waitedSeconds,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Internal Helpers (reaction name enrichment)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Enriches reaction display names via the fetchShortProfile batch API.
 *
 * Collects MRIs missing display names from the given reactions, resolves
 * them in a single batch call, then updates the reaction objects in-place.
 */
async function enrichReactionNamesFromProfiles(reactions?: Reaction[]): Promise<void> {
  if (!reactions || reactions.length === 0) return;

  const unresolvedMris = reactions
    .filter(r => !r.user.displayName && r.user.mri)
    .map(r => r.user.mri);

  if (unresolvedMris.length === 0) return;

  // Deduplicate MRIs for the batch call
  const uniqueMris = [...new Set(unresolvedMris)];
  const nameMap = await resolveNames(uniqueMris);

  if (nameMap.size === 0) return;

  for (const reaction of reactions) {
    if (!reaction.user.displayName) {
      const name = nameMap.get(reaction.user.mri);
      if (name) reaction.user.displayName = name;
    }
  }
}
