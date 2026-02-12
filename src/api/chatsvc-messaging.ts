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
import { stripHtml, extractLinks, buildMessageLink, buildOneOnOneConversationId, extractObjectId, markdownToTeamsHtml, type ExtractedLink } from '../utils/parsers.js';
import { SELF_CHAT_ID, MRI_ORGID_PREFIX } from '../constants.js';
import { formatHumanReadableDate } from './chatsvc-common.js';
import type { RawChatsvcMessage, RawConversationResponse, RawCreateThreadResponse } from '../types/api-responses.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Result of sending a message. */
export interface SendMessageResult {
  messageId: string;
  timestamp?: number;
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

/** A mention to include in a message (internal). */
interface Mention {
  /** The user's MRI (e.g., '8:orgid:uuid'). */
  mri: string;
  /** Display name to show for the mention. */
  displayName: string;
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

  const { replyToMessageId } = options;
  const displayName = getUserDisplayName() || 'User';

  // Generate unique message ID
  const clientMessageId = Date.now().toString();

  // Process content: handle mentions, links, and markdown formatting.
  // Always convert through markdown→HTML pipeline (never pass user content through
  // without sanitization, as Teams requires proper block-level wrapping like <p> tags)
  const parsed = parseContentWithMentionsAndLinks(content);
  const htmlContent = parsed.html;
  const mentionsToSend = parsed.mentions;

  // Build the message body
  const body: Record<string, unknown> = {
    content: htmlContent,
    messagetype: 'RichText/Html',
    contenttype: 'text',
    imdisplayname: displayName,
    clientmessageid: clientMessageId,
  };

  // Add mentions property if mentions were found
  if (mentionsToSend.length > 0) {
    body.properties = {
      mentions: buildMentionsProperty(mentionsToSend),
    };
  }

  const url = CHATSVC_API.messages(region, conversationId, replyToMessageId, baseUrl);

  const response = await httpRequest<{ OriginalArrivalTime?: number }>(
    url,
    {
      method: 'POST',
      headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    messageId: clientMessageId,
    timestamp: response.value.data.OriginalArrivalTime,
  });
}

/**
 * Sends a message to your own notes/self-chat.
 */
export async function sendNoteToSelf(content: string): Promise<Result<SendMessageResult>> {
  return sendMessage(SELF_CHAT_ID, content);
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
 */
export async function getThreadMessages(
  conversationId: string,
  options: { limit?: number; startTime?: number; order?: 'asc' | 'desc' } = {}
): Promise<Result<GetThreadResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const limit = options.limit ?? 50;

  let url = CHATSVC_API.messages(region, conversationId, undefined, baseUrl);
  url += `?view=msnp24Equivalent&pageSize=${limit}`;

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
      (() => {
        const parsed = parseInt(id, 10);
        return !isNaN(parsed) && parsed > 0 ? new Date(parsed).toISOString() : new Date().toISOString();
      })();

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
    });
  }

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
 * Note: You can only edit your own messages. The API will reject
 * attempts to edit messages from other users.
 * 
 * @param conversationId - The conversation containing the message
 * @param messageId - The ID of the message to edit
 * @param newContent - The new content for the message
 */
export async function editMessage(
  conversationId: string,
  messageId: string,
  newContent: string
): Promise<Result<EditMessageResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { auth, region, baseUrl } = authResult.value;
  const displayName = getUserDisplayName() || 'User';
  
  // Always convert through markdown→HTML pipeline (never pass raw HTML through,
  // as Teams requires proper block-level wrapping like <p> tags)
  const htmlContent = markdownToTeamsHtml(newContent);

  // Build the edit request body
  // The API requires the message structure with updated content
  const body = {
    id: messageId,
    type: 'Message',
    conversationid: conversationId,
    content: htmlContent,
    messagetype: 'RichText/Html',
    contenttype: 'text',
    imdisplayname: displayName,
  };

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
 * Escapes HTML special characters.
 */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Builds the HTML for a single mention.
 */
function buildMentionHtml(displayName: string, itemId: number): string {
  return `<readonly class="skipProofing" itemtype="http://schema.skype.com/Mention" contenteditable="false" spellcheck="false"><span itemtype="http://schema.skype.com/Mention" itemscope itemid="${itemId}">${escapeHtml(displayName)}</span></readonly>`;
}

/**
 * Builds the mentions property array for the API request.
 */
function buildMentionsProperty(mentions: Mention[]): string {
  const mentionObjects = mentions.map((mention, index) => ({
    '@type': 'http://schema.skype.com/Mention',
    'itemid': String(index),
    'mri': mention.mri,
    'mentionType': 'person',
    'displayName': mention.displayName,
  }));
  return JSON.stringify(mentionObjects);
}

/**
 * Parses content for both mentions @[Name](mri) and links [text](url).
 * Processes them in a single pass to avoid escaping conflicts.
 */
function parseContentWithMentionsAndLinks(content: string): { html: string; mentions: Mention[] } {
  // Patterns for mentions and links
  // Note: Link pattern uses [^)\s] to reject URLs with spaces
  const mentionPattern = /@\[([^\]]+)\]\(([^)]+)\)/g;
  const linkPattern = /(?<!@)\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g;
  
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
      html = buildMentionHtml(m.text, mentionId);
      mentionId++;
    } else {
      const safeText = escapeHtml(m.text);
      const safeUrl = m.target.replace(/"/g, '&quot;');
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
