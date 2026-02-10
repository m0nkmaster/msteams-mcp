/**
 * Chat Service API client for messaging operations.
 * 
 * Handles all calls to teams.microsoft.com/api/chatsvc endpoints.
 * Base URL is extracted from session config to support different Teams environments.
 */

import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getMessagingHeaders, getSkypeAuthHeaders, getTeamsHeaders } from '../utils/api-config.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import { getUserDisplayName } from '../auth/token-extractor.js';
import { requireMessageAuth, getRegion, getTeamsBaseUrl } from '../utils/auth-guards.js';
import { stripHtml, extractLinks, buildMessageLink, buildOneOnOneConversationId, extractObjectId, extractActivityTimestamp, parseVirtualConversationMessage, markdownToTeamsHtml, type ExtractedLink } from '../utils/parsers.js';
import { DEFAULT_ACTIVITY_LIMIT, SAVED_MESSAGES_ID, FOLLOWED_THREADS_ID, VIRTUAL_CONVERSATION_PREFIX, SELF_CHAT_ID, MRI_ORGID_PREFIX } from '../constants.js';

// Reusable date formatter for human-readable timestamps with day of week
// Hoisted to module scope to avoid creating a new formatter per message
const humanReadableDateFormatter = new Intl.DateTimeFormat('en-US', {
  weekday: 'long',
  day: 'numeric',
  month: 'long',
  year: 'numeric',
  hour: '2-digit',
  minute: '2-digit',
  timeZone: 'UTC',
  timeZoneName: 'short',
});

/**
 * Formats an ISO timestamp into a human-readable string with day of week.
 * This helps LLMs correctly identify the day without needing to calculate it.
 * Example: "Friday, January 30, 2026, 10:45 AM UTC"
 */
function formatHumanReadableDate(isoTimestamp: string): string {
  try {
    const date = new Date(isoTimestamp);
    if (isNaN(date.getTime())) return '';
    return humanReadableDateFormatter.format(date);
  } catch {
    return '';
  }
}

/** Gets region and base URL together for API calls. */
function getApiConfig() {
  return {
    region: getRegion(),
    baseUrl: getTeamsBaseUrl(),
  };
}

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

/** Result of saving/unsaving a message. */
export interface SaveMessageResult {
  conversationId: string;
  messageId: string;
  saved: boolean;
}

/** A saved message from the virtual 48:saved conversation. */
export interface SavedMessage {
  id: string;
  content: string;
  contentType: string;
  sender: {
    mri: string;
    displayName?: string;
  };
  timestamp: string;
  /** The original conversation where this message lives. */
  sourceConversationId: string;
  /** The original message ID in the source conversation. */
  sourceMessageId?: string;
  messageLink?: string;
  links?: ExtractedLink[];
}

/** Result of getting saved messages. */
export interface GetSavedMessagesResult {
  messages: SavedMessage[];
}

/** A followed thread from the virtual 48:threads conversation. */
export interface FollowedThread {
  id: string;
  content: string;
  contentType: string;
  sender: {
    mri: string;
    displayName?: string;
  };
  timestamp: string;
  /** The original conversation/thread where this post lives. */
  sourceConversationId: string;
  /** The original post ID in the source conversation. */
  sourcePostId?: string;
  messageLink?: string;
  links?: ExtractedLink[];
}

/** Result of getting followed threads. */
export interface GetFollowedThreadsResult {
  threads: FollowedThread[];
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
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { replyToMessageId } = options;
  const { region, baseUrl } = getApiConfig();
  const displayName = getUserDisplayName() || 'User';

  // Generate unique message ID
  const clientMessageId = Date.now().toString();

  // Process content: handle mentions, links, and markdown formatting
  let htmlContent: string;
  let mentionsToSend: Mention[] = [];

  // If content is already HTML, pass through as-is
  if (content.startsWith('<')) {
    htmlContent = content;
  } else {
    // Check for mentions/links that need special processing
    const parsed = parseContentWithMentionsAndLinks(content);
    htmlContent = parsed.html;
    mentionsToSend = parsed.mentions;
  }

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
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
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
    const msg = raw as Record<string, unknown>;

    // Skip non-message types
    const messageType = msg.messagetype as string;
    if (!messageType || messageType.startsWith('Control/') || messageType === 'ThreadActivity/AddMember') {
      continue;
    }

    // Skip deleted messages (they have empty content and a deletetime property)
    const properties = msg.properties as Record<string, unknown> | undefined;
    if (properties?.deletetime) {
      continue;
    }

    const id = msg.id as string || msg.originalarrivaltime as string;
    if (!id) continue;

    const content = msg.content as string || '';
    const contentType = msg.messagetype as string || 'Text';

    const fromMri = msg.from as string || '';
    const displayName = msg.imdisplayname as string || msg.displayName as string;

    const timestamp = msg.originalarrivaltime as string ||
      msg.composetime as string ||
      new Date(parseInt(id, 10)).toISOString();

    // Build message link
    const messageLink = /^\d+$/.test(id)
      ? buildMessageLink(conversationId, id)
      : undefined;

    // Extract links before stripping HTML
    const links = extractLinks(content);

    // Format human-readable date with day of week to help LLMs
    const when = formatHumanReadableDate(timestamp);

    // Extract thread root ID for channel messages
    // When rootMessageId differs from id, this message is a reply within a thread
    const rootMessageId = msg.rootMessageId as string | undefined;
    const isThreadReply = !!rootMessageId && rootMessageId !== id;

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
      clientMessageId: msg.clientmessageid as string,
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

// Regex pattern for extracting message ID from saved messages: T_{conversationId}_M_{messageId}
const SAVED_MESSAGE_PATTERN = /_M_(\d+)$/;

// Regex pattern for extracting post ID from followed threads: T_{conversationId}_P_{postId}_Threads
const FOLLOWED_THREAD_PATTERN = /_P_(\d+)_Threads$/;

/**
 * Gets saved (bookmarked) messages from the virtual 48:saved conversation.
 * Returns messages the user has bookmarked across all conversations.
 */
export async function getSavedMessages(
  options: { limit?: number } = {}
): Promise<Result<GetSavedMessagesResult>> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const limit = options.limit ?? 50;

  let url = CHATSVC_API.messages(region, SAVED_MESSAGES_ID, undefined, baseUrl);
  url += `?view=msnp24Equivalent|supportsMessageProperties&pageSize=${limit}&startTime=1`;

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
    return ok({ messages: [] });
  }

  const messages: SavedMessage[] = [];

  for (const raw of rawMessages) {
    const parsed = parseVirtualConversationMessage(
      raw as Record<string, unknown>,
      SAVED_MESSAGE_PATTERN
    );
    if (!parsed) continue;

    messages.push({
      id: parsed.id,
      content: parsed.content,
      contentType: parsed.contentType,
      sender: parsed.sender,
      timestamp: parsed.timestamp,
      sourceConversationId: parsed.sourceConversationId,
      sourceMessageId: parsed.sourceReferenceId,
      messageLink: parsed.messageLink,
    });
  }

  // Sort by timestamp (newest first for saved messages)
  messages.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  return ok({ messages });
}

/**
 * Gets followed threads from the virtual 48:threads conversation.
 * Returns threads the user is following for updates.
 */
export async function getFollowedThreads(
  options: { limit?: number } = {}
): Promise<Result<GetFollowedThreadsResult>> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const limit = options.limit ?? 50;

  let url = CHATSVC_API.messages(region, FOLLOWED_THREADS_ID, undefined, baseUrl);
  url += `?view=msnp24Equivalent|supportsMessageProperties&pageSize=${limit}&startTime=1`;

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
    return ok({ threads: [] });
  }

  const threads: FollowedThread[] = [];

  for (const raw of rawMessages) {
    const parsed = parseVirtualConversationMessage(
      raw as Record<string, unknown>,
      FOLLOWED_THREAD_PATTERN
    );
    if (!parsed) continue;

    threads.push({
      id: parsed.id,
      content: parsed.content,
      contentType: parsed.contentType,
      sender: parsed.sender,
      timestamp: parsed.timestamp,
      sourceConversationId: parsed.sourceConversationId,
      sourcePostId: parsed.sourceReferenceId,
      messageLink: parsed.messageLink,
    });
  }

  // Sort by timestamp (newest first for followed threads)
  threads.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  return ok({ threads });
}

/**
 * Saves (bookmarks) a message.
 * 
 * @param rootMessageId - For channel threaded replies, the ID of the thread root post.
 *                        For top-level posts and non-channel messages, omit or pass undefined.
 */
export async function saveMessage(
  conversationId: string,
  messageId: string,
  rootMessageId?: string
): Promise<Result<SaveMessageResult>> {
  return setMessageSavedState(conversationId, messageId, true, rootMessageId);
}

/**
 * Unsaves (removes bookmark from) a message.
 * 
 * @param rootMessageId - For channel threaded replies, the ID of the thread root post.
 *                        For top-level posts and non-channel messages, omit or pass undefined.
 */
export async function unsaveMessage(
  conversationId: string,
  messageId: string,
  rootMessageId?: string
): Promise<Result<SaveMessageResult>> {
  return setMessageSavedState(conversationId, messageId, false, rootMessageId);
}

/**
 * Internal function to set the saved state of a message.
 * 
 * The rcmetadata API uses a two-ID system:
 * - URL path: rootMessageId (thread root for channel replies, or messageId for top-level)
 * - Body: mid (the actual message being saved/unsaved)
 */
async function setMessageSavedState(
  conversationId: string,
  messageId: string,
  saved: boolean,
  rootMessageId: string | undefined
): Promise<Result<SaveMessageResult>> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  // For channel threaded replies, rootMessageId is the thread root post ID
  // For top-level posts and non-channel messages, use the messageId itself
  const urlMessageId = rootMessageId ?? messageId;
  const url = CHATSVC_API.messageMetadata(region, conversationId, urlMessageId, baseUrl);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'PUT',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify({
        s: saved ? 1 : 0,
        mid: parseInt(messageId, 10),
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    conversationId,
    messageId,
    saved,
  });
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
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const displayName = getUserDisplayName() || 'User';
  
  // Convert markdown formatting to Teams HTML
  const htmlContent = newContent.startsWith('<') ? newContent : markdownToTeamsHtml(newContent);

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
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
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

/**
 * Gets properties for a single conversation.
 */
export async function getConversationProperties(
  conversationId: string
): Promise<Result<{ displayName?: string; conversationType?: string }>> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const url = CHATSVC_API.conversation(region, conversationId, baseUrl) + '?view=msnp24Equivalent';

  const response = await httpRequest<Record<string, unknown>>(
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
  const threadProps = data.threadProperties as Record<string, unknown> | undefined;
  const productType = threadProps?.productThreadType as string | undefined;

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
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return ok(undefined); // Non-critical: just return undefined if not authenticated
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
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
    const m = msg as Record<string, unknown>;
    const fromMri = m.from as string || '';
    const displayName = m.imdisplayname as string;

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
  
  // Build result with mentions/links as inline elements
  // Text segments between mentions/links get markdown conversion
  const mentions: Mention[] = [];
  let result = '';
  let lastIndex = 0;
  let mentionId = 0;
  
  for (const m of matches) {
    // Add markdown-converted text before this match
    const textBefore = content.substring(lastIndex, m.index);
    if (textBefore) {
      result += markdownToTeamsHtml(textBefore);
    }
    
    if (m.type === 'mention') {
      mentions.push({ mri: m.target, displayName: m.text });
      result += buildMentionHtml(m.text, mentionId);
      mentionId++;
    } else {
      // Link
      const safeText = escapeHtml(m.text);
      const safeUrl = m.target.replace(/"/g, '&quot;');
      result += `<a href="${safeUrl}">${safeText}</a>`;
    }
    
    lastIndex = m.index + m.length;
  }
  
  // Add remaining text with markdown conversion
  const remaining = content.substring(lastIndex);
  if (remaining) {
    result += markdownToTeamsHtml(remaining);
  }
  
  return { html: result, mentions };
}

/** Result of getting a 1:1 conversation. */
export interface GetOneOnOneChatResult {
  conversationId: string;
  otherUserId: string;
  currentUserId: string;
}

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
// Group Chat Operations
// ─────────────────────────────────────────────────────────────────────────────

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

/**
 * Creates a new group chat with multiple members.
 * 
 * Unlike 1:1 chats which have a predictable ID format, group chats require
 * an API call to create. The conversation ID is returned by the server.
 * 
 * @param memberIdentifiers - Array of user identifiers (MRI, object ID, or GUID)
 * @param topic - Optional chat topic/name
 * @param region - API region (default: 'amer')
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
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  // Validate we have at least 2 other members (plus current user = 3+ total)
  if (!memberIdentifiers || memberIdentifiers.length < 2) {
    return err(createError(
      ErrorCode.INVALID_INPUT,
      'Group chat requires at least 2 other members. For 1:1 chats, use teams_get_chat instead.'
    ));
  }

  const { region, baseUrl } = getApiConfig();

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

  const response = await httpRequest<{ threadResource?: { id?: string } }>(
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
  const responseData = response.value.data as Record<string, unknown>;
  const threadResource = responseData.threadResource as Record<string, unknown> | undefined;
  let conversationId = (typeof threadResource?.id === 'string' ? threadResource.id : undefined)
    || (typeof responseData.id === 'string' ? responseData.id : undefined)
    || (typeof responseData.threadId === 'string' ? responseData.threadId : undefined);
  
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
// Consumption Horizon (Read Status) Operations
// ─────────────────────────────────────────────────────────────────────────────

/** Consumption horizon information for a conversation. */
export interface ConsumptionHorizonInfo {
  /** The conversation/thread ID. */
  conversationId: string;
  /** The version timestamp of the consumption horizon. */
  version?: string;
  /** The last read message ID (timestamp). */
  lastReadMessageId?: string;
  /** The last read timestamp. */
  lastReadTimestamp?: number;
  /** Raw consumption horizons array from API. */
  consumptionHorizons: Array<{
    id: string;
    consumptionHorizon: string;
  }>;
}

/** Result of marking a conversation as read. */
export interface MarkAsReadResult {
  conversationId: string;
  markedUpTo: string;
}

/**
 * Gets the consumption horizon (read receipts) for a conversation.
 * The consumption horizon indicates where each user has read up to.
 */
export async function getConsumptionHorizon(
  conversationId: string
): Promise<Result<ConsumptionHorizonInfo>> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const url = CHATSVC_API.consumptionHorizons(region, conversationId, baseUrl);

  const response = await httpRequest<{
    id?: string;
    version?: string;
    consumptionhorizons?: Array<{
      id: string;
      consumptionhorizon: string;
    }>;
  }>(
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
  const horizons = data.consumptionhorizons || [];

  // Find the current user's consumption horizon
  let lastReadMessageId: string | undefined;
  let lastReadTimestamp: number | undefined;

  for (const h of horizons) {
    if (h.id === auth.userMri || h.id.includes(auth.userMri)) {
      // Consumption horizon format: "{timestamp};{timestamp};{messageId}"
      const parts = h.consumptionhorizon.split(';');
      if (parts.length >= 3) {
        lastReadMessageId = parts[2];
        lastReadTimestamp = parseInt(parts[0], 10);
      }
      break;
    }
  }

  return ok({
    conversationId,
    version: data.version,
    lastReadMessageId,
    lastReadTimestamp,
    consumptionHorizons: horizons.map(h => ({
      id: h.id,
      consumptionHorizon: h.consumptionhorizon,
    })),
  });
}

/**
 * Marks a conversation as read up to a specific message.
 */
export async function markAsRead(
  conversationId: string,
  messageId: string
): Promise<Result<MarkAsReadResult>> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const url = CHATSVC_API.updateConsumptionHorizon(region, conversationId, baseUrl);

  // Format: "{messageId};{messageId};{messageId}" - all three values are the same
  const consumptionHorizon = `${messageId};${messageId};${messageId}`;

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'PUT',
      headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify({
        consumptionhorizon: consumptionHorizon,
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    conversationId,
    markedUpTo: messageId,
  });
}

/**
 * Gets unread count for a conversation by comparing consumption horizon
 * with recent messages.
 */
export async function getUnreadStatus(
  conversationId: string
): Promise<Result<{
  conversationId: string;
  unreadCount: number;
  lastReadMessageId?: string;
  latestMessageId?: string;
}>> {
  // Get consumption horizon
  const horizonResult = await getConsumptionHorizon(conversationId);
  if (!horizonResult.ok) {
    return horizonResult;
  }

  // Get recent messages
  const messagesResult = await getThreadMessages(conversationId, { limit: 50 });
  if (!messagesResult.ok) {
    return messagesResult;
  }

  const lastReadId = horizonResult.value.lastReadMessageId;
  const messages = messagesResult.value.messages;

  // Count messages after the last read position
  let unreadCount = 0;
  let latestMessageId: string | undefined;

  // Messages are sorted oldest-first, so reverse to process newest-first
  const reversedMessages = [...messages].reverse();

  for (const msg of reversedMessages) {
    if (!latestMessageId && !msg.isFromMe) {
      latestMessageId = msg.id;
    }

    if (lastReadId && msg.id === lastReadId) {
      break;
    }

    // Count messages not from the current user
    if (!msg.isFromMe) {
      unreadCount++;
    }
  }

  // If last read message wasn't in our window, count is a lower bound

  return ok({
    conversationId,
    unreadCount,
    lastReadMessageId: lastReadId,
    latestMessageId,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Activity Feed Operations
// ─────────────────────────────────────────────────────────────────────────────

/** Type of activity item. */
export type ActivityType = 'mention' | 'reaction' | 'reply' | 'message' | 'unknown';

/** An activity item from the notifications feed. */
export interface ActivityItem {
  /** Activity/message ID. */
  id: string;
  /** Type of activity (mention, reaction, reply, etc.). */
  type: ActivityType;
  /** Activity content (HTML stripped). */
  content: string;
  /** Raw content type from API. */
  contentType: string;
  /** Sender information. */
  sender: {
    mri: string;
    displayName?: string;
  };
  /** When the activity occurred. */
  timestamp: string;
  /** The source conversation ID (where the activity happened). */
  conversationId?: string;
  /** The conversation/thread topic name. */
  topic?: string;
  /** Direct link to the activity in Teams. */
  activityLink?: string;
  /** Links extracted from content. */
  links?: ExtractedLink[];
}

/** Result of fetching the activity feed. */
export interface GetActivityResult {
  /** Activity items (newest first). */
  activities: ActivityItem[];
  /** Sync state for incremental polling. */
  syncState?: string;
}

/**
 * Determines the activity type from message content and properties.
 */
function detectActivityType(msg: Record<string, unknown>): ActivityType {
  const content = msg.content as string || '';
  const messageType = msg.messagetype as string || '';
  
  // Check for @mention
  if (content.includes('itemtype="http://schema.skype.com/Mention"') ||
      content.includes('itemscope itemtype="http://schema.skype.com/Mention"')) {
    return 'mention';
  }
  
  // Check for reaction-related message types
  if (messageType.toLowerCase().includes('reaction')) {
    return 'reaction';
  }
  
  // Check for thread/reply indicators
  if (msg.threadtopic || msg.parentMessageId) {
    return 'reply';
  }
  
  // Standard message
  if (messageType.includes('RichText') || messageType.includes('Text')) {
    return 'message';
  }
  
  return 'unknown';
}

/**
 * Gets the activity feed (notifications) for the current user.
 * Includes mentions, reactions, replies, and other notifications.
 */
export async function getActivityFeed(
  options: { limit?: number } = {}
): Promise<Result<GetActivityResult>> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const limit = options.limit ?? DEFAULT_ACTIVITY_LIMIT;

  let url = CHATSVC_API.activityFeed(region, baseUrl);
  url += `?view=msnp24Equivalent&pageSize=${limit}`;

  const response = await httpRequest<{ messages?: unknown[]; syncState?: string }>(
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
  const syncState = response.value.data.syncState;
  
  if (!Array.isArray(rawMessages)) {
    return ok({
      activities: [],
      syncState,
    });
  }

  const activities: ActivityItem[] = [];

  for (const raw of rawMessages) {
    const msg = raw as Record<string, unknown>;

    // Skip control/system messages that aren't relevant
    const messageType = msg.messagetype as string;
    if (!messageType || 
        messageType.startsWith('Control/') || 
        messageType === 'ThreadActivity/AddMember' ||
        messageType === 'ThreadActivity/DeleteMember') {
      continue;
    }

    const id = msg.id as string || msg.originalarrivaltime as string;
    if (!id) continue;

    const content = msg.content as string || '';
    const contentType = msg.messagetype as string || 'Text';

    const fromMri = msg.from as string || '';
    const displayName = msg.imdisplayname as string || msg.displayName as string;

    // Safely extract timestamp - returns null if no valid timestamp found
    const timestamp = extractActivityTimestamp(msg);
    if (!timestamp) continue;

    // Get source conversation - prefer clumpId (actual source) over conversationid
    // Some activity items have conversationid as "48:notifications" (the virtual conversation)
    // which doesn't work for deep links. clumpId contains the real source conversation.
    const rawConversationId = msg.conversationid as string || msg.conversationId as string;
    const clumpId = msg.clumpId as string;
    
    // Use clumpId if conversationid is a virtual conversation (48:xxx format)
    const isVirtualConversation = rawConversationId?.startsWith(VIRTUAL_CONVERSATION_PREFIX);
    const conversationId = (isVirtualConversation && clumpId) ? clumpId : rawConversationId;
    
    const topic = msg.threadtopic as string || msg.topic as string;

    // Build activity link if we have a valid source conversation context
    // Skip virtual conversations (48:xxx) as they don't produce working deep links
    let activityLink: string | undefined;
    if (conversationId && !conversationId.startsWith(VIRTUAL_CONVERSATION_PREFIX) && /^\d+$/.test(id)) {
      activityLink = buildMessageLink(conversationId, id);
    }

    const activityType = detectActivityType(msg);

    // Extract links before stripping HTML
    const links = extractLinks(content);

    activities.push({
      id,
      type: activityType,
      content: stripHtml(content),
      contentType,
      sender: {
        mri: fromMri,
        displayName,
      },
      timestamp,
      conversationId,
      topic,
      activityLink,
      links: links.length > 0 ? links : undefined,
    });
  }

  // Sort by timestamp (newest first for activity feed)
  activities.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

  return ok({
    activities,
    syncState,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Reaction (Emotions) Operations
// ─────────────────────────────────────────────────────────────────────────────

/** Result of adding/removing a reaction. */
export interface ReactionResult {
  conversationId: string;
  messageId: string;
  emoji: string;
}

/**
 * Adds a reaction (emoji) to a message.
 * 
 * @param conversationId - The conversation containing the message
 * @param messageId - The message ID to react to
 * @param emojiKey - The emoji key (e.g., 'like', 'heart', 'laugh')
 */
export async function addReaction(
  conversationId: string,
  messageId: string,
  emojiKey: string
): Promise<Result<ReactionResult>> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const url = CHATSVC_API.messageEmotions(region, conversationId, messageId, baseUrl);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'PUT',
      headers: {
        ...getTeamsHeaders(),
        'Authentication': `skypetoken=${auth.skypeToken}`,
        'Authorization': `Bearer ${auth.authToken}`,
      },
      body: JSON.stringify({
        emotions: {
          key: emojiKey,
          value: Date.now(),
        },
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    conversationId,
    messageId,
    emoji: emojiKey,
  });
}

/**
 * Removes a reaction (emoji) from a message.
 * 
 * @param conversationId - The conversation containing the message
 * @param messageId - The message ID to remove the reaction from
 * @param emojiKey - The emoji key to remove (e.g., 'like', 'heart')
 */
export async function removeReaction(
  conversationId: string,
  messageId: string,
  emojiKey: string
): Promise<Result<ReactionResult>> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const auth = authResult.value;

  const { region, baseUrl } = getApiConfig();
  const url = CHATSVC_API.messageEmotions(region, conversationId, messageId, baseUrl);

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'DELETE',
      headers: {
        ...getTeamsHeaders(),
        'Authentication': `skypetoken=${auth.skypeToken}`,
        'Authorization': `Bearer ${auth.authToken}`,
      },
      body: JSON.stringify({
        emotions: {
          key: emojiKey,
        },
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({
    conversationId,
    messageId,
    emoji: emojiKey,
  });
}
