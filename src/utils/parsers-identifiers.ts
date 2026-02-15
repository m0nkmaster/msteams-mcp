/**
 * Message link building, conversation type detection, ID extraction, and timestamp utilities.
 */

// ─────────────────────────────────────────────────────────────────────────────
// Conversation Type
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Determines the conversation type from a Teams conversation ID.
 * 
 * Conversation ID formats:
 * - Channels: 19:xxx@thread.tacv2
 * - Meetings: 19:meeting_xxx@thread.v2
 * - 1:1 chats: 19:guid_guid@unq.gbl.spaces
 * - Group chats: 19:xxx@thread.v2 (non-meeting)
 */
export function getConversationType(conversationId: string): 'channel' | 'meeting' | 'chat' {
  if (conversationId.includes('@thread.tacv2')) {
    return 'channel';
  }
  if (conversationId.includes('meeting_')) {
    return 'meeting';
  }
  // 1:1 chats (@unq.gbl.spaces) and group chats (@thread.v2) both need chat context
  return 'chat';
}

// ─────────────────────────────────────────────────────────────────────────────
// Message Deep Links
// ─────────────────────────────────────────────────────────────────────────────

/** Default Teams base URL for message links. */
const DEFAULT_TEAMS_LINK_BASE = 'https://teams.microsoft.com';

/** Options for building a message deep link. */
export interface MessageLinkOptions {
  /** The conversation/channel/chat ID (e.g., "19:xxx@thread.tacv2"). */
  conversationId: string;
  /** The message timestamp in epoch milliseconds. */
  messageId: string | number;
  /** Tenant ID (GUID) - required for reliable deep links. */
  tenantId?: string;
  /** For channel messages: the team's group ID (GUID). */
  groupId?: string;
  /** For channel messages: the parent/root message ID (epoch ms). If omitted for channels, messageId is used. */
  parentMessageId?: string;
  /** Teams base URL (for GCC/GCC-High support). */
  teamsBaseUrl?: string;
}

/** Context for building reliable message deep links. */
export interface LinkContext {
  /** Tenant ID (GUID) from session. */
  tenantId?: string;
  /** Teams base URL (for GCC/GCC-High support). */
  teamsBaseUrl?: string;
}

/**
 * Builds a deep link to open a message in Teams.
 * 
 * Uses Microsoft's documented deep link formats:
 * - Channels: /l/message/{channelId}/{msgId}?tenantId=...&groupId=...&parentMessageId=...
 * - Chats/Meetings: /l/message/{chatId}/{msgId}?tenantId=...&context={"contextType":"chat"}
 * 
 * @see https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-link-teams
 */
export function buildMessageLink(opts: MessageLinkOptions): string {
  const base = opts.teamsBaseUrl ?? DEFAULT_TEAMS_LINK_BASE;
  const msgId = typeof opts.messageId === 'string' ? opts.messageId : String(opts.messageId);
  const convType = getConversationType(opts.conversationId);

  // Build the base URL path — encode the conversationId for URL safety
  const linkUrl = `${base}/l/message/${encodeURIComponent(opts.conversationId)}/${msgId}`;

  const params = new URLSearchParams();

  if (convType === 'channel') {
    // Channel deep links require tenantId, groupId, and parentMessageId
    if (opts.tenantId) params.set('tenantId', opts.tenantId);
    if (opts.groupId) params.set('groupId', opts.groupId);
    // parentMessageId is always required — for top-level posts it equals the messageId
    params.set('parentMessageId', opts.parentMessageId ?? msgId);
    params.set('createdTime', msgId);
  } else {
    // Chat and meeting deep links require tenantId and context
    if (opts.tenantId) params.set('tenantId', opts.tenantId);
    params.set('context', '{"contextType":"chat"}');
  }

  const qs = params.toString();
  return qs ? `${linkUrl}?${qs}` : linkUrl;
}

// ─────────────────────────────────────────────────────────────────────────────
// Message Timestamps
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Extracts a timestamp-based message ID from various sources.
 * Teams uses epoch milliseconds as message IDs in URLs.
 * 
 * IMPORTANT: For channel threaded replies, the ;messageid= in ClientConversationId
 * is the PARENT thread's ID, not this message's ID. We must prefer the actual
 * message timestamp (DateTimeReceived/DateTimeSent) for accurate deep links.
 */
export function extractMessageTimestamp(
  source: Record<string, unknown> | undefined,
  timestamp?: string
): string | undefined {
  // FIRST: Try to compute from the message's own timestamp
  // This is the most reliable for channel threaded replies
  if (timestamp) {
    try {
      const date = new Date(timestamp);
      if (!isNaN(date.getTime())) {
        return String(date.getTime());
      }
    } catch {
      // Ignore parsing errors
    }
  }
  
  // SECOND: Try explicit MessageId fields
  if (source) {
    // Check for MessageId or Id in various formats
    const messageId = source.MessageId ?? source.OriginalMessageId ?? source.ReferenceObjectId;
    if (typeof messageId === 'string' && /^\d{13}$/.test(messageId)) {
      return messageId;
    }
    
    // LAST RESORT: Check ClientConversationId for ;messageid=xxx suffix
    // NOTE: For threaded replies, this is the PARENT message ID, so only use
    // if we couldn't get the actual timestamp above
    const clientConvId = source.ClientConversationId as string | undefined;
    if (clientConvId && clientConvId.includes(';messageid=')) {
      const match = clientConvId.match(/;messageid=(\d+)/);
      if (match) {
        return match[1];
      }
    }
  }
  
  return undefined;
}

/**
 * Safely extracts a timestamp from an activity feed message.
 * 
 * Tries multiple sources in order of preference:
 * 1. originalarrivaltime - Primary timestamp field
 * 2. composetime - When message was composed
 * 3. id as numeric timestamp - Fallback if ID is a Unix timestamp
 * 
 * Returns null if no valid timestamp can be determined, preventing
 * RangeError from Date operations on invalid values.
 * 
 * @param msg - Raw message object from activity feed API
 * @returns ISO timestamp string, or null if no valid timestamp found
 */
export function extractActivityTimestamp(msg: Record<string, unknown>): string | null {
  const arrivalTime = msg.originalarrivaltime as string;
  const composeTime = msg.composetime as string;
  
  if (arrivalTime) return arrivalTime;
  if (composeTime) return composeTime;

  // Try parsing the message ID as a numeric timestamp
  const id = msg.id as string;
  if (id) {
    const numericId = parseInt(id, 10);
    if (!isNaN(numericId) && numericId > 0) {
      return new Date(numericId).toISOString();
    }
  }
  
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// ID Extraction
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Decodes a base64-encoded GUID to its standard string representation.
 * 
 * Microsoft encodes GUIDs as 16 bytes with little-endian ordering for the
 * first three groups (Data1, Data2, Data3).
 * 
 * @param base64 - Base64-encoded GUID (typically 24 chars with == padding)
 * @returns The GUID string in standard format, or null if invalid
 */
export function decodeBase64Guid(base64: string): string | null {
  try {
    // Decode base64 to bytes
    const bytes = Buffer.from(base64, 'base64');
    
    // GUID is exactly 16 bytes
    if (bytes.length !== 16) {
      return null;
    }
    
    // Convert to hex
    const hex = bytes.toString('hex');
    
    // Format as GUID with little-endian byte ordering for first 3 groups
    // Data1 (4 bytes), Data2 (2 bytes), Data3 (2 bytes) are little-endian
    // Data4 (8 bytes) is big-endian
    const guid = [
      hex.slice(6, 8) + hex.slice(4, 6) + hex.slice(2, 4) + hex.slice(0, 2), // Data1
      hex.slice(10, 12) + hex.slice(8, 10),   // Data2
      hex.slice(14, 16) + hex.slice(12, 14),  // Data3
      hex.slice(16, 20),                       // Data4a
      hex.slice(20, 32),                       // Data4b
    ].join('-');
    
    return guid.toLowerCase();
  } catch {
    return null;
  }
}

/**
 * Checks if a string appears to be a base64-encoded GUID.
 * Base64-encoded 16 bytes = 24 characters (22 chars + 2 padding or no padding).
 */
function isLikelyBase64Guid(str: string): boolean {
  // Check length (22-24 chars for 16 bytes)
  if (str.length < 22 || str.length > 24) {
    return false;
  }
  
  // Must contain only base64 characters
  if (!/^[A-Za-z0-9+/]+=*$/.test(str)) {
    return false;
  }
  
  // Typically ends with == for 16 bytes
  return true;
}

/**
 * Extracts the Azure AD object ID (GUID) from various formats.
 * 
 * Handles:
 * - MRI format: "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24"
 * - MRI with base64: "8:orgid:93qkaTtFGWpUHjyRafgdhg=="
 * - Skype ID format: "orgid:ab76f827-27e2-4c67-a765-f1a53145fa24"
 * - ID with tenant: "ab76f827-27e2-4c67-a765-f1a53145fa24@56b731a8-..."
 * - Raw GUID: "ab76f827-27e2-4c67-a765-f1a53145fa24"
 * - Base64-encoded GUID: "93qkaTtFGWpUHjyRafgdhg=="
 * 
 * @param identifier - User identifier in any supported format
 * @returns The extracted GUID or null if invalid format
 */
export function extractObjectId(identifier: string): string | null {
  if (!identifier) return null;

  // Pattern for a GUID (with or without hyphens)
  const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

  /** Resolves an ID part to a lowercase GUID, trying direct match then base64 decode. */
  const resolveIdPart = (idPart: string): string | null => {
    if (guidPattern.test(idPart)) return idPart.toLowerCase();
    if (isLikelyBase64Guid(idPart)) return decodeBase64Guid(idPart);
    return null;
  };

  // Handle MRI format: "8:orgid:GUID" or "8:orgid:base64"
  if (identifier.startsWith('8:orgid:')) {
    return resolveIdPart(identifier.substring(8));
  }

  // Handle Skype ID format: "orgid:GUID" (from skype token's skypeid field)
  if (identifier.startsWith('orgid:')) {
    return resolveIdPart(identifier.substring(6));
  }

  // Handle ID with tenant: "GUID@tenantId"
  if (identifier.includes('@')) {
    return resolveIdPart(identifier.split('@')[0]);
  }

  // Handle raw GUID or base64-encoded GUID
  return resolveIdPart(identifier);
}

/**
 * Builds a 1:1 conversation ID from two user object IDs.
 * 
 * The conversation ID format for 1:1 chats in Teams is:
 * `19:{userId1}_{userId2}@unq.gbl.spaces`
 * 
 * The user IDs are sorted lexicographically to ensure consistency -
 * both participants will generate the same conversation ID.
 * 
 * @param userId1 - First user's object ID (GUID, MRI, or ID with tenant)
 * @param userId2 - Second user's object ID (GUID, MRI, or ID with tenant)
 * @returns The constructed conversation ID, or null if either ID is invalid
 */
export function buildOneOnOneConversationId(
  userId1: string,
  userId2: string
): string | null {
  const id1 = extractObjectId(userId1);
  const id2 = extractObjectId(userId2);

  if (!id1 || !id2) {
    return null;
  }

  // Sort lexicographically for consistent ID regardless of who initiates
  const sorted = [id1, id2].sort();

  return `19:${sorted[0]}_${sorted[1]}@unq.gbl.spaces`;
}
