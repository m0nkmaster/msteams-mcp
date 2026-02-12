/**
 * Pure parsing functions for Teams API responses.
 * 
 * These functions transform raw API responses into our internal types.
 * They are extracted here for testability - no side effects or external dependencies.
 */

import type { TeamsSearchResult, ExtractedLink } from '../types/teams.js';
import { MIN_CONTENT_LENGTH } from '../constants.js';

// Re-export ExtractedLink so existing imports from parsers.ts continue to work
export type { ExtractedLink };

/** Person search result from Substrate suggestions API. */
export interface PersonSearchResult {
  id: string;              // Azure AD object ID
  mri: string;             // Teams MRI (8:orgid:guid)
  displayName: string;     // Full display name
  email?: string;          // Primary email address
  givenName?: string;      // First name
  surname?: string;        // Last name
  jobTitle?: string;       // Job title
  department?: string;     // Department
  companyName?: string;    // Company name
}

/** User profile extracted from JWT tokens. */
export interface UserProfile {
  id: string;           // Azure AD object ID (oid)
  mri: string;          // Teams MRI (8:orgid:guid)
  email: string;        // User principal name / email
  displayName: string;  // Full display name
  givenName?: string;   // First name
  surname?: string;     // Last name
  tenantId?: string;    // Azure tenant ID
}

/**
 * Extracts links from HTML content before stripping.
 * Returns an array of { url, text } objects.
 */
export function extractLinks(html: string): ExtractedLink[] {
  const links: ExtractedLink[] = [];
  const linkRegex = /<a\s+[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  
  let match;
  while ((match = linkRegex.exec(html)) !== null) {
    const url = match[1];
    const text = stripHtml(match[2]); // Clean nested HTML in link text
    if (url && !url.startsWith('javascript:')) {
      links.push({ url, text: text || url });
    }
  }
  
  return links;
}

/**
 * Strips HTML tags from content for display.
 */
export function stripHtml(html: string): string {
  return html
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

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

/**
 * Builds a deep link to open a message in Teams.
 * 
 * Uses Microsoft's documented deep link formats:
 * - Channels: /l/message/{channelId}/{msgId}?tenantId=...&groupId=...&parentMessageId=...
 * - Chats/Meetings: /l/message/{chatId}/{msgId}?tenantId=...&context={"contextType":"chat"}
 * 
 * @see https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-link-teams
 */
export function buildMessageLink(opts: MessageLinkOptions): string;
/**
 * @deprecated Use the options object overload instead.
 */
export function buildMessageLink(
  conversationId: string,
  messageTimestamp: string | number,
  parentMessageId?: string,
  teamsBaseUrl?: string
): string;
export function buildMessageLink(
  optsOrConversationId: MessageLinkOptions | string,
  messageTimestamp?: string | number,
  parentMessageId?: string,
  teamsBaseUrl?: string
): string {
  // Normalise arguments: support both old positional and new options-object signatures
  let opts: MessageLinkOptions;
  if (typeof optsOrConversationId === 'string') {
    opts = {
      conversationId: optsOrConversationId,
      messageId: messageTimestamp!,
      parentMessageId,
      teamsBaseUrl,
    };
  } else {
    opts = optsOrConversationId;
  }

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
 * Parses a person suggestion from the Substrate API response.
 * 
 * The API can return IDs in various formats:
 * - GUID with tenant: "ab76f827-...@tenant.onmicrosoft.com"
 * - Base64-encoded GUID: "93qkaTtFGWpUHjyRafgdhg=="
 */
export function parsePersonSuggestion(item: Record<string, unknown>): PersonSearchResult | null {
  const rawId = item.Id as string;
  if (!rawId) return null;

  // Extract the ID part (strip tenant suffix if present)
  const idPart = rawId.includes('@') ? rawId.split('@')[0] : rawId;
  
  // Convert to a proper GUID format
  const objectId = extractObjectId(idPart);
  if (!objectId) {
    // If we can't parse the ID, skip this result
    return null;
  }
  
  // Build MRI from the decoded GUID if not provided
  const mri = (item.MRI as string) || `8:orgid:${objectId}`;
  
  const displayName = item.DisplayName as string || '';
  
  // EmailAddresses can be an array
  const emailAddresses = item.EmailAddresses as string[] | undefined;
  const email = emailAddresses?.[0];

  return {
    id: objectId,
    mri: mri.includes('orgid:') && !mri.includes('-') 
      ? `8:orgid:${objectId}`  // Rebuild MRI if it has base64
      : mri,
    displayName,
    email,
    givenName: item.GivenName as string | undefined,
    surname: item.Surname as string | undefined,
    jobTitle: item.JobTitle as string | undefined,
    department: item.Department as string | undefined,
    companyName: item.CompanyName as string | undefined,
  };
}

/** Context for building reliable message deep links. */
export interface LinkContext {
  /** Tenant ID (GUID) from session. */
  tenantId?: string;
  /** Teams base URL (for GCC/GCC-High support). */
  teamsBaseUrl?: string;
}

/**
 * Parses a v2 query result item into a search result.
 */
export function parseV2Result(item: Record<string, unknown>, linkContext?: LinkContext): TeamsSearchResult | null {
  const content = item.HitHighlightedSummary as string || 
                  item.Summary as string || 
                  '';
  
  if (content.length < MIN_CONTENT_LENGTH) return null;

  const id = item.Id as string || 
             item.ReferenceId as string || 
             `v2-${Date.now()}`;

  // Extract links before stripping HTML
  const links = extractLinks(content);
  const cleanContent = stripHtml(content);

  const source = item.Source as Record<string, unknown> | undefined;

  // Extract conversationId from extension fields or source properties
  // For channel threaded replies, we want the thread ID (ClientThreadId) not the channel ID
  let conversationId: string | undefined;
  if (source) {
    // Check ClientThreadId first - this is the specific thread for channel replies
    // Using this ensures the deep link goes to the correct thread context
    const clientThreadId = source.ClientThreadId;
    if (typeof clientThreadId === 'string' && clientThreadId.length > 0) {
      conversationId = clientThreadId;
    }
    
    // Fallback to Extensions.SkypeGroupId (the channel ID)
    if (!conversationId) {
      const extensions = source.Extensions as Record<string, unknown> | undefined;
      if (extensions) {
        const extId = extensions.SkypeSpaces_ConversationPost_Extension_SkypeGroupId;
        if (typeof extId === 'string' && extId.length > 0) {
          conversationId = extId;
        }
      }
    }
    
    // Fallback to ClientConversationId (strip ;messageid= suffix if present)
    if (!conversationId) {
      const clientConvId = source.ClientConversationId;
      if (typeof clientConvId === 'string' && clientConvId.length > 0) {
        conversationId = clientConvId.split(';')[0];
      }
    }
  }

  // Note: The API returns DateTimeReceived, DateTimeSent, DateTimeCreated (not ReceivedTime/CreatedDateTime)
  const timestamp = source?.DateTimeReceived as string || 
                    source?.DateTimeSent as string || 
                    source?.DateTimeCreated as string ||
                    source?.ReceivedTime as string ||  // Legacy fallback
                    source?.CreatedDateTime as string; // Legacy fallback
  
  // Extract message timestamp - used for both deep links and thread replies
  const messageTimestamp = extractMessageTimestamp(source, timestamp);
  
  // Extract parent message ID from ClientConversationId for thread replies
  // Format: "19:xxx@thread.tacv2;messageid=1769237777958"
  // If the messageid differs from the message's own timestamp, it's a thread reply
  let parentMessageId: string | undefined;
  if (source) {
    const clientConvId = source.ClientConversationId as string | undefined;
    if (clientConvId?.includes(';messageid=')) {
      const match = clientConvId.match(/;messageid=(\d+)/);
      if (match) {
        parentMessageId = match[1];
      }
    }
  }
  
  // Build message link if we have the required data
  let messageLink: string | undefined;
  if (conversationId && messageTimestamp) {
    messageLink = buildMessageLink({
      conversationId,
      messageId: messageTimestamp,
      tenantId: linkContext?.tenantId,
      parentMessageId,
      teamsBaseUrl: linkContext?.teamsBaseUrl,
    });
  }

  return {
    id,
    type: 'message',
    content: cleanContent,
    sender: source?.From as string || source?.Sender as string,
    timestamp,
    channelName: source?.ChannelName as string || source?.Topic as string,
    teamName: source?.TeamName as string || source?.GroupName as string,
    conversationId,
    // Use the timestamp as messageId (required for thread replies)
    // Fallback to ReferenceId if timestamp extraction fails
    messageId: messageTimestamp || item.ReferenceId as string,
    messageLink,
    links: links.length > 0 ? links : undefined,
  };
}

/**
 * Parses user profile from a JWT payload.
 * 
 * @param payload - Decoded JWT payload object
 * @returns User profile or null if required fields are missing
 */
export function parseJwtProfile(payload: Record<string, unknown>): UserProfile | null {
  const oid = payload.oid as string | undefined;
  const name = payload.name as string | undefined;
  
  if (!oid || !name) {
    return null;
  }
  
  const profile: UserProfile = {
    id: oid,
    mri: `8:orgid:${oid}`,
    email: (payload.upn || payload.preferred_username || payload.email || '') as string,
    displayName: name,
    tenantId: payload.tid as string | undefined,
  };
  
  // Try to extract given name and surname
  if (payload.given_name) {
    profile.givenName = payload.given_name as string;
  }
  if (payload.family_name) {
    profile.surname = payload.family_name as string;
  }
  
  // If no given/family name, try to parse from displayName
  if (!profile.givenName && profile.displayName.includes(',')) {
    // Format: "Surname, GivenName"
    const parts = profile.displayName.split(',').map(s => s.trim());
    if (parts.length === 2) {
      profile.surname = parts[0];
      profile.givenName = parts[1];
    }
  } else if (!profile.givenName && profile.displayName.includes(' ')) {
    // Format: "GivenName Surname"
    const parts = profile.displayName.split(' ');
    profile.givenName = parts[0];
    profile.surname = parts.slice(1).join(' ');
  }
  
  return profile;
}

/**
 * Calculates token expiry status from an expiry timestamp.
 * 
 * @param expiryMs - Token expiry time in milliseconds since epoch
 * @param nowMs - Current time in milliseconds (for testing)
 * @returns Token status including whether it's valid and time remaining
 */
export function calculateTokenStatus(
  expiryMs: number,
  nowMs: number = Date.now()
): {
  isValid: boolean;
  expiresAt: string;
  minutesRemaining: number;
} {
  const expiryDate = new Date(expiryMs);
  
  return {
    isValid: expiryMs > nowMs,
    expiresAt: expiryDate.toISOString(),
    minutesRemaining: Math.max(0, Math.round((expiryMs - nowMs) / 1000 / 60)),
  };
}

/**
 * Parses the pagination result from a search API response.
 * 
 * @param entitySets - Raw EntitySets array from API response
 * @returns Parsed results and total count if available
 */
export function parseSearchResults(
  entitySets: unknown[] | undefined,
  linkContext?: LinkContext
): { results: TeamsSearchResult[]; total?: number } {
  const results: TeamsSearchResult[] = [];
  let total: number | undefined;

  if (!Array.isArray(entitySets)) {
    return { results, total };
  }

  for (const entitySet of entitySets) {
    const es = entitySet as Record<string, unknown>;
    const resultSets = es.ResultSets as unknown[] | undefined;
    
    if (Array.isArray(resultSets)) {
      for (const resultSet of resultSets) {
        const rs = resultSet as Record<string, unknown>;
        
        // Try to get total
        const rsTotal = rs.Total ?? rs.TotalCount ?? rs.TotalEstimate;
        if (typeof rsTotal === 'number') {
          total = rsTotal;
        }
        
        const items = rs.Results as unknown[] | undefined;
        if (Array.isArray(items)) {
          for (const item of items) {
            const parsed = parseV2Result(item as Record<string, unknown>, linkContext);
            if (parsed) results.push(parsed);
          }
        }
      }
    }
  }

  return { results, total };
}

/**
 * Parses people search results from the Groups/Suggestions structure.
 * 
 * @param groups - Raw Groups array from suggestions API response
 * @returns Array of parsed person results
 */
export function parsePeopleResults(groups: unknown[] | undefined): PersonSearchResult[] {
  const results: PersonSearchResult[] = [];
  
  if (!Array.isArray(groups)) {
    return results;
  }

  for (const group of groups) {
    const g = group as Record<string, unknown>;
    const suggestions = g.Suggestions as unknown[] | undefined;
    
    if (Array.isArray(suggestions)) {
      for (const suggestion of suggestions) {
        const parsed = parsePersonSuggestion(suggestion as Record<string, unknown>);
        if (parsed) results.push(parsed);
      }
    }
  }

  return results;
}

/** Channel search result from Substrate suggestions API or Teams List API. */
export interface ChannelSearchResult {
  channelId: string;         // Conversation ID (19:xxx@thread.tacv2)
  channelName: string;       // Channel display name
  teamName: string;          // Parent team name
  teamId: string;            // Team group ID
  channelType: string;       // "Standard", "Private", etc.
  description?: string;      // Channel description if available
  isMember?: boolean;        // Whether user is a member of this channel's team
}

/**
 * Parses a single channel suggestion from the API response.
 * 
 * @param suggestion - Raw suggestion object from API
 * @returns Parsed channel result or null if required fields are missing
 */
export function parseChannelSuggestion(
  suggestion: Record<string, unknown>
): ChannelSearchResult | null {
  const name = suggestion.Name as string | undefined;
  const threadId = suggestion.ThreadId as string | undefined;
  const teamName = suggestion.TeamName as string | undefined;
  const groupId = suggestion.GroupId as string | undefined;
  
  // All required fields must be present
  if (!name || !threadId || !teamName || !groupId) {
    return null;
  }

  return {
    channelId: threadId,
    channelName: name,
    teamName,
    teamId: groupId,
    channelType: (suggestion.ChannelType as string) || 'Standard',
    description: suggestion.Description as string | undefined,
  };
}

/**
 * Parses channel search results from the Groups/Suggestions structure.
 * 
 * @param groups - Raw Groups array from suggestions API response
 * @returns Array of parsed channel results
 */
export function parseChannelResults(groups: unknown[] | undefined): ChannelSearchResult[] {
  const results: ChannelSearchResult[] = [];
  
  if (!Array.isArray(groups)) {
    return results;
  }

  for (const group of groups) {
    const g = group as Record<string, unknown>;
    const suggestions = g.Suggestions as unknown[] | undefined;
    
    if (Array.isArray(suggestions)) {
      for (const suggestion of suggestions) {
        const s = suggestion as Record<string, unknown>;
        // Only parse ChannelSuggestion entities
        if (s.EntityType === 'ChannelSuggestion') {
          const parsed = parseChannelSuggestion(s);
          if (parsed) results.push(parsed);
        }
      }
    }
  }

  return results;
}

/** Team with channels from the Teams List API response. */
export interface TeamWithChannels {
  teamId: string;           // Team group ID (GUID)
  teamName: string;         // Team display name
  threadId: string;         // Team root conversation ID
  description?: string;     // Team description
  channels: ChannelSearchResult[];
}

/**
 * Parses the Teams List API response to extract all teams and channels.
 * 
 * @param data - Raw response data from /api/csa/{region}/api/v3/teams/users/me
 * @returns Array of teams with their channels
 */
export function parseTeamsList(data: Record<string, unknown> | undefined): TeamWithChannels[] {
  const results: TeamWithChannels[] = [];
  
  if (!data) return results;
  
  const teams = data.teams as unknown[] | undefined;
  if (!Array.isArray(teams)) return results;
  
  for (const team of teams) {
    const t = team as Record<string, unknown>;
    // Team's id IS the thread ID (format: 19:xxx@thread.tacv2)
    const threadId = t.id as string | undefined;
    const displayName = t.displayName as string | undefined;
    
    if (!threadId || !displayName) continue;
    
    const channels: ChannelSearchResult[] = [];
    const channelList = t.channels as unknown[] | undefined;
    
    if (Array.isArray(channelList)) {
      for (const channel of channelList) {
        const c = channel as Record<string, unknown>;
        const channelId = c.id as string | undefined;
        const channelName = c.displayName as string | undefined;
        
        if (!channelId || !channelName) continue;
        
        // Channel has groupId directly, and channelType as a number
        const groupId = (c.groupId as string) || '';
        // Map numeric channelType to string (0=Standard, 1=Private, 2=Shared)
        const channelTypeNum = c.channelType as number | undefined;
        const channelType = channelTypeNum === 1 ? 'Private' 
          : channelTypeNum === 2 ? 'Shared' 
          : 'Standard';
        
        channels.push({
          channelId,
          channelName,
          teamName: displayName,
          teamId: groupId,
          channelType,
          description: c.description as string | undefined,
          isMember: true, // User is always a member for channels returned by this API
        });
      }
    }
    
    results.push({
      teamId: threadId, // Use thread ID as team identifier
      teamName: displayName,
      threadId,
      description: t.description as string | undefined,
      channels,
    });
  }
  
  return results;
}

/**
 * Filters channels from the Teams List by name.
 * 
 * @param teams - Array of teams with channels from parseTeamsList
 * @param query - Search query (case-insensitive partial match)
 * @returns Matching channels flattened into a single array
 */
export function filterChannelsByName(
  teams: TeamWithChannels[],
  query: string
): ChannelSearchResult[] {
  const lowerQuery = query.toLowerCase();
  const results: ChannelSearchResult[] = [];
  
  for (const team of teams) {
    for (const channel of team.channels) {
      if (channel.channelName.toLowerCase().includes(lowerQuery)) {
        results.push(channel);
      }
    }
  }
  
  return results;
}

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
// Markdown to Teams HTML Conversion
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Escapes HTML special characters in text.
 */
function escapeHtmlChars(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Converts inline markdown formatting to Teams HTML within a single line.
 * Handles: bold, italic, strikethrough, inline code.
 * Text outside of formatting markers is HTML-escaped.
 */
function convertInlineFormatting(line: string): string {
  // Process inline code first (to prevent other formatting inside code spans)
  // Split on `code` patterns, escape and format alternately
  const codeParts = line.split(/`([^`]+)`/);
  let result = '';
  
  for (let i = 0; i < codeParts.length; i++) {
    if (i % 2 === 1) {
      // Inside backticks - render as code, only escape HTML
      result += `<code>${escapeHtmlChars(codeParts[i])}</code>`;
    } else {
      // Outside backticks - process other inline formatting
      let segment = escapeHtmlChars(codeParts[i]);
      
      // Bold: **text** or __text__
      segment = segment.replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
      segment = segment.replace(/__(.+?)__/g, '<b>$1</b>');
      
      // Italic: *text* or _text_ (but not inside words for underscore)
      segment = segment.replace(/\*(.+?)\*/g, '<i>$1</i>');
      segment = segment.replace(/(?<!\w)_(.+?)_(?!\w)/g, '<i>$1</i>');
      
      // Strikethrough: ~~text~~
      segment = segment.replace(/~~(.+?)~~/g, '<s>$1</s>');
      
      result += segment;
    }
  }
  
  return result;
}

/**
 * Converts markdown-formatted text to Teams-compatible HTML.
 * 
 * Supports:
 * - **bold** / __bold__ → <b>
 * - *italic* / _italic_ → <i>
 * - ~~strikethrough~~ → <s>
 * - `inline code` → <code>
 * - ```code blocks``` → <pre><code>
 * - Newlines → paragraph breaks
 * - Ordered lists (1. item) → <ol><li>
 * - Unordered lists (- item, * item) → <ul><li>
 * 
 * Plain text without any formatting is returned as-is (HTML-escaped).
 */
export function markdownToTeamsHtml(text: string): string {
  // Handle fenced code blocks first (```...```)
  // Split text into code blocks and non-code-block segments
  const segments: { type: 'text' | 'codeblock'; content: string; lang?: string }[] = [];
  const codeBlockRegex = /```(\w*)\n?([\s\S]*?)```/g;
  let lastIndex = 0;
  let match;
  
  while ((match = codeBlockRegex.exec(text)) !== null) {
    // Text before this code block
    if (match.index > lastIndex) {
      segments.push({ type: 'text', content: text.substring(lastIndex, match.index) });
    }
    segments.push({ type: 'codeblock', content: match[2], lang: match[1] || undefined });
    lastIndex = match.index + match[0].length;
  }
  // Remaining text after last code block
  if (lastIndex < text.length) {
    segments.push({ type: 'text', content: text.substring(lastIndex) });
  }
  
  const htmlParts: string[] = [];
  
  for (const segment of segments) {
    if (segment.type === 'codeblock') {
      // Code blocks: escape HTML but preserve whitespace
      const escaped = escapeHtmlChars(segment.content.replace(/\n$/, ''));
      htmlParts.push(`<pre><code>${escaped}</code></pre>`);
      continue;
    }
    
    // Process text segments: split into paragraphs on double newlines
    const paragraphs = segment.content.split(/\n{2,}/);
    
    for (const para of paragraphs) {
      const trimmed = para.trim();
      if (!trimmed) continue;
      
      const lines = trimmed.split('\n');
      
      // Check if this paragraph is a list
      const isUnorderedList = lines.every(l => /^\s*[-*]\s+/.test(l));
      const isOrderedList = lines.every(l => /^\s*\d+[.)]\s+/.test(l));
      
      if (isUnorderedList) {
        const items = lines.map(l => {
          const content = l.replace(/^\s*[-*]\s+/, '');
          return `<li>${convertInlineFormatting(content)}</li>`;
        });
        htmlParts.push(`<ul>${items.join('')}</ul>`);
      } else if (isOrderedList) {
        const items = lines.map(l => {
          const content = l.replace(/^\s*\d+[.)]\s+/, '');
          return `<li>${convertInlineFormatting(content)}</li>`;
        });
        htmlParts.push(`<ol>${items.join('')}</ol>`);
      } else {
        // Regular paragraph - join lines with <br>
        const htmlLines = lines.map(l => convertInlineFormatting(l));
        htmlParts.push(`<p>${htmlLines.join('<br>')}</p>`);
      }
    }
  }
  
  return htmlParts.join('') || '<p></p>';
}

/**
 * Checks whether text contains any markdown formatting that would
 * benefit from conversion to HTML.
 */
export function hasMarkdownFormatting(text: string): boolean {
  // Code blocks
  if (/```[\s\S]*```/.test(text)) return true;
  // Inline code
  if (/`[^`]+`/.test(text)) return true;
  // Bold
  if (/\*\*.+?\*\*/.test(text) || /__.+?__/.test(text)) return true;
  // Italic (single * or _)
  if (/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/.test(text)) return true;
  // Strikethrough
  if (/~~.+?~~/.test(text)) return true;
  // Lists
  if (/^\s*[-*]\s+/m.test(text)) return true;
  if (/^\s*\d+[.)]\s+/m.test(text)) return true;
  // Multiple newlines (paragraph breaks)
  if (/\n/.test(text)) return true;
  
  return false;
}

// ─────────────────────────────────────────────────────────────────────────────
// Virtual Conversation Parsing
// ─────────────────────────────────────────────────────────────────────────────

/** Common fields from a virtual conversation message (48:saved, 48:threads, etc). */
export interface VirtualConversationItem {
  id: string;
  content: string;
  contentType: string;
  sender: {
    mri: string;
    displayName?: string;
  };
  timestamp: string;
  sourceConversationId: string;
  sourceReferenceId?: string;
  messageLink?: string;
  links?: ExtractedLink[];
}

/**
 * Parses a raw message from a virtual conversation (48:saved, 48:threads, etc).
 * 
 * Virtual conversations contain references to messages in other conversations.
 * The clumpId field contains the source conversation ID, and secondaryReferenceId
 * contains a composite key with the source message/post ID.
 * 
 * @param msg - Raw message object from virtual conversation API
 * @param referencePattern - Regex to extract source ID from secondaryReferenceId
 * @returns Parsed virtual conversation item, or null if message should be skipped
 */
export function parseVirtualConversationMessage(
  msg: Record<string, unknown>,
  referencePattern: RegExp,
  linkContext?: LinkContext
): VirtualConversationItem | null {
  // Skip non-message types
  const messageType = msg.messagetype as string || msg.type as string;
  if (!messageType || messageType.startsWith('Control/')) {
    return null;
  }

  const id = msg.id as string;
  if (!id) return null;

  const content = msg.content as string || '';
  const contentType = messageType || 'Text';

  const fromMri = msg.from as string || '';
  const displayName = msg.imdisplayname as string || msg.displayName as string;

  // Safe timestamp extraction - use extractActivityTimestamp pattern
  const timestamp = extractActivityTimestamp(msg);
  if (!timestamp) return null;

  // clumpId contains the original conversation where the message lives
  const sourceConversationId = msg.clumpId as string || '';
  
  // Extract source reference ID from secondaryReferenceId if available
  let sourceReferenceId: string | undefined;
  const secondaryRef = msg.secondaryReferenceId as string;
  if (secondaryRef) {
    const match = secondaryRef.match(referencePattern);
    if (match) {
      sourceReferenceId = match[1];
    }
  }

  // Build message link to original message
  const messageLink = sourceConversationId && sourceReferenceId
    ? buildMessageLink({
        conversationId: sourceConversationId,
        messageId: sourceReferenceId,
        tenantId: linkContext?.tenantId,
        teamsBaseUrl: linkContext?.teamsBaseUrl,
      })
    : undefined;

  // Extract links before stripping HTML
  const links = extractLinks(content);

  return {
    id,
    content: stripHtml(content),
    contentType,
    sender: {
      mri: fromMri,
      displayName,
    },
    timestamp,
    sourceConversationId,
    sourceReferenceId,
    messageLink,
    links: links.length > 0 ? links : undefined,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Transcript Formatting
// ─────────────────────────────────────────────────────────────────────────────

/** A single entry from a meeting transcript. */
export interface TranscriptEntry {
  /** Start time (e.g., "00:00:22.287"). */
  startTime: string;
  /** End time (e.g., "00:00:23.167"). */
  endTime: string;
  /** Speaker display name. */
  speaker: string;
  /** Spoken text content. */
  text: string;
}

/**
 * Formats transcript entries into a readable text format.
 * 
 * Merges consecutive entries from the same speaker into a single block
 * to reduce noise and improve readability.
 * 
 * @param entries - Transcript entries
 * @returns Formatted transcript string
 */
export function formatTranscriptText(entries: TranscriptEntry[]): string {
  if (entries.length === 0) return '';
  
  const blocks: string[] = [];
  let currentSpeaker: string | null = null;
  let currentTexts: string[] = [];
  let blockStartTime = '';
  
  for (const entry of entries) {
    if (entry.speaker !== currentSpeaker) {
      // Flush previous block
      if (currentTexts.length > 0) {
        const prefix = currentSpeaker
          ? `[${blockStartTime}] ${currentSpeaker}:`
          : `[${blockStartTime}]`;
        blocks.push(`${prefix}\n${currentTexts.join(' ')}`);
      }
      currentSpeaker = entry.speaker;
      currentTexts = [entry.text];
      blockStartTime = entry.startTime;
    } else {
      currentTexts.push(entry.text);
    }
  }
  
  // Flush last block
  if (currentTexts.length > 0) {
    const prefix = currentSpeaker
      ? `[${blockStartTime}] ${currentSpeaker}:`
      : `[${blockStartTime}]`;
    blocks.push(`${prefix}\n${currentTexts.join(' ')}`);
  }
  
  return blocks.join('\n\n');
}
