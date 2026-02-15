/**
 * Search result and email result parsing from Substrate API responses.
 */

import type { TeamsSearchResult, EmailSearchResult, EmailType } from '../types/teams.js';
import { MIN_CONTENT_LENGTH } from '../constants.js';
import { extractLinks, stripHtml } from './parsers-html.js';
import { extractMessageTimestamp, buildMessageLink, type LinkContext } from './parsers-identifiers.js';

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

// ─────────────────────────────────────────────────────────────────────────────
// Email Parsing
// ─────────────────────────────────────────────────────────────────────────────

/** Subject prefixes that indicate calendar responses (case-insensitive). */
const CALENDAR_RESPONSE_PREFIXES = [
  'accepted:',
  'declined:',
  'tentative:',
  'canceled:',
  'cancelled:',
];

/** Subject prefixes that indicate automated notifications (case-insensitive). */
const NOTIFICATION_PREFIXES = [
  'automatic reply:',
  'auto-reply:',
  'out of office:',
];

/**
 * Classifies an email result based on its subject line.
 */
export function classifyEmailType(subject: string): EmailType {
  const lower = subject.toLowerCase().trimStart();
  for (const prefix of CALENDAR_RESPONSE_PREFIXES) {
    if (lower.startsWith(prefix)) return 'calendar-response';
  }
  for (const prefix of NOTIFICATION_PREFIXES) {
    if (lower.startsWith(prefix)) return 'notification';
  }
  return 'email';
}

/**
 * Parses a single email result from the Substrate v2 query API response.
 * 
 * Email results have a different structure to Teams message results:
 * - Subject is a top-level field
 * - From/To/Cc are structured address objects
 * - Body preview is in HitHighlightedSummary or Preview
 */
export function parseEmailResult(item: Record<string, unknown>): EmailSearchResult | null {
  const source = item.Source as Record<string, unknown> | undefined;

  // Subject from top-level or source
  const subject = item.Subject as string
    || source?.Subject as string
    || '';

  // Preview/body snippet
  const rawPreview = item.HitHighlightedSummary as string
    || item.Summary as string
    || source?.Preview as string
    || '';

  const links = extractLinks(rawPreview);
  const preview = stripHtml(rawPreview);

  // Skip results with no meaningful content
  if (!subject && preview.length < MIN_CONTENT_LENGTH) return null;

  const id = item.Id as string
    || item.ReferenceId as string
    || `email-${Date.now()}`;

  // Sender info — Substrate returns From as a structured object or string
  let sender = '';
  let senderEmail: string | undefined;
  const from = source?.From as Record<string, unknown> | string | undefined;
  if (typeof from === 'string') {
    sender = from;
  } else if (from) {
    const emailAddress = from.EmailAddress as Record<string, unknown> | undefined;
    if (emailAddress) {
      sender = emailAddress.Name as string || '';
      senderEmail = emailAddress.Address as string | undefined;
    } else {
      sender = from.Name as string || from.DisplayName as string || '';
      senderEmail = from.Address as string || from.EmailAddress as string | undefined;
    }
  }
  // Fallback to flat fields
  if (!sender) {
    sender = source?.FromName as string || source?.Sender as string || '';
  }
  if (!senderEmail) {
    senderEmail = source?.FromAddress as string || source?.SenderSmtpAddress as string || undefined;
  }

  // Timestamp
  const receivedAt = source?.DateTimeReceived as string
    || source?.DateTimeSent as string
    || source?.ReceivedTime as string
    || undefined;

  // Attachments
  let hasAttachments: boolean | undefined;
  if (typeof source?.HasAttachments === 'boolean') {
    hasAttachments = source.HasAttachments;
  } else if (typeof source?.HasAttachment === 'boolean') {
    hasAttachments = source.HasAttachment;
  }

  // Importance
  const importance = source?.Importance as string || undefined;

  // Read status
  const isRead = typeof source?.IsRead === 'boolean' ? source.IsRead as boolean : undefined;

  // Recipients
  const toRecipients = parseRecipients(source?.ToRecipients ?? source?.DisplayTo);
  const ccRecipients = parseRecipients(source?.CcRecipients ?? source?.DisplayCc);

  // Web link
  const webLink = source?.WebLink as string
    || source?.OwaLink as string
    || undefined;

  // Conversation ID (email thread) — can be a string or {Id: string} object
  let conversationId: string | undefined;
  const rawConvId = source?.ConversationId;
  if (typeof rawConvId === 'string') {
    conversationId = rawConvId;
  } else if (rawConvId && typeof rawConvId === 'object') {
    conversationId = (rawConvId as Record<string, unknown>).Id as string | undefined;
  }
  if (!conversationId) {
    conversationId = source?.InternetMessageId as string || undefined;
  }

  return {
    id,
    emailType: classifyEmailType(subject),
    subject,
    sender,
    senderEmail,
    preview,
    receivedAt,
    hasAttachments,
    importance,
    isRead,
    toRecipients,
    ccRecipients,
    webLink,
    conversationId,
    links: links.length > 0 ? links : undefined,
  };
}

/**
 * Parses recipient fields from Substrate email results.
 * 
 * Recipients can be:
 * - A string (e.g., "John Smith; Jane Doe")
 * - An array of objects with Name/Address fields
 * - undefined
 */
function parseRecipients(recipients: unknown): string[] | undefined {
  if (!recipients) return undefined;

  if (typeof recipients === 'string') {
    // Semicolon-separated display names
    const parsed = recipients.split(';').map(r => r.trim()).filter(Boolean);
    return parsed.length > 0 ? parsed : undefined;
  }

  if (Array.isArray(recipients)) {
    const parsed: string[] = [];
    for (const r of recipients) {
      const recipient = r as Record<string, unknown>;
      const emailAddress = recipient.EmailAddress as Record<string, unknown> | undefined;
      if (emailAddress) {
        const name = emailAddress.Name as string;
        const address = emailAddress.Address as string;
        const display = name || address;
        if (display) parsed.push(display);
      } else {
        const name = recipient.Name as string || recipient.DisplayName as string;
        const address = recipient.Address as string || recipient.EmailAddress as string;
        const display = name || address;
        if (display) parsed.push(display);
      }
    }
    return parsed.length > 0 ? parsed : undefined;
  }

  return undefined;
}

/**
 * Parses email search results from the Substrate v2 query API response.
 * 
 * Same EntitySets/ResultSets structure as Teams message search, but
 * items are email messages with different fields.
 * 
 * By default, calendar responses (Accepted/Declined/Tentative/Canceled)
 * are filtered out. Set `excludeCalendarResponses: false` to include them.
 */
export function parseEmailSearchResults(
  entitySets: unknown[] | undefined,
  options: { excludeCalendarResponses?: boolean } = {},
): { results: EmailSearchResult[]; total?: number; filteredCount?: number } {
  const excludeCalendar = options.excludeCalendarResponses ?? true;
  const results: EmailSearchResult[] = [];
  let total: number | undefined;
  let filteredCount = 0;

  if (!Array.isArray(entitySets)) {
    return { results, total };
  }

  for (const entitySet of entitySets) {
    const es = entitySet as Record<string, unknown>;
    const resultSets = es.ResultSets as unknown[] | undefined;

    if (Array.isArray(resultSets)) {
      for (const resultSet of resultSets) {
        const rs = resultSet as Record<string, unknown>;

        const rsTotal = rs.Total ?? rs.TotalCount ?? rs.TotalEstimate;
        if (typeof rsTotal === 'number') {
          total = rsTotal;
        }

        const items = rs.Results as unknown[] | undefined;
        if (Array.isArray(items)) {
          for (const item of items) {
            const parsed = parseEmailResult(item as Record<string, unknown>);
            if (parsed) {
              if (excludeCalendar && parsed.emailType === 'calendar-response') {
                filteredCount++;
              } else {
                results.push(parsed);
              }
            }
          }
        }
      }
    }
  }

  return { results, total, filteredCount: filteredCount > 0 ? filteredCount : undefined };
}
