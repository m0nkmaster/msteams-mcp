/**
 * TypeScript interfaces for Teams data structures.
 */

/** A link extracted from message content. */
export interface ExtractedLink {
  url: string;
  text: string;
}

export interface TeamsSearchResult {
  id: string;
  type: 'message' | 'file' | 'person';
  content: string;
  sender?: string;
  timestamp?: string;
  channelName?: string;
  teamName?: string;
  conversationId?: string;
  messageId?: string;
  /** Direct link to open this message in Teams */
  messageLink?: string;
  /** Links extracted from message content */
  links?: ExtractedLink[];
}

/** Pagination metadata returned with search results. */
export interface SearchPaginationResult {
  /** Number of results returned in this response. */
  returned: number;
  /** Starting offset used for this request. */
  from: number;
  /** Page size requested. */
  size: number;
  /** Total results available (if known). */
  total?: number;
  /** Whether more results are available. */
  hasMore: boolean;
}

/** Classification of an email result. */
export type EmailType = 'email' | 'calendar-response' | 'notification';

/** An email search result from Substrate. */
export interface EmailSearchResult {
  /** Unique ID for this result. */
  id: string;
  /** Classification: 'email', 'calendar-response', or 'notification'. */
  emailType: EmailType;
  /** Email subject line. */
  subject: string;
  /** Sender display name. */
  sender: string;
  /** Sender email address. */
  senderEmail?: string;
  /** Preview/snippet of the email body (HTML stripped). */
  preview: string;
  /** When the email was received (ISO timestamp). */
  receivedAt?: string;
  /** Whether the email has attachments. */
  hasAttachments?: boolean;
  /** Importance level (Normal, High, Low). */
  importance?: string;
  /** Whether the email has been read. */
  isRead?: boolean;
  /** To recipients (display names or emails). */
  toRecipients?: string[];
  /** CC recipients (display names or emails). */
  ccRecipients?: string[];
  /** Web link to open the email in Outlook. */
  webLink?: string;
  /** The conversation/thread ID for this email chain. */
  conversationId?: string;
  /** Links extracted from email body. */
  links?: ExtractedLink[];
}
