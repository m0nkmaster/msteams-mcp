/**
 * TypeScript interfaces for Teams data structures.
 */

/** A link extracted from message content. */
export interface ExtractedLink {
  url: string;
  text: string;
  /** Link type from <item type="..."> elements (e.g., amsTranscript, onedriveForBusinessVideo). */
  contentType?: string;
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
  nextFrom?: number;
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

/** Full email detail from Microsoft Graph /me/messages/{id}. */
export interface EmailDetail {
  /** Graph message id. */
  id: string;
  /** Email subject line. */
  subject: string;
  /** Sender display name. */
  sender: string;
  /** Sender email address. */
  senderEmail?: string;
  /** To recipients, formatted "Name <address>". */
  toRecipients?: string[];
  /** CC recipients, formatted "Name <address>". */
  ccRecipients?: string[];
  /** BCC recipients, formatted "Name <address>". */
  bccRecipients?: string[];
  /** When the email was sent (ISO timestamp). */
  sentAt?: string;
  /** When the email was received (ISO timestamp). */
  receivedAt?: string;
  /** Importance level (normal, high, low). */
  importance?: string;
  /** Whether the email has been read. */
  isRead?: boolean;
  /** Web link to open the email in Outlook. */
  webLink?: string;
  /** The conversation/thread ID for this email chain. */
  conversationId?: string;
  /** RFC 5322 message id. */
  internetMessageId?: string;
  /** Classification: 'email', 'calendar-response', or 'notification'. */
  emailType: EmailType;
  /** Full email body as plain text. */
  body: string;
  /** Links extracted from the HTML body. */
  links?: ExtractedLink[];
  /** Non-inline attachments (metadata only; use the email's webLink to download). */
  attachments?: { name: string; contentType?: string; size?: number }[];
}
