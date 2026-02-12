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


