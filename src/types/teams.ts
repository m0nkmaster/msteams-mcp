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

export interface TeamsMessage {
  id: string;
  content: string;
  sender: string;
  timestamp: string;
  channelName?: string;
  teamName?: string;
  replyCount?: number;
  reactions?: TeamsReaction[];
}

export interface TeamsReaction {
  type: string;
  count: number;
}

export interface InterceptedRequest {
  url: string;
  method: string;
  headers: Record<string, string>;
  postData?: string;
  timestamp: Date;
}

export interface InterceptedResponse {
  url: string;
  status: number;
  headers: Record<string, string>;
  body?: string;
  timestamp: Date;
}

// Note: SessionState is defined in auth/session-store.ts (the authoritative source)
// Import from there if needed: import { type SessionState } from '../auth/session-store.js';

export interface SearchApiEndpoint {
  url: string;
  method: string;
  description: string;
  requestFormat?: unknown;
  responseFormat?: unknown;
}

/** Pagination options for search requests. */
export interface SearchPaginationOptions {
  /** Starting offset (0-based). Default: 0 */
  from?: number;
  /** Page size. Default: 25 */
  size?: number;
  /** Maximum total results to fetch across all pages. */
  maxResults?: number;
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

/** Search results with pagination metadata. */
export interface TeamsSearchResultsWithPagination {
  results: TeamsSearchResult[];
  pagination: SearchPaginationResult;
}
