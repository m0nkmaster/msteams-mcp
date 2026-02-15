/**
 * Pure parsing functions for Teams API responses.
 * 
 * This barrel module re-exports all parser submodules for backward compatibility.
 * Implementations are split into focused submodules:
 * - parsers-html: HTML stripping, link extraction, escaping
 * - parsers-markdown: Markdown to Teams HTML conversion
 * - parsers-identifiers: Message links, conversation types, ID extraction, timestamps
 * - parsers-search: Teams message and email search result parsing
 * - parsers-people: People search, JWT profiles, token status
 * - parsers-channels: Channel and team list parsing
 * - parsers-virtual: Virtual conversation messages and transcript formatting
 */

export { extractLinks, stripHtml, escapeHtmlChars, type ExtractedLink } from './parsers-html.js';
export { markdownToTeamsHtml, hasMarkdownFormatting } from './parsers-markdown.js';
export {
  getConversationType,
  buildMessageLink,
  extractMessageTimestamp,
  extractActivityTimestamp,
  decodeBase64Guid,
  extractObjectId,
  buildOneOnOneConversationId,
  type MessageLinkOptions,
  type LinkContext,
} from './parsers-identifiers.js';
export {
  parseV2Result,
  parseSearchResults,
  classifyEmailType,
  parseEmailResult,
  parseEmailSearchResults,
} from './parsers-search.js';
export {
  parsePersonSuggestion,
  parsePeopleResults,
  parseJwtProfile,
  calculateTokenStatus,
  type PersonSearchResult,
  type UserProfile,
} from './parsers-people.js';
export {
  parseChannelSuggestion,
  parseChannelResults,
  parseTeamsList,
  filterChannelsByName,
  type ChannelSearchResult,
  type TeamWithChannels,
} from './parsers-channels.js';
export {
  parseVirtualConversationMessage,
  formatTranscriptText,
  type VirtualConversationItem,
  type TranscriptEntry,
} from './parsers-virtual.js';
