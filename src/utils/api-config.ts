/**
 * API endpoint configuration and header utilities.
 * 
 * Centralises all API URLs and common request headers.
 * 
 * Region and base URLs are extracted from the user's session via DISCOVER-REGION-GTM,
 * supporting different Teams environments (commercial, GCC, GCC-High, etc.).
 */

import { NOTIFICATIONS_ID } from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Default URLs (fallbacks when session config unavailable)
// ─────────────────────────────────────────────────────────────────────────────

/** Default Teams base URL (commercial cloud). */
export const DEFAULT_TEAMS_BASE_URL = 'https://teams.microsoft.com';

/** Default Substrate base URL (commercial cloud). */
export const DEFAULT_SUBSTRATE_BASE_URL = 'https://substrate.office.com';

// ─────────────────────────────────────────────────────────────────────────────
// Substrate API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Substrate API endpoints.
 * 
 * Note: Substrate URLs may differ for GCC/GCC-High tenants. The base URL is
 * hardcoded to substrate.office.com as we haven't found a config source for it.
 * If issues are reported from government cloud users, we may need to add config.
 */
export const SUBSTRATE_API = {
  /** Full-text message search. */
  search: `${DEFAULT_SUBSTRATE_BASE_URL}/searchservice/api/v2/query`,
  
  /** People and message typeahead suggestions. */
  suggestions: `${DEFAULT_SUBSTRATE_BASE_URL}/search/api/v1/suggestions`,
  
  /** Frequent contacts list. */
  frequentContacts: `${DEFAULT_SUBSTRATE_BASE_URL}/search/api/v1/suggestions?scenario=peoplecache`,
  
  /** People search. */
  peopleSearch: `${DEFAULT_SUBSTRATE_BASE_URL}/search/api/v1/suggestions?scenario=powerbar`,
  
  /** Channel search (org-wide, not just user's teams). */
  channelSearch: `${DEFAULT_SUBSTRATE_BASE_URL}/search/api/v1/suggestions?scenario=powerbar&setflight=TurnOffMPLSuppressionTeams,EnableTeamsChannelDomainPowerbar&domain=TeamsChannel`,
} as const;

// ─────────────────────────────────────────────────────────────────────────────
// Chatsvc API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Chat service API endpoint builders.
 * 
 * All functions accept an optional `baseUrl` parameter to support different
 * Teams environments. If not provided, uses DEFAULT_TEAMS_BASE_URL.
 */
export const CHATSVC_API = {
  /**
   * Get messages URL for a conversation.
   * 
   * For thread replies in channels, provide replyToMessageId to append
   * `;messageid={id}` to the conversation path. This tells Teams the message
   * is a reply to an existing thread rather than a new top-level post.
   */
  messages: (region: string, conversationId: string, replyToMessageId?: string, baseUrl = DEFAULT_TEAMS_BASE_URL) => {
    const conversationPath = replyToMessageId
      ? `${conversationId};messageid=${replyToMessageId}`
      : conversationId;
    return `${baseUrl}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationPath)}/messages`;
  },
  
  /** Get conversation metadata URL. */
  conversation: (region: string, conversationId: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}`,
  
  /** Save/unsave message metadata URL. */
  messageMetadata: (region: string, conversationId: string, messageId: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/rcmetadata/${messageId}`,
  
  /** Edit a specific message URL. */
  editMessage: (region: string, conversationId: string, messageId: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${messageId}`,
  
  /** Delete a specific message URL (soft delete). */
  deleteMessage: (region: string, conversationId: string, messageId: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${messageId}?behavior=softDelete`,
  
  /** Get consumption horizons (read receipts) for a thread. */
  consumptionHorizons: (region: string, threadId: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/chatsvc/${region}/v1/threads/${encodeURIComponent(threadId)}/consumptionhorizons`,
  
  /** Update consumption horizon (mark as read) for a conversation. */
  updateConsumptionHorizon: (region: string, conversationId: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/properties?name=consumptionhorizon`,
  
  /** Activity feed (notifications) messages. */
  activityFeed: (region: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(NOTIFICATIONS_ID)}/messages`,
  
  /** Message emotions (reactions) URL. */
  messageEmotions: (region: string, conversationId: string, messageId: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/chatsvc/${region}/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages/${messageId}/properties?name=emotions`,
  
  /** Create a new thread (group chat). */
  createThread: (region: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/chatsvc/${region}/v1/threads`,
} as const;

// ─────────────────────────────────────────────────────────────────────────────
// Calendar API
// ─────────────────────────────────────────────────────────────────────────────

/** 
 * Calendar/Meeting API endpoints.
 * 
 * The mt/part endpoints use partitioned regions (e.g., amer-02, emea-03).
 * Some tenants use non-partitioned URLs (e.g., /api/mt/emea).
 * The correct format is extracted from the user's session via DISCOVER-REGION-GTM.
 */
export const CALENDAR_API = {
  /**
   * Get calendar view (meetings) for a date range.
   * 
   * Uses OData-style query parameters for filtering and pagination.
   * 
   * @param regionPartition - The full region-partition (e.g., "amer-02") or just region (e.g., "emea")
   * @param hasPartition - Whether the tenant uses partitioned URLs
   * @param baseUrl - Teams base URL (from config or default)
   */
  calendarView: (regionPartition: string, hasPartition: boolean, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    hasPartition
      ? `${baseUrl}/api/mt/part/${regionPartition}/v2.1/me/calendars/calendarView`
      : `${baseUrl}/api/mt/${regionPartition}/v2.1/me/calendars/calendarView`,
} as const;

// ─────────────────────────────────────────────────────────────────────────────
// CSA API
// ─────────────────────────────────────────────────────────────────────────────

/** CSA (Chat Service Aggregator) API endpoints. */
export const CSA_API = {
  /** Conversation folders (favourites) URL. */
  conversationFolders: (region: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/csa/${region}/api/v1/teams/users/me/conversationFolders?supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true`,
  
  /** Teams list (all teams/channels user is a member of). */
  teamsList: (region: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/csa/${region}/api/v3/teams/users/me?isPrefetch=false&enableMembershipSummary=true`,
  
  /** Custom emoji metadata. */
  customEmojis: (region: string, baseUrl = DEFAULT_TEAMS_BASE_URL) =>
    `${baseUrl}/api/csa/${region}/api/v1/customemoji/metadata`,
} as const;

// ─────────────────────────────────────────────────────────────────────────────
// Request Headers
// ─────────────────────────────────────────────────────────────────────────────

/** Common request headers for Teams API calls. */
export function getTeamsHeaders(baseUrl = DEFAULT_TEAMS_BASE_URL): HeadersInit {
  return {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'Origin': baseUrl,
    'Referer': `${baseUrl}/`,
  };
}

/** Headers for Bearer token authentication. */
export function getBearerHeaders(token: string, baseUrl = DEFAULT_TEAMS_BASE_URL): HeadersInit {
  return {
    ...getTeamsHeaders(baseUrl),
    'Authorization': `Bearer ${token}`,
  };
}

/** Headers for Skype token + Bearer authentication. */
export function getSkypeAuthHeaders(skypeToken: string, authToken: string, baseUrl = DEFAULT_TEAMS_BASE_URL): HeadersInit {
  return {
    ...getTeamsHeaders(baseUrl),
    'Authentication': `skypetoken=${skypeToken}`,
    'Authorization': `Bearer ${authToken}`,
  };
}

/** Headers for CSA API (Skype token + CSA Bearer). */
export function getCsaHeaders(skypeToken: string, csaToken: string, baseUrl = DEFAULT_TEAMS_BASE_URL): HeadersInit {
  return {
    ...getTeamsHeaders(baseUrl),
    'Authentication': `skypetoken=${skypeToken}`,
    'Authorization': `Bearer ${csaToken}`,
  };
}

/** Client version header for messaging API. */
export function getMessagingHeaders(skypeToken: string, authToken: string, baseUrl = DEFAULT_TEAMS_BASE_URL): HeadersInit {
  return {
    ...getSkypeAuthHeaders(skypeToken, authToken, baseUrl),
    'X-Ms-Client-Version': '1415/1.0.0.2025010401',
  };
}
