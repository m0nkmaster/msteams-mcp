/**
 * Token extraction from session state.
 * 
 * Extracts various authentication tokens from Playwright's saved session state.
 * Teams stores MSAL tokens in localStorage; we parse these to get bearer tokens
 * for various APIs (Substrate search, chatsvc messaging, etc.).
 */

import {
  readSessionState,
  readTokenCache,
  writeTokenCache,
  clearTokenCache,
  getTeamsOrigin,
  type SessionState,
  type TokenCache,
} from './session-store.js';
import { parseJwtProfile, type UserProfile } from '../utils/parsers.js';
import { MRI_TYPE_PREFIX, ORGID_PREFIX, MRI_ORGID_PREFIX, MAX_DEBUG_CONFIG_VALUE_LENGTH } from '../constants.js';

// ============================================================================
// JWT Utilities
// ============================================================================

/**
 * Decodes a JWT token's payload without verifying the signature.
 */
function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;
    return JSON.parse(Buffer.from(parts[1], 'base64').toString());
  } catch {
    return null;
  }
}

/**
 * Gets the expiry date from a JWT token's `exp` claim.
 */
function getJwtExpiry(token: string): Date | null {
  const payload = decodeJwtPayload(token);
  if (!payload?.exp || typeof payload.exp !== 'number') return null;
  return new Date(payload.exp * 1000);
}

/**
 * Checks if a string looks like a JWT (starts with 'ey').
 */
function isJwtToken(value: unknown): value is string {
  return typeof value === 'string' && value.startsWith('ey');
}

// ============================================================================
// Session Helpers
// ============================================================================

/**
 * Resolves session state and Teams origin in one call.
 * Many functions need both, so this reduces boilerplate.
 */
function getTeamsLocalStorage(state?: SessionState): Array<{ name: string; value: string }> | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  const teamsOrigin = getTeamsOrigin(sessionState);
  return teamsOrigin?.localStorage ?? null;
}

// ============================================================================
// Types
// ============================================================================

/** Substrate search token (for search/people APIs). */
export interface SubstrateTokenInfo {
  token: string;
  expiry: Date;
}

/** Teams chat API token (for chatsvc). */
export interface TeamsTokenInfo {
  token: string;
  expiry: Date;
  userMri: string;
}

/** Cookie-based auth for messaging APIs. */
export interface MessageAuthInfo {
  skypeToken: string;
  authToken: string;
  userMri: string;
}

// ============================================================================
// Token Extraction
// ============================================================================

/**
 * Extracts the Substrate search token from session state.
 * This token is used for search and people APIs.
 */
export function extractSubstrateToken(state?: SessionState): SubstrateTokenInfo | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  // Collect all valid Substrate tokens and pick the one with longest expiry
  let bestToken: SubstrateTokenInfo | null = null;

  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      
      // Look for Substrate search tokens by target scope
      // Match both old format (substrate.office.com/search/SubstrateSearch)
      // and new format (substrate.office.com/SubstrateSearch-Internal.ReadWrite)
      const target = entry.target as string | undefined;
      if (!target?.includes('substrate.office.com')) continue;
      if (!target.includes('SubstrateSearch')) continue;

      if (!isJwtToken(entry.secret)) continue;

      const expiry = getJwtExpiry(entry.secret);
      if (!expiry) continue;

      // Skip expired tokens
      if (expiry.getTime() <= Date.now()) continue;

      // Keep the token with longest remaining validity
      if (!bestToken || expiry.getTime() > bestToken.expiry.getTime()) {
        bestToken = { token: entry.secret, expiry };
      }
    } catch {
      continue;
    }
  }

  return bestToken;
}

// ============================================================================
// Cached Token Access
// ============================================================================

/**
 * Gets a valid Substrate token, either from cache or by extracting from session.
 */
export function getValidSubstrateToken(): string | null {
  // Try cache first
  const cache = readTokenCache();
  if (cache && cache.substrateTokenExpiry > Date.now()) {
    return cache.substrateToken;
  }

  // Extract from session
  const extracted = extractSubstrateToken();
  if (!extracted) return null;

  // Check if not expired
  if (extracted.expiry.getTime() <= Date.now()) {
    return null;
  }

  // Cache the token
  const newCache: TokenCache = {
    substrateToken: extracted.token,
    substrateTokenExpiry: extracted.expiry.getTime(),
    extractedAt: Date.now(),
  };
  writeTokenCache(newCache);

  return extracted.token;
}

/**
 * Checks if we have a valid Substrate token.
 */
export function hasValidSubstrateToken(): boolean {
  return getValidSubstrateToken() !== null;
}

/**
 * Gets Substrate token status for diagnostics.
 */
export function getSubstrateTokenStatus(): {
  hasToken: boolean;
  expiresAt?: string;
  minutesRemaining?: number;
} {
  const extracted = extractSubstrateToken();
  if (!extracted) {
    return { hasToken: false };
  }

  const now = Date.now();
  const expiryMs = extracted.expiry.getTime();

  return {
    hasToken: expiryMs > now,
    expiresAt: extracted.expiry.toISOString(),
    minutesRemaining: Math.max(0, Math.round((expiryMs - now) / 1000 / 60)),
  };
}

/** Candidate token found during extraction. */
interface TokenCandidate {
  token: string;
  expiry: Date;
  userMri?: string;
}

/**
 * Extracts the Teams chat API token from session state.
 * 
 * Teams stores multiple tokens for different services. We prefer:
 * 1. chatsvcagg.teams.microsoft.com (primary chat API)
 * 2. api.spaces.skype.com (fallback)
 */
export function extractTeamsToken(state?: SessionState): TeamsTokenInfo | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  let chatsvcCandidate: TokenCandidate | null = null;
  let skypeCandidate: TokenCandidate | null = null;
  let userMri: string | null = null;

  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      if (!entry.target || !isJwtToken(entry.secret)) continue;

      const payload = decodeJwtPayload(entry.secret);
      if (!payload?.exp || typeof payload.exp !== 'number') continue;

      const expiry = new Date(payload.exp * 1000);

      // Capture user MRI from any token's oid claim
      if (typeof payload.oid === 'string' && !userMri) {
        userMri = `${MRI_ORGID_PREFIX}${payload.oid}`;
      }

      // Track best candidate for each service
      if (entry.target.includes('chatsvcagg.teams.microsoft.com')) {
        if (!chatsvcCandidate || expiry > chatsvcCandidate.expiry) {
          chatsvcCandidate = { token: entry.secret, expiry };
        }
      } else if (entry.target.includes('api.spaces.skype.com')) {
        if (!skypeCandidate || expiry > skypeCandidate.expiry) {
          skypeCandidate = { token: entry.secret, expiry };
        }
      }
    } catch {
      continue;
    }
  }

  // Fallback: extract userMri from Substrate token if not found
  if (!userMri) {
    userMri = extractUserMriFromSubstrate(state);
  }

  // Prefer chatsvc, fall back to skype
  const best = chatsvcCandidate ?? skypeCandidate;
  if (!best || !userMri || best.expiry.getTime() <= Date.now()) {
    return null;
  }

  return { token: best.token, expiry: best.expiry, userMri };
}

/**
 * Extracts the Skype Spaces API token from session state.
 * 
 * This token is required for the calendar/meetings API (mt/part endpoints).
 * It has scope: https://api.spaces.skype.com/Authorization.ReadWrite
 */
export function extractSkypeSpacesToken(state?: SessionState): string | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  let bestCandidate: { token: string; expiry: Date } | null = null;

  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      if (!entry.target || !isJwtToken(entry.secret)) continue;

      // Look for api.spaces.skype.com token
      if (!entry.target.includes('api.spaces.skype.com')) continue;

      const payload = decodeJwtPayload(entry.secret);
      if (!payload?.exp || typeof payload.exp !== 'number') continue;

      const expiry = new Date(payload.exp * 1000);
      
      // Skip expired tokens
      if (expiry.getTime() <= Date.now()) continue;

      // Keep the one with the latest expiry
      if (!bestCandidate || expiry > bestCandidate.expiry) {
        bestCandidate = { token: entry.secret, expiry };
      }
    } catch {
      continue;
    }
  }

  return bestCandidate?.token ?? null;
}

/** Region configuration from Teams discovery. */
export interface RegionConfig {
  /** Base region (e.g., "amer", "emea", "apac") - used by chatsvc, csa APIs. */
  region: string;
  /** Partition number (e.g., "02", "01") - only needed for mt/part APIs. */
  partition: string;
  /** Full region with partition (e.g., "amer-02") - for mt/part APIs. */
  regionPartition: string;
  /** Whether this tenant uses partitioned mt/part URLs. */
  hasPartition: boolean;
  
  // Full service URLs from DISCOVER-REGION-GTM (use these directly)
  /** Full middleTier URL (e.g., "https://teams.microsoft.com/api/mt/part/amer-02"). */
  middleTierUrl: string;
  /** Chat service URL base (e.g., "https://teams.microsoft.com/api/chatsvc/amer"). */
  chatServiceUrl: string;
  /** CSA service URL base (e.g., "https://teams.microsoft.com/api/csa/amer"). */
  csaServiceUrl: string;
  
  // Extracted base URLs for constructing other endpoints
  /** Teams base URL (e.g., "https://teams.microsoft.com" or "https://teams.microsoft.us" for GCC). */
  teamsBaseUrl: string;
}

/**
 * Extracts the user's region and partition from the Teams discovery config.
 * 
 * Teams stores a DISCOVER-REGION-GTM config in localStorage that contains
 * region-specific URLs for all APIs. There are two formats:
 * 
 * **Partitioned (most Enterprise tenants):**
 * - middleTier: "https://teams.microsoft.com/api/mt/part/amer-02"
 * - chatServiceAfd: "https://teams.microsoft.com/api/chatsvc/amer"
 * 
 * **Non-partitioned (some tenants, e.g., UK):**
 * - middleTier: "https://teams.microsoft.com/api/mt/emea"
 * - chatServiceAfd: "https://teams.microsoft.com/api/chatsvc/uk"
 * 
 * We use the full URLs directly from config rather than reconstructing them,
 * which ensures compatibility with GCC/GCC-High tenants that may use different
 * base URLs (e.g., teams.microsoft.us).
 */
export function extractRegionConfig(state?: SessionState): RegionConfig | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  // Find the DISCOVER-REGION-GTM key
  for (const item of localStorage) {
    if (!item.name.includes('DISCOVER-REGION-GTM')) continue;

    try {
      const data = JSON.parse(item.value) as { item?: Record<string, string> };
      const middleTierUrl = data.item?.middleTier;
      const chatServiceUrl = data.item?.chatServiceAfd;
      const csaServiceUrl = data.item?.chatSvcAggAfd;
      
      if (!chatServiceUrl) continue;

      // Extract Teams base URL from any of the URLs (chatServiceAfd is reliable)
      let teamsBaseUrl = 'https://teams.microsoft.com'; // fallback
      try {
        const url = new URL(chatServiceUrl);
        teamsBaseUrl = `${url.protocol}//${url.host}`;
      } catch {
        // Use fallback
      }

      // Extract region from chatServiceAfd (e.g., /api/chatsvc/amer or /api/chatsvc/uk)
      const chatMatch = chatServiceUrl.match(/\/api\/chatsvc\/([a-z]+)$/);
      if (!chatMatch) continue;
      const region = chatMatch[1];

      // Try to extract partition from middleTier if it's partitioned
      // Format: /api/mt/part/amer-02 (partitioned) or /api/mt/emea (non-partitioned)
      let partition: string | undefined;
      let regionPartition: string | undefined;
      let hasPartition = false;
      
      if (middleTierUrl) {
        const partitionMatch = middleTierUrl.match(/\/api\/mt\/part\/([a-z]+)-(\d+)$/);
        if (partitionMatch) {
          hasPartition = true;
          partition = partitionMatch[2];
          regionPartition = `${partitionMatch[1]}-${partition}`;
        } else {
          // Non-partitioned format: /api/mt/emea
          const simpleMatch = middleTierUrl.match(/\/api\/mt\/([a-z]+)$/);
          if (simpleMatch) {
            // No partition - calendar API uses non-partitioned URL
            regionPartition = simpleMatch[1];
          }
        }
      }

      return {
        region,
        partition: partition ?? '',
        regionPartition: regionPartition ?? region,
        hasPartition,
        middleTierUrl: middleTierUrl ?? '',
        chatServiceUrl,
        csaServiceUrl: csaServiceUrl ?? `${teamsBaseUrl}/api/csa/${region}`,
        teamsBaseUrl,
      };
    } catch {
      continue;
    }
  }

  return null;
}

/** User details from DISCOVER-USER-DETAILS. */
export interface UserDetails {
  /** User's MRI (e.g., "8:orgid:abc..."). */
  mri: string;
  /** User's region (e.g., "amer", "emea"). */
  region: string;
  /** User's partition (e.g., "amer01"). */
  userPartition: string;
  /** Tenant's partition (e.g., "amer02"). */
  tenantPartition: string;
  /** License details. */
  licenses: {
    isFreemium: boolean;
    isTrial: boolean;
    isTeamsEnabled: boolean;
    isCopilot: boolean;
    isTranscriptEnabled: boolean;
    isFrontline: boolean;
  };
}

/**
 * Extracts user details from DISCOVER-USER-DETAILS in localStorage.
 * 
 * This provides user-specific info including:
 * - User's MRI
 * - Region and partition info
 * - License details (Copilot, transcription, etc.)
 */
export function extractUserDetails(state?: SessionState): UserDetails | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  for (const item of localStorage) {
    if (!item.name.includes('DISCOVER-USER-DETAILS')) continue;

    try {
      const data = JSON.parse(item.value) as { 
        item?: {
          id?: string;
          region?: string;
          userPartition?: string;
          partition?: string;
          licenseDetails?: Record<string, unknown>;
        };
      };
      
      const details = data.item;
      if (!details?.id || !details?.region) continue;

      const licenses = details.licenseDetails ?? {};

      return {
        mri: details.id,
        region: details.region,
        userPartition: details.userPartition ?? '',
        tenantPartition: details.partition ?? '',
        licenses: {
          isFreemium: licenses.isFreemium === true,
          isTrial: licenses.isTrial === true,
          isTeamsEnabled: licenses.isTeamsEnabled === true,
          isCopilot: licenses.isCopilot === true,
          isTranscriptEnabled: licenses.isTranscriptEnabled === true,
          isFrontline: licenses.isFrontline === true,
        },
      };
    } catch {
      continue;
    }
  }

  return null;
}

/**
 * Extracts user MRI from the Substrate token's oid claim.
 */
function extractUserMriFromSubstrate(state?: SessionState): string | null {
  const substrateInfo = extractSubstrateToken(state);
  if (!substrateInfo) return null;

  const payload = decodeJwtPayload(substrateInfo.token);
  if (typeof payload?.oid === 'string') {
    return `${MRI_ORGID_PREFIX}${payload.oid}`;
  }
  return null;
}

/**
 * Extracts authentication info needed for messaging API.
 * Unlike other APIs, messaging uses cookies rather than localStorage tokens.
 */
export function extractMessageAuth(state?: SessionState): MessageAuthInfo | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  const cookies = sessionState.cookies ?? [];
  const teamsCookies = cookies.filter(c => c.domain?.includes('teams.microsoft.com'));

  // Extract the two required cookies
  const skypeToken = teamsCookies.find(c => c.name === 'skypetoken_asm')?.value ?? null;
  const rawAuthToken = teamsCookies.find(c => c.name === 'authtoken')?.value ?? null;
  
  if (!skypeToken || !rawAuthToken) return null;

  // Decode authtoken (URL-encoded, may have 'Bearer=' prefix)
  let authToken = decodeURIComponent(rawAuthToken);
  if (authToken.startsWith('Bearer=')) {
    authToken = authToken.substring(7);
  }

  // Extract userMri from skypeToken's skypeid claim, or fall back to authToken's oid
  const userMri = extractMriFromSkypeToken(skypeToken) 
    ?? extractMriFromAuthToken(authToken);

  if (!userMri) return null;

  return { skypeToken, authToken, userMri };
}

/**
 * Gets messaging token status for diagnostics.
 * The skypetoken_asm cookie is a JWT with an exp claim.
 */
export function getMessageAuthStatus(): {
  hasToken: boolean;
  expiresAt?: string;
  minutesRemaining?: number;
} {
  const sessionState = readSessionState();
  if (!sessionState) {
    return { hasToken: false };
  }

  const cookies = sessionState.cookies ?? [];
  const skypeToken = cookies.find(
    c => c.domain?.includes('teams.microsoft.com') && c.name === 'skypetoken_asm'
  )?.value;

  if (!skypeToken) {
    return { hasToken: false };
  }

  const expiry = getJwtExpiry(skypeToken);
  if (!expiry) {
    // Token exists but can't parse expiry - assume valid
    return { hasToken: true };
  }

  const now = Date.now();
  const expiryMs = expiry.getTime();

  return {
    hasToken: expiryMs > now,
    expiresAt: expiry.toISOString(),
    minutesRemaining: Math.max(0, Math.round((expiryMs - now) / 1000 / 60)),
  };
}

function extractMriFromSkypeToken(token: string): string | null {
  const payload = decodeJwtPayload(token);
  if (typeof payload?.skypeid !== 'string') return null;
  
  // The skypeid claim may be 'orgid:guid' without the '8:' prefix
  // Ensure we return the full MRI format '8:orgid:guid'
  const skypeid = payload.skypeid;
  if (skypeid.startsWith(MRI_TYPE_PREFIX)) {
    return skypeid;
  } else if (skypeid.startsWith(ORGID_PREFIX)) {
    return `${MRI_TYPE_PREFIX}${skypeid}`;
  }
  return skypeid;
}

function extractMriFromAuthToken(token: string): string | null {
  const payload = decodeJwtPayload(token);
  return typeof payload?.oid === 'string' ? `${MRI_ORGID_PREFIX}${payload.oid}` : null;
}

/**
 * Extracts the CSA token for the conversationFolders API.
 * This searches all origins, not just teams.microsoft.com.
 */
export function extractCsaToken(state?: SessionState): string | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;

  for (const origin of sessionState.origins ?? []) {
    for (const item of origin.localStorage ?? []) {
      // Skip temporary entries, look for chatsvcagg tokens
      if (item.name.startsWith('tmp.')) continue;
      if (!item.name.includes('chatsvcagg.teams.microsoft.com')) continue;

      try {
        const entry = JSON.parse(item.value) as { secret?: string };
        if (entry.secret) return entry.secret;
      } catch {
        // Ignore parse errors
      }
    }
  }

  return null;
}

// ============================================================================
// User Profile
// ============================================================================

/**
 * Gets the current user's profile from cached JWT tokens.
 */
export function getUserProfile(state?: SessionState): UserProfile | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      if (!isJwtToken(entry.secret)) continue;

      const payload = decodeJwtPayload(entry.secret);
      if (payload) {
        const profile = parseJwtProfile(payload);
        if (profile) return profile;
      }
    } catch {
      continue;
    }
  }

  return null;
}

/**
 * Gets user's display name from session state.
 * Searches localStorage entries first, then falls back to JWT claims.
 */
export function getUserDisplayName(state?: SessionState): string | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  // First pass: look for explicit displayName in localStorage
  for (const item of localStorage) {
    // Quick filter before parsing
    if (!item.value?.includes('displayName') && !item.value?.includes('givenName')) {
      continue;
    }

    try {
      const entry = JSON.parse(item.value);
      if (entry.displayName) return entry.displayName;
      if (entry.name?.displayName) return entry.name.displayName;
    } catch {
      continue;
    }
  }

  // Fallback: extract from Teams token's name claim
  const teamsToken = extractTeamsToken(state);
  if (teamsToken) {
    const payload = decodeJwtPayload(teamsToken.token);
    if (typeof payload?.name === 'string') return payload.name;
  }

  return null;
}

// ============================================================================
// Token Status Checks
// ============================================================================

/**
 * Checks if tokens in session state are expired.
 */
export function areTokensExpired(state?: SessionState): boolean {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return true;

  const substrate = extractSubstrateToken(sessionState);
  return !substrate || substrate.expiry.getTime() <= Date.now();
}

// Re-export clearTokenCache for convenience
export { clearTokenCache };

// ============================================================================
// Config Discovery (for debugging/research)
// ============================================================================

/**
 * Discovered configuration from localStorage.
 * Used for debugging and understanding what config is available.
 */
export interface DiscoveredConfig {
  /** All DISCOVER-* keys with their parsed values. */
  discoveryConfigs: Record<string, unknown>;
  /** All keys that look like configuration (not tokens). */
  configKeys: string[];
  /** Content of interesting config keys (settings, flags). */
  configContents: Record<string, unknown>;
  /** Teams base URL extracted from discovery config. */
  teamsBaseUrl: string | null;
  /** Substrate URL if found in any config. */
  substrateUrl: string | null;
  /** All unique hosts found in config URLs. */
  uniqueHosts: string[];
}

/**
 * Extracts all configuration data from localStorage for debugging.
 * This helps discover what config is available from different tenants.
 */
export function discoverConfig(state?: SessionState): DiscoveredConfig | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;

  const discoveryConfigs: Record<string, unknown> = {};
  const configKeys: string[] = [];
  const configContents: Record<string, unknown> = {};
  const allUrls: string[] = [];

  for (const item of localStorage) {
    // Collect all DISCOVER-* keys
    if (item.name.includes('DISCOVER')) {
      try {
        discoveryConfigs[item.name] = JSON.parse(item.value);
      } catch {
        discoveryConfigs[item.name] = item.value;
      }
    }

    // Collect keys that look like config (not tokens/cache)
    const isConfigKey = 
      item.name.includes('config') ||
      item.name.includes('CONFIG') ||
      item.name.includes('settings') ||
      item.name.includes('SETTINGS') ||
      item.name.includes('environment') ||
      item.name.includes('ENVIRONMENT') ||
      item.name.includes('endpoint') ||
      item.name.includes('ENDPOINT') ||
      item.name.includes('DISCOVER') ||
      item.name.includes('flags') ||
      item.name.includes('FLAGS');

    if (isConfigKey) {
      configKeys.push(item.name);
      // Also capture content for settings/flags keys (but not large token keys)
      if (!item.name.includes('accesstoken') && !item.name.includes('DISCOVER')) {
        try {
          configContents[item.name] = JSON.parse(item.value);
        } catch {
          // Only capture if it's short enough to be useful
          if (item.value.length < MAX_DEBUG_CONFIG_VALUE_LENGTH) {
            configContents[item.name] = item.value;
          }
        }
      }
    }

    // Extract URLs from any JSON values
    try {
      const extractUrls = (obj: unknown, depth = 0): void => {
        if (depth > 5) return; // Prevent infinite recursion
        if (typeof obj === 'string') {
          // Match URLs
          const urlMatch = obj.match(/https?:\/\/[^\s"'<>]+/g);
          if (urlMatch) allUrls.push(...urlMatch);
        } else if (typeof obj === 'object' && obj !== null) {
          for (const value of Object.values(obj)) {
            extractUrls(value, depth + 1);
          }
        }
      };
      extractUrls(JSON.parse(item.value));
    } catch {
      // Not JSON, skip
    }
  }

  // Extract unique hosts from URLs
  const uniqueHosts = [...new Set(
    allUrls
      .map(url => {
        try {
          return new URL(url).host;
        } catch {
          return null;
        }
      })
      .filter((h): h is string => h !== null)
  )].sort();

  // Try to find Teams base URL from discovery config
  let teamsBaseUrl: string | null = null;
  let substrateUrl: string | null = null;

  for (const [key, value] of Object.entries(discoveryConfigs)) {
    if (key.includes('DISCOVER-REGION-GTM') && typeof value === 'object' && value !== null) {
      const item = (value as Record<string, unknown>).item as Record<string, string> | undefined;
      if (item?.chatServiceAfd) {
        try {
          const url = new URL(item.chatServiceAfd);
          teamsBaseUrl = `${url.protocol}//${url.host}`;
        } catch {
          // Invalid URL
        }
      }
    }
  }

  // Look for Substrate URLs
  for (const host of uniqueHosts) {
    if (host.includes('substrate')) {
      substrateUrl = `https://${host}`;
      break;
    }
  }

  return {
    discoveryConfigs,
    configKeys,
    configContents,
    teamsBaseUrl,
    substrateUrl,
    uniqueHosts,
  };
}
