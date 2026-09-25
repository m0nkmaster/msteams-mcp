/**
 * Token extraction from session state.
 * 
 * Extracts various authentication tokens from Playwright's saved session state.
 * Teams stores MSAL tokens in localStorage; we parse these to get bearer tokens
 * for various APIs (Substrate search, chatsvc messaging, etc.).
 */

import {
  readSessionState,
  writeSessionState,
  readTokenCache,
  writeTokenCache,
  getTeamsOrigin,
  type SessionState,
} from './session-store.js';
import { parseJwtProfile, type UserProfile } from '../utils/parsers.js';
import { MRI_TYPE_PREFIX, ORGID_PREFIX, MRI_ORGID_PREFIX, ASSIGNMENTS_APP_ID, GRAPH_AUDIENCES } from '../constants.js';

// ============================================================================
// JWT Utilities
// ============================================================================

/**
 * Decodes a JWT token's payload without verifying the signature.
 */
export function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;
    return JSON.parse(Buffer.from(parts[1], 'base64url').toString());
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

/** localStorage entry from Teams origin. */
type LocalStorageEntry = { name: string; value: string };

/** Teams origin localStorage entries, or [] when there is no session. */
function teamsLocalStorage(state?: SessionState): LocalStorageEntry[] {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return [];
  return getTeamsOrigin(sessionState)?.localStorage ?? [];
}

// ============================================================================
// Types
// ============================================================================

/** A JWT access token and its expiry. */
export interface SubstrateTokenInfo {
  token: string;
  expiry: Date;
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
 * Longest-lived unexpired JWT among MSAL cache entries matching `match`.
 * `payload` is the decoded JWT payload of the entry's secret.
 */
function bestJwt(
  state: SessionState | undefined,
  match: (entry: { credentialType?: string; target?: string }, payload: Record<string, unknown>) => boolean,
): SubstrateTokenInfo | null {
  let best: SubstrateTokenInfo | null = null;
  for (const item of teamsLocalStorage(state)) {
    try {
      const entry = JSON.parse(item.value);
      if (!isJwtToken(entry.secret)) continue;
      const payload = decodeJwtPayload(entry.secret);
      if (!payload || typeof payload.exp !== 'number' || !match(entry, payload)) continue;
      const expiry = new Date(payload.exp * 1000);
      if (expiry.getTime() <= Date.now()) continue;
      if (!best || expiry > best.expiry) best = { token: entry.secret, expiry };
    } catch {
      continue;
    }
  }
  return best;
}

/**
 * Extracts the Substrate search token from session state.
 * This token is used for search and people APIs.
 *
 * The target scope containing 'SubstrateSearch' matches all known formats:
 *   - substrate.office.com/search/SubstrateSearch (old)
 *   - substrate.office.com/SubstrateSearch-Internal.ReadWrite (new)
 *   - outlook.office.com/search/SubstrateSearch-Internal.ReadWrite (some Enterprise tenants)
 */
export function extractSubstrateToken(state?: SessionState): SubstrateTokenInfo | null {
  return bestJwt(state, entry => entry.target?.includes('SubstrateSearch') ?? false);
}

// ============================================================================
// Cached Token Access
// ============================================================================

/**
 * Gets a valid Substrate token, either from cache or by extracting from session.
 */
export function getValidSubstrateToken(): string | null {
  const cache = readTokenCache();
  if (cache && cache.substrateTokenExpiry > Date.now()) {
    return cache.substrateToken;
  }

  const extracted = extractSubstrateToken();
  if (!extracted) return null;

  writeTokenCache({
    substrateToken: extracted.token,
    substrateTokenExpiry: extracted.expiry.getTime(),
    extractedAt: Date.now(),
  });
  return extracted.token;
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
  if (!extracted) return { hasToken: false };
  return {
    hasToken: true,
    expiresAt: extracted.expiry.toISOString(),
    minutesRemaining: Math.round((extracted.expiry.getTime() - Date.now()) / 60000),
  };
}

// ============================================================================
// Optional Tokens (EDU Assignments, Microsoft Graph)
// ============================================================================

/**
 * Extracts the EDU Assignments API token from session state.
 *
 * This token authenticates calls to `assignments.edu.cloud.microsoft`. It is
 * minted on demand by HTTP token refresh and stored in the MSAL cache.
 * Select AccessToken entries by the JWT audience, since Microsoft Graph uses
 * the same EduAssignments permission names for a different resource.
 */
export function extractAssignmentsToken(state?: SessionState): SubstrateTokenInfo | null {
  return extractTokenByAudience([ASSIGNMENTS_APP_ID], state);
}

/**
 * Extracts a Microsoft Graph token from session state. Teams web caches one for
 * its own client; the on-demand refresh (`refreshTokensViaHttp('graph')`) renews it.
 */
export function extractGraphToken(state?: SessionState): SubstrateTokenInfo | null {
  return extractTokenByAudience(GRAPH_AUDIENCES, state);
}

/** Longest-lived unexpired AccessToken whose JWT audience is one of `audiences`. */
function extractTokenByAudience(audiences: readonly string[], state?: SessionState): SubstrateTokenInfo | null {
  return bestJwt(state, (entry, payload) =>
    entry.credentialType === 'AccessToken' && audiences.includes(String(payload.aud)));
}

/** Remove only the rejected credential; leave other API tokens untouched. */
export function invalidateAccessToken(token: string): void {
  const state = readSessionState();
  if (!state) return;
  const origin = getTeamsOrigin(state);
  if (!origin) return;
  origin.localStorage = origin.localStorage.filter(item => {
    try {
      return JSON.parse(item.value).secret !== token;
    } catch {
      return true;
    }
  });
  writeSessionState(state);
}

/**
 * Gets a valid EDU Assignments token from the session, or null. Refreshed on
 * demand by `refreshTokensViaHttp('assignments')`.
 */
export function getValidAssignmentsToken(): string | null {
  return extractAssignmentsToken()?.token ?? null;
}

/** Gets a valid Microsoft Graph token from the session, or null. */
export function getValidGraphToken(): string | null {
  return extractGraphToken()?.token ?? null;
}

/**
 * Extracts the Skype Spaces API token from session state.
 *
 * This token is required for the calendar/meetings API (mt/part endpoints).
 * It has scope: https://api.spaces.skype.com/Authorization.ReadWrite
 */
export function extractSkypeSpacesToken(state?: SessionState): string | null {
  return bestJwt(state, entry => entry.target?.includes('api.spaces.skype.com') ?? false)?.token ?? null;
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
  for (const item of teamsLocalStorage(state)) {
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
  for (const item of teamsLocalStorage(state)) {
    try {
      const entry = JSON.parse(item.value);
      if (!isJwtToken(entry.secret)) continue;
      const payload = decodeJwtPayload(entry.secret);
      const profile = payload && parseJwtProfile(payload);
      if (profile) return profile;
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
  for (const item of teamsLocalStorage(state)) {
    // Quick filter before parsing
    if (!item.value?.includes('displayName') && !item.value?.includes('givenName')) continue;
    try {
      const entry = JSON.parse(item.value);
      if (entry.displayName) return entry.displayName;
      if (entry.name?.displayName) return entry.name.displayName;
    } catch {
      continue;
    }
  }
  return getUserProfile(state)?.displayName ?? null;
}
