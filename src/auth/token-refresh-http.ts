/**
 * Browserless token refresh via direct HTTP calls.
 * 
 * Refreshes MSAL tokens by extracting the refresh token from session state
 * and POSTing to Azure AD's OAuth2 token endpoint. This eliminates the need
 * to spawn a headless browser for token refresh (~100ms vs ~8s).
 * 
 * The refresh token grant works identically for standard Microsoft login and
 * corporate SSO (ADFS/Okta federation) — once you have a refresh token, it's
 * a standard Azure AD token regardless of how the user originally authenticated.
 * 
 * First login still requires a browser — this module only handles refresh of
 * existing sessions where a refresh token is already available.
 * 
 * Flow:
 * 1. Extract refresh token, client ID, tenant ID from MSAL cache in session state
 * 2. POST to Azure AD token endpoint for each required scope
 * 3. Update session state localStorage with new MSAL cache entries
 * 4. Exchange Skype Spaces token for skypetoken_asm cookie
 * 5. Update session state cookies
 * 6. Write updated session state back to encrypted storage
 */

import {
  readSessionState,
  writeSessionState,
  getTeamsOrigin,
  type SessionState,
} from './session-store.js';
import { clearTokenCache } from './token-extractor.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';

// ============================================================================
// Types
// ============================================================================

/** MSAL cache entry for a refresh token. */
interface MsalRefreshToken {
  credentialType: 'RefreshToken';
  homeAccountId: string;
  environment: string;
  clientId: string;
  secret: string;
  /** Unix timestamp string (seconds). */
  expiresOn?: string;
  /** Timestamp string (ms since epoch). */
  lastUpdatedAt?: string;
}

/** MSAL cache entry for an access token. */
interface MsalAccessToken {
  credentialType: 'AccessToken';
  homeAccountId: string;
  environment: string;
  clientId: string;
  realm: string;
  target: string;
  tokenType: string;
  secret: string;
  /** Unix timestamp string (seconds). */
  expiresOn: string;
  /** Unix timestamp string (seconds). */
  extendedExpiresOn: string;
  /** Unix timestamp string (seconds). */
  cachedAt: string;
}

/** Extracted MSAL cache info needed for refresh. */
interface MsalCacheInfo {
  refreshToken: string;
  clientId: string;
  tenantId: string;
  homeAccountId: string;
  environment: string;
  /** The localStorage key for the refresh token entry. */
  refreshTokenKey: string;
}

/** Azure AD token response. */
interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  token_type: string;
  expires_in: number;
  scope: string;
  ext_expires_in?: number;
}

/** Authsvc response for skype token exchange. */
interface AuthsvcResponse {
  tokens?: {
    skypeToken?: string;
    expiresIn?: number;
  };
  regionGtms?: Record<string, unknown>;
}

/** Result of a successful HTTP token refresh. */
export interface HttpRefreshResult {
  /** Number of access tokens refreshed. */
  tokensRefreshed: number;
  /** Whether the skype token was refreshed. */
  skypeTokenRefreshed: boolean;
  /** Whether the refresh token itself was rotated. */
  refreshTokenRotated: boolean;
}

// ============================================================================
// Constants
// ============================================================================

/** Azure AD OAuth2 token endpoint template. */
const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token';

/** Teams authsvc endpoint for skype token exchange. */
const AUTHSVC_ENDPOINT = 'https://authsvc.teams.microsoft.com/v1.0/authz';

/**
 * Scopes to refresh. Each entry maps a resource identifier to the scopes
 * we request. The resource identifier is used to match existing MSAL cache
 * entries so we can update them in-place.
 */
const REFRESH_SCOPES = [
  {
    /** Substrate search/people APIs. */
    resource: 'substrate.office.com',
    scopes: 'https://substrate.office.com/.default offline_access',
  },
  {
    /** Calendar/Skype Spaces + skypetoken_asm derivation. */
    resource: 'api.spaces.skype.com',
    scopes: 'https://api.spaces.skype.com/.default offline_access',
  },
  {
    /** CSA token for messaging APIs. */
    resource: 'chatsvcagg.teams.microsoft.com',
    scopes: 'https://chatsvcagg.teams.microsoft.com/.default offline_access',
  },
  {
    /** Microsoft Graph API (spike: send message via Graph). */
    resource: 'graph.microsoft.com',
    scopes: 'https://graph.microsoft.com/.default offline_access',
  },
] as const;

/** HTTP request timeout for token refresh calls (ms). */
const REFRESH_TIMEOUT_MS = 10000;

// ============================================================================
// MSAL Cache Extraction
// ============================================================================

/**
 * Extracts MSAL cache info (refresh token, client ID, tenant ID) from session state.
 */
function extractMsalCacheInfo(state: SessionState): MsalCacheInfo | null {
  const teamsOrigin = getTeamsOrigin(state);
  if (!teamsOrigin?.localStorage) return null;

  let refreshToken: MsalRefreshToken | null = null;
  let refreshTokenKey: string | null = null;
  let tenantId: string | null = null;

  for (const item of teamsOrigin.localStorage) {
    try {
      const entry = JSON.parse(item.value);

      // Find the refresh token entry
      if (entry.credentialType === 'RefreshToken' && entry.secret && entry.clientId) {
        refreshToken = entry as MsalRefreshToken;
        refreshTokenKey = item.name;
      }

      // Extract tenant ID from any access token's realm field
      if (entry.credentialType === 'AccessToken' && entry.realm && !tenantId) {
        tenantId = entry.realm;
      }
    } catch {
      continue;
    }
  }

  if (!refreshToken || !refreshTokenKey || !tenantId) return null;

  return {
    refreshToken: refreshToken.secret,
    clientId: refreshToken.clientId,
    tenantId,
    homeAccountId: refreshToken.homeAccountId,
    environment: refreshToken.environment,
    refreshTokenKey,
  };
}

// ============================================================================
// OAuth2 Token Refresh
// ============================================================================

/**
 * Refreshes an access token via Azure AD's OAuth2 token endpoint.
 * 
 * Uses the refresh_token grant type with the Teams SPA public client ID.
 * No client secret is needed — Teams is a public client (SPA).
 */
async function refreshAccessToken(
  tenantId: string,
  clientId: string,
  refreshToken: string,
  scopes: string,
): Promise<Result<TokenResponse>> {
  const url = TOKEN_ENDPOINT.replace('{tenantId}', tenantId);

  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: clientId,
    refresh_token: refreshToken,
    scope: scopes,
  });

  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), REFRESH_TIMEOUT_MS);

    // The Origin header is required because the Teams client ID is registered as
    // a Single-Page Application (SPA). Azure AD validates that refresh token grants
    // from SPA clients include a cross-origin Origin header matching a registered
    // redirect URI. Without this, Azure AD returns AADSTS9002327.
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': 'https://teams.microsoft.com',
      },
      body: body.toString(),
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      let errorDetail = `HTTP ${response.status}`;

      // Parse Azure AD error response for better diagnostics
      try {
        const errorJson = JSON.parse(errorText);
        if (errorJson.error_description) {
          errorDetail = errorJson.error_description;
        } else if (errorJson.error) {
          errorDetail = `${errorJson.error}: ${errorJson.error_description || errorText}`;
        }
      } catch {
        errorDetail = `HTTP ${response.status}: ${errorText.substring(0, 200)}`;
      }

      // Specific error codes that indicate the refresh token is invalid/expired
      const isAuthError = response.status === 400 || response.status === 401;

      return err(createError(
        isAuthError ? ErrorCode.AUTH_EXPIRED : ErrorCode.UNKNOWN,
        `Token refresh failed: ${errorDetail}`,
        { retryable: !isAuthError }
      ));
    }

    const data = await response.json() as TokenResponse;
    return ok(data);

  } catch (error) {
    if (error instanceof Error && error.name === 'AbortError') {
      return err(createError(
        ErrorCode.TIMEOUT,
        'Token refresh request timed out',
        { retryable: true }
      ));
    }

    return err(createError(
      ErrorCode.NETWORK_ERROR,
      `Token refresh network error: ${error instanceof Error ? error.message : String(error)}`,
      { retryable: true }
    ));
  }
}

// ============================================================================
// Skype Token Exchange
// ============================================================================

/**
 * Exchanges a Skype Spaces access token for a skypetoken_asm.
 * 
 * POST to authsvc.teams.microsoft.com with the Skype Spaces bearer token.
 * Returns the skype token which is used as a cookie for messaging APIs.
 */
async function exchangeSkypeToken(
  skypeSpacesToken: string,
): Promise<Result<{ skypeToken: string; expiresIn: number }>> {
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), REFRESH_TIMEOUT_MS);

    const response = await fetch(AUTHSVC_ENDPOINT, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${skypeSpacesToken}`,
        'Content-Type': 'application/json',
      },
      body: '{}',
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      return err(createError(
        ErrorCode.AUTH_EXPIRED,
        `Skype token exchange failed: HTTP ${response.status}: ${errorText.substring(0, 200)}`,
        { retryable: false }
      ));
    }

    const data = await response.json() as AuthsvcResponse;
    const skypeToken = data.tokens?.skypeToken;
    const expiresIn = data.tokens?.expiresIn;

    if (!skypeToken) {
      return err(createError(
        ErrorCode.UNKNOWN,
        'Skype token exchange returned no token',
        { retryable: false }
      ));
    }

    return ok({ skypeToken, expiresIn: expiresIn ?? 86400 });

  } catch (error) {
    if (error instanceof Error && error.name === 'AbortError') {
      return err(createError(
        ErrorCode.TIMEOUT,
        'Skype token exchange timed out',
        { retryable: true }
      ));
    }

    return err(createError(
      ErrorCode.NETWORK_ERROR,
      `Skype token exchange error: ${error instanceof Error ? error.message : String(error)}`,
      { retryable: true }
    ));
  }
}

// ============================================================================
// Session State Update
// ============================================================================

/**
 * Finds the localStorage key for an existing MSAL access token entry
 * that matches the given resource.
 */
function findAccessTokenKey(
  localStorage: Array<{ name: string; value: string }>,
  resource: string,
): { key: string; entry: MsalAccessToken } | null {
  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      if (entry.credentialType !== 'AccessToken') continue;
      if (!entry.target?.includes(resource)) continue;
      return { key: item.name, entry: entry as MsalAccessToken };
    } catch {
      continue;
    }
  }
  return null;
}

/**
 * Updates a localStorage entry in the session state.
 * If the key exists, updates it. Otherwise, adds a new entry.
 */
function updateLocalStorageEntry(
  localStorage: Array<{ name: string; value: string }>,
  key: string,
  value: string,
): void {
  const existing = localStorage.findIndex(item => item.name === key);
  if (existing >= 0) {
    localStorage[existing].value = value;
  } else {
    localStorage.push({ name: key, value });
  }
}

/**
 * Updates the MSAL access token cache entry with a new token.
 * Maintains the exact MSAL cache format so token-extractor.ts can find it.
 */
function updateAccessTokenInCache(
  localStorage: Array<{ name: string; value: string }>,
  resource: string,
  tokenResponse: TokenResponse,
  cacheInfo: MsalCacheInfo,
): boolean {
  const existing = findAccessTokenKey(localStorage, resource);
  if (!existing) {
    // No existing entry for this resource — create one
    // Build the key in MSAL format: {homeAccountId}-{environment}-accesstoken-{clientId}-{realm}-{target}
    const now = Math.floor(Date.now() / 1000);
    const newEntry: MsalAccessToken = {
      credentialType: 'AccessToken',
      homeAccountId: cacheInfo.homeAccountId,
      environment: cacheInfo.environment,
      clientId: cacheInfo.clientId,
      realm: cacheInfo.tenantId,
      target: tokenResponse.scope,
      tokenType: tokenResponse.token_type || 'Bearer',
      secret: tokenResponse.access_token,
      expiresOn: String(now + tokenResponse.expires_in),
      extendedExpiresOn: String(now + (tokenResponse.ext_expires_in ?? tokenResponse.expires_in)),
      cachedAt: String(now),
    };

    // Build key in MSAL format
    const scopeKey = tokenResponse.scope.replace(/ /g, ' ').toLowerCase();
    const key = `${cacheInfo.homeAccountId}-${cacheInfo.environment}-accesstoken-${cacheInfo.clientId}-${cacheInfo.tenantId}-${scopeKey}`;
    updateLocalStorageEntry(localStorage, key, JSON.stringify(newEntry));
    return true;
  }

  // Update existing entry
  const now = Math.floor(Date.now() / 1000);
  const updated: MsalAccessToken = {
    ...existing.entry,
    secret: tokenResponse.access_token,
    expiresOn: String(now + tokenResponse.expires_in),
    extendedExpiresOn: String(now + (tokenResponse.ext_expires_in ?? tokenResponse.expires_in)),
    cachedAt: String(now),
  };

  // If the response includes new scopes, update target
  if (tokenResponse.scope) {
    updated.target = tokenResponse.scope;
  }

  updateLocalStorageEntry(localStorage, existing.key, JSON.stringify(updated));
  return true;
}

/**
 * Updates the refresh token in the MSAL cache.
 * Azure AD may rotate the refresh token on each use.
 */
function updateRefreshTokenInCache(
  localStorage: Array<{ name: string; value: string }>,
  refreshTokenKey: string,
  newRefreshToken: string,
  existingEntry: string,
): void {
  try {
    const entry = JSON.parse(existingEntry) as MsalRefreshToken;
    entry.secret = newRefreshToken;
    entry.lastUpdatedAt = String(Date.now());
    updateLocalStorageEntry(localStorage, refreshTokenKey, JSON.stringify(entry));
  } catch {
    // If we can't parse the existing entry, just update the secret
    updateLocalStorageEntry(localStorage, refreshTokenKey, JSON.stringify({
      credentialType: 'RefreshToken',
      secret: newRefreshToken,
      lastUpdatedAt: String(Date.now()),
    }));
  }
}

/**
 * Updates the skypetoken_asm cookies in session state.
 */
function updateSkypeTokenCookies(
  state: SessionState,
  skypeToken: string,
  expiresIn: number,
): void {
  const expiryTimestamp = Date.now() / 1000 + expiresIn;

  // Domains where skypetoken_asm is set
  const skypeTokenDomains = ['.asyncgw.teams.microsoft.com', '.asm.skype.com'];

  for (const domain of skypeTokenDomains) {
    const existingIdx = state.cookies.findIndex(
      c => c.name === 'skypetoken_asm' && c.domain === domain
    );

    const cookie = {
      name: 'skypetoken_asm',
      value: skypeToken,
      domain,
      path: '/',
      expires: expiryTimestamp,
      httpOnly: true,
      secure: true,
      sameSite: 'None' as const,
    };

    if (existingIdx >= 0) {
      state.cookies[existingIdx] = cookie;
    } else {
      state.cookies.push(cookie);
    }
  }
}

/**
 * Updates the authtoken cookie in session state.
 * The authtoken is the Skype Spaces access token stored as a cookie.
 */
function updateAuthTokenCookie(
  state: SessionState,
  skypeSpacesToken: string,
  expiresIn: number,
): void {
  const expiryTimestamp = Date.now() / 1000 + expiresIn;

  const existingIdx = state.cookies.findIndex(
    c => c.name === 'authtoken' && c.domain === 'teams.microsoft.com'
  );

  const cookie = {
    name: 'authtoken',
    value: `Bearer%3D${encodeURIComponent(skypeSpacesToken)}`,
    domain: 'teams.microsoft.com',
    path: '/',
    expires: expiryTimestamp,
    httpOnly: false,
    secure: true,
    sameSite: 'None' as const,
  };

  if (existingIdx >= 0) {
    state.cookies[existingIdx] = cookie;
  } else {
    state.cookies.push(cookie);
  }
}

// ============================================================================
// Main Refresh Function
// ============================================================================

/**
 * Refreshes tokens via direct HTTP calls (no browser needed).
 * 
 * This is the primary token refresh mechanism. It:
 * 1. Extracts the MSAL refresh token from session state
 * 2. Calls Azure AD's token endpoint for each required scope
 * 3. Updates the MSAL cache in session state with new tokens
 * 4. Exchanges the Skype Spaces token for skypetoken_asm
 * 5. Updates cookies in session state
 * 6. Writes the updated session state back to encrypted storage
 * 
 * Falls back to browser-based refresh if this fails (e.g., refresh token
 * expired, Conditional Access policy requires interactive auth).
 */
export async function refreshTokensViaHttp(): Promise<Result<HttpRefreshResult>> {
  // Read current session state
  const state = readSessionState();
  if (!state) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No session state found. Browser login is required for first authentication.',
      { suggestions: ['Call teams_login to authenticate via browser'] }
    ));
  }

  // Extract MSAL cache info
  const cacheInfo = extractMsalCacheInfo(state);
  if (!cacheInfo) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No MSAL refresh token found in session state. Browser login is required.',
      { suggestions: ['Call teams_login to authenticate via browser'] }
    ));
  }

  const teamsOrigin = getTeamsOrigin(state);
  if (!teamsOrigin?.localStorage) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No Teams localStorage found in session state.',
    ));
  }

  // Get the existing refresh token entry for later update
  const refreshTokenEntry = teamsOrigin.localStorage.find(
    item => item.name === cacheInfo.refreshTokenKey
  );

  let tokensRefreshed = 0;
  let refreshTokenRotated = false;
  let skypeSpacesToken: string | null = null;
  let skypeSpacesExpiresIn: number | null = null;

  // Use the current refresh token; it may be rotated by Azure AD
  let currentRefreshToken = cacheInfo.refreshToken;

  // Refresh each scope
  for (const scope of REFRESH_SCOPES) {
    const result = await refreshAccessToken(
      cacheInfo.tenantId,
      cacheInfo.clientId,
      currentRefreshToken,
      scope.scopes,
    );

    if (!result.ok) {
      // If any refresh fails with auth error, the refresh token is likely expired
      if (result.error.code === ErrorCode.AUTH_EXPIRED) {
        return err(createError(
          ErrorCode.AUTH_EXPIRED,
          `HTTP token refresh failed for ${scope.resource}: ${result.error.message}. Browser login required.`,
          { suggestions: ['Call teams_login to re-authenticate via browser'] }
        ));
      }
      // For other errors (network, timeout), log and continue with remaining scopes
      console.error(`[token-refresh-http] Failed to refresh ${scope.resource}: ${result.error.message}`);
      continue;
    }

    const tokenResponse = result.value;

    // Update the access token in MSAL cache
    updateAccessTokenInCache(
      teamsOrigin.localStorage,
      scope.resource,
      tokenResponse,
      cacheInfo,
    );
    tokensRefreshed++;

    // If Azure AD rotated the refresh token, use the new one for subsequent calls
    if (tokenResponse.refresh_token && tokenResponse.refresh_token !== currentRefreshToken) {
      currentRefreshToken = tokenResponse.refresh_token;
      refreshTokenRotated = true;
    }

    // Capture the Skype Spaces token for skypetoken_asm exchange
    if (scope.resource === 'api.spaces.skype.com') {
      skypeSpacesToken = tokenResponse.access_token;
      skypeSpacesExpiresIn = tokenResponse.expires_in;
    }
  }

  if (tokensRefreshed === 0) {
    return err(createError(
      ErrorCode.UNKNOWN,
      'HTTP token refresh failed: no tokens were successfully refreshed.',
      { retryable: true }
    ));
  }

  // Update the refresh token if it was rotated
  if (refreshTokenRotated && refreshTokenEntry) {
    updateRefreshTokenInCache(
      teamsOrigin.localStorage,
      cacheInfo.refreshTokenKey,
      currentRefreshToken,
      refreshTokenEntry.value,
    );
  }

  // Exchange Skype Spaces token for skypetoken_asm
  let skypeTokenRefreshed = false;
  if (skypeSpacesToken) {
    const skypeResult = await exchangeSkypeToken(skypeSpacesToken);
    if (skypeResult.ok) {
      updateSkypeTokenCookies(state, skypeResult.value.skypeToken, skypeResult.value.expiresIn);
      updateAuthTokenCookie(state, skypeSpacesToken, skypeSpacesExpiresIn ?? 3600);
      skypeTokenRefreshed = true;
    } else {
      console.error(`[token-refresh-http] Skype token exchange failed: ${skypeResult.error.message}`);
    }
  }

  // Write updated session state back to encrypted storage
  writeSessionState(state);

  // Clear the in-memory token cache so it re-reads from the updated session
  clearTokenCache();

  return ok({
    tokensRefreshed,
    skypeTokenRefreshed,
    refreshTokenRotated,
  });
}
