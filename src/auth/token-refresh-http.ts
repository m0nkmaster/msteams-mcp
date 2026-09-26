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

import { ASSIGNMENTS_APP_ID, GRAPH_AUDIENCES } from '../constants.js';
import {
  readSessionState,
  writeSessionState,
  clearTokenCache,
  getTeamsOrigin,
  type SessionState,
} from './session-store.js';
import { decodeJwtPayload } from './token-extractor.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import * as log from '../utils/logger.js';

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

// ============================================================================
// Constants
// ============================================================================

/** Azure AD OAuth2 token endpoint template. */
const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token';

/** Teams authsvc endpoint for skype token exchange. */
const AUTHSVC_ENDPOINT = 'https://authsvc.teams.microsoft.com/v1.0/authz';

/** Resource key used by the on-demand Assignments refresh. */
const ASSIGNMENTS_RESOURCE = 'EduAssignments';

/** Resource key used by the on-demand Microsoft Graph refresh. */
const GRAPH_RESOURCE = 'MicrosoftGraph';

/** Optional resources requested only when a tool needs them. */
export type OnDemandResource = 'assignments' | 'graph';

/**
 * Optional resources are cached by JWT audience rather than MSAL target, because
 * Graph and Assignments grant overlapping bare permission names.
 */
const ON_DEMAND_AUDIENCES: Record<string, readonly string[]> = {
  [ASSIGNMENTS_RESOURCE]: [ASSIGNMENTS_APP_ID],
  [GRAPH_RESOURCE]: GRAPH_AUDIENCES,
};

/** Core scopes are refreshed together; optional resources only on demand. */
const REFRESH_SCOPES: ReadonlyArray<{ resource: string; scopes: string; onDemand?: OnDemandResource; label?: string }> = [
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
    /**
     * EDU Assignments API (assignments.edu.cloud.microsoft, backed by OneNote EDU).
     * Use the app GUID to request a token. Cache selection checks the audience,
     * since Graph grants similarly named EduAssignments permissions.
     */
    resource: ASSIGNMENTS_RESOURCE,
    scopes: `${ASSIGNMENTS_APP_ID}/.default offline_access`,
    // EDU access may be unavailable even when the core Teams session is valid.
    onDemand: 'assignments',
    label: 'Assignments',
  },
  {
    /** Microsoft Graph, used to download files such as assignment attachments. */
    resource: GRAPH_RESOURCE,
    scopes: 'https://graph.microsoft.com/.default offline_access',
    onDemand: 'graph',
    label: 'Microsoft Graph',
  },
];

/** HTTP request timeout for token refresh calls (ms). */
const REFRESH_TIMEOUT_MS = 10000;

// ============================================================================
// MSAL Cache Extraction
// ============================================================================

/**
 * Extracts MSAL cache info (refresh token, client ID, tenant ID) from session state.
 * On failure returns a diagnostic string describing what was missing.
 */
function extractMsalCacheInfo(state: SessionState): { cacheInfo: MsalCacheInfo; refreshTokenEntry: MsalRefreshToken } | { missing: string } {
  const localStorage = getTeamsOrigin(state)?.localStorage ?? [];
  let refreshToken: MsalRefreshToken | null = null;
  let refreshTokenKey: string | null = null;
  let tenantId: string | null = null;

  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      if (entry.credentialType === 'RefreshToken' && entry.secret && entry.clientId) {
        refreshToken = entry as MsalRefreshToken;
        refreshTokenKey = item.name;
      }
      // Tenant ID comes from any access token's realm field
      if (entry.credentialType === 'AccessToken' && entry.realm && !tenantId) {
        tenantId = entry.realm;
      }
    } catch {
      continue;
    }
  }

  if (!refreshToken || !refreshTokenKey || !tenantId) {
    return { missing: `localStorage items: ${localStorage.length}, refreshToken: ${!!refreshToken}, tenantId: ${!!tenantId}` };
  }

  return {
    cacheInfo: {
      refreshToken: refreshToken.secret,
      clientId: refreshToken.clientId,
      tenantId,
      homeAccountId: refreshToken.homeAccountId,
      environment: refreshToken.environment,
      refreshTokenKey,
    },
    refreshTokenEntry: refreshToken,
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
      signal: AbortSignal.timeout(REFRESH_TIMEOUT_MS),
    });

    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      let errorDetail = `HTTP ${response.status}`;

      // Parse Azure AD error response for better diagnostics
      try {
        const errorJson = JSON.parse(errorText);
        if (errorJson.error_description) {
          errorDetail = errorJson.error_description;
        } else if (errorJson.error) {
          errorDetail = `${errorJson.error}: ${errorText}`;
        }
      } catch {
        errorDetail = `HTTP ${response.status}: ${errorText.substring(0, 200)}`;
      }

      // Consent/resource refusals are not an expired Teams login. Do not classify
      // all HTTP 400s this way: invalid_grant may require renewed SSO or MFA.
      // 50105 unassigned user, 53003 Conditional Access block, 90094 admin
      // consent, 650057 invalid resource: none are fixed by re-authenticating.
      const optional = REFRESH_SCOPES.find(scope => scope.onDemand && scope.scopes === scopes);
      if (optional && /AADSTS(?:50105|53003|65001|65004|90094|500011|650057|700016)\b/.test(errorDetail)) {
        return err(createError(ErrorCode.ACCESS_DENIED,
          `${optional.label} access was refused: ${errorDetail}`, {
            retryable: false,
            suggestions: [
              `${optional.label} is optional; all other Teams tools are unaffected`,
              `Check ${optional.label} availability and required consent with your tenant administrator`,
            ],
          }));
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
    if (error instanceof Error && error.name === 'TimeoutError') {
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
    const response = await fetch(AUTHSVC_ENDPOINT, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${skypeSpacesToken}`,
        'Content-Type': 'application/json',
      },
      body: '{}',
      signal: AbortSignal.timeout(REFRESH_TIMEOUT_MS),
    });

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
    if (error instanceof Error && error.name === 'TimeoutError') {
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

type LocalStorage = Array<{ name: string; value: string }>;

/**
 * Finds the existing MSAL access token entry for the given resource.
 */
function findAccessToken(
  localStorage: LocalStorage,
  resource: string,
): { key: string; entry: MsalAccessToken } | null {
  const audiences = ON_DEMAND_AUDIENCES[resource];
  for (const item of localStorage) {
    try {
      const entry = JSON.parse(item.value);
      if (entry.credentialType !== 'AccessToken') continue;
      if (audiences) {
        if (!audiences.includes(String(decodeJwtPayload(entry.secret)?.aud))) continue;
      } else if (!entry.target?.includes(resource)) continue;
      return { key: item.name, entry: entry as MsalAccessToken };
    } catch {
      continue;
    }
  }
  return null;
}

/** Sets a localStorage entry, adding it if the key is new. */
function upsertLocalStorage(localStorage: LocalStorage, key: string, value: string): void {
  const existing = localStorage.find(item => item.name === key);
  if (existing) existing.value = value;
  else localStorage.push({ name: key, value });
}

/** Sets a cookie, replacing any with the same name and domain. */
function upsertCookie(state: SessionState, cookie: SessionState['cookies'][number]): void {
  const idx = state.cookies.findIndex(c => c.name === cookie.name && c.domain === cookie.domain);
  if (idx >= 0) state.cookies[idx] = cookie;
  else state.cookies.push(cookie);
}

/**
 * Writes a new access token into the MSAL cache, updating the resource's
 * existing entry or creating one. Keeps the exact MSAL cache format so
 * token-extractor.ts can find it.
 */
function updateAccessTokenInCache(
  localStorage: LocalStorage,
  resource: string,
  tokenResponse: TokenResponse,
  cacheInfo: MsalCacheInfo,
): void {
  // Qualify Assignments scopes: Graph can return the same bare permission names
  // for the same account/client, but a different audience.
  const target = resource === ASSIGNMENTS_RESOURCE
    ? tokenResponse.scope.split(/\s+/).filter(Boolean).map(scope => scope.includes('/') ? scope : `${ASSIGNMENTS_APP_ID}/${scope}`).join(' ')
    : tokenResponse.scope;
  const existing = findAccessToken(localStorage, resource);
  const now = Math.floor(Date.now() / 1000);
  const entry: MsalAccessToken = {
    ...(existing?.entry ?? {
      credentialType: 'AccessToken',
      homeAccountId: cacheInfo.homeAccountId,
      environment: cacheInfo.environment,
      clientId: cacheInfo.clientId,
      realm: cacheInfo.tenantId,
      tokenType: tokenResponse.token_type || 'Bearer',
    }),
    target: target || existing?.entry.target || '',
    secret: tokenResponse.access_token,
    expiresOn: String(now + tokenResponse.expires_in),
    extendedExpiresOn: String(now + (tokenResponse.ext_expires_in ?? tokenResponse.expires_in)),
    cachedAt: String(now),
  };
  // MSAL key format: {homeAccountId}-{environment}-accesstoken-{clientId}-{realm}-{target}
  const key = existing?.key
    ?? `${cacheInfo.homeAccountId}-${cacheInfo.environment}-accesstoken-${cacheInfo.clientId}-${cacheInfo.tenantId}-${target.toLowerCase()}`;
  upsertLocalStorage(localStorage, key, JSON.stringify(entry));
}

/**
 * Stores a freshly exchanged skypetoken_asm cookie and the Skype Spaces
 * `authtoken` cookie (the Skype Spaces access token, used by messaging APIs).
 */
function updateSkypeCookies(
  state: SessionState,
  skypeToken: string,
  skypeExpiresIn: number,
  skypeSpacesToken: string,
  spacesExpiresIn: number,
): void {
  const nowSec = Date.now() / 1000;
  for (const domain of ['.asyncgw.teams.microsoft.com', '.asm.skype.com']) {
    upsertCookie(state, {
      name: 'skypetoken_asm', value: skypeToken, domain, path: '/',
      expires: nowSec + skypeExpiresIn, httpOnly: true, secure: true, sameSite: 'None',
    });
  }
  upsertCookie(state, {
    name: 'authtoken', value: `Bearer%3D${encodeURIComponent(skypeSpacesToken)}`,
    domain: 'teams.microsoft.com', path: '/',
    expires: nowSec + spacesExpiresIn, httpOnly: false, secure: true, sameSite: 'None',
  });
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
export async function refreshTokensViaHttp(resource: 'core' | OnDemandResource = 'core'): Promise<Result<void>> {
  const state = readSessionState();
  if (!state) {
    log.warn('token-refresh-http', 'No session state file found');
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'No session state found. Browser login is required for first authentication.',
      { suggestions: ['Call teams_login to authenticate via browser'] }
    ));
  }

  const extracted = extractMsalCacheInfo(state);
  if ('missing' in extracted) {
    log.warn('token-refresh-http', `MSAL cache extraction failed: ${extracted.missing}`);
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      `No MSAL refresh token found in session state (${extracted.missing}). Browser login is required.`,
      { suggestions: ['Call teams_login to authenticate via browser'] }
    ));
  }

  const { cacheInfo, refreshTokenEntry } = extracted;
  // Non-null: extraction found the refresh token in this origin's localStorage.
  const localStorage = getTeamsOrigin(state)!.localStorage;

  let tokensRefreshed = 0;
  let skypeSpaces: TokenResponse | null = null;
  const scopeErrors: string[] = [];

  // Azure AD may rotate the refresh token on each use
  let currentRefreshToken = cacheInfo.refreshToken;

  const scopes = REFRESH_SCOPES.filter(scope => resource === 'core' ? !scope.onDemand : scope.onDemand === resource);
  for (const scope of scopes) {
    const result = await refreshAccessToken(
      cacheInfo.tenantId,
      cacheInfo.clientId,
      currentRefreshToken,
      scope.scopes,
    );

    if (!result.ok) {
      // A resource-specific caller needs the actual failure, not core success.
      if (resource !== 'core') return result;
      if (result.error.code === ErrorCode.AUTH_EXPIRED) return result;
      // A transient core resource failure must not discard other refreshed tokens.
      log.warn('token-refresh-http', `Failed to refresh ${scope.resource}: ${result.error.message}`);
      scopeErrors.push(`${scope.resource}: ${result.error.message}`);
      continue;
    }

    updateAccessTokenInCache(localStorage, scope.resource, result.value, cacheInfo);
    tokensRefreshed++;
    if (result.value.refresh_token) currentRefreshToken = result.value.refresh_token;
    if (scope.resource === 'api.spaces.skype.com') skypeSpaces = result.value;
  }

  if (tokensRefreshed === 0) {
    return err(createError(
      ErrorCode.UNKNOWN,
      `HTTP token refresh failed: ${scopeErrors.length} of ${scopes.length} scopes failed. ${scopeErrors.join('; ')}`,
      { retryable: true }
    ));
  }

  if (currentRefreshToken !== cacheInfo.refreshToken) {
    upsertLocalStorage(localStorage, cacheInfo.refreshTokenKey, JSON.stringify({
      ...refreshTokenEntry,
      secret: currentRefreshToken,
      lastUpdatedAt: String(Date.now()),
    }));
  }

  // Exchange Skype Spaces token for skypetoken_asm
  let skypeTokenRefreshed = false;
  if (skypeSpaces) {
    const skypeResult = await exchangeSkypeToken(skypeSpaces.access_token);
    if (skypeResult.ok) {
      updateSkypeCookies(state, skypeResult.value.skypeToken, skypeResult.value.expiresIn,
        skypeSpaces.access_token, skypeSpaces.expires_in);
      skypeTokenRefreshed = true;
    } else {
      log.warn('token-refresh-http', `Skype token exchange failed: ${skypeResult.error.message}`);
    }
  }

  writeSessionState(state);

  // Clear the in-memory token cache so it re-reads from the updated session
  if (resource === 'core') clearTokenCache();

  log.info('token-refresh-http', `Refreshed ${tokensRefreshed} of ${scopes.length} tokens` +
    (skypeTokenRefreshed ? ', skype token refreshed' : '') +
    (currentRefreshToken !== cacheInfo.refreshToken ? ', refresh token rotated' : ''));
  return ok(undefined);
}
