/**
 * Authentication guard utilities.
 * 
 * Provides reusable auth checks that return Result types for consistent
 * error handling across API modules.
 */

import { ErrorCode, createError, type McpError } from '../types/errors.js';
import { type Result, err, ok } from '../types/result.js';
import {
  getValidSubstrateToken,
  extractMessageAuth,
  extractCsaToken,
  extractSubstrateToken,
  extractAssignmentsToken,
  extractGraphToken,
  extractSkypeSpacesToken,
  extractRegionConfig,
  getUserProfile,
  type MessageAuthInfo,
  type RegionConfig,
} from '../auth/token-extractor.js';
import { clearTokenCache } from '../auth/session-store.js';
import { DEFAULT_TEAMS_BASE_URL } from './api-config.js';
import { TOKEN_REFRESH_THRESHOLD_MS, ASSIGNMENTS_UNAVAILABLE_TTL_MS } from '../constants.js';
import { refreshTokensViaBrowser, refreshAssignmentsToken, refreshGraphToken } from '../auth/token-refresh.js';

// ─────────────────────────────────────────────────────────────────────────────
// Guard Types
// ─────────────────────────────────────────────────────────────────────────────

/** Authentication info for messaging and CSA APIs. */
export interface CsaAuthInfo {
  auth: MessageAuthInfo;
  csaToken: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Guard Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Requires a valid Substrate token, refreshing first if it is expired or
 * within the refresh threshold. A failed refresh still returns a token that
 * has not yet expired.
 */
export async function requireSubstrateTokenAsync(): Promise<Result<string, McpError>> {
  const substrate = extractSubstrateToken();
  if (!substrate || substrate.expiry.getTime() - Date.now() < TOKEN_REFRESH_THRESHOLD_MS) {
    await refreshTokensViaBrowser();
  }

  const token = getValidSubstrateToken();
  if (!token) {
    return err(createError(
      ErrorCode.AUTH_EXPIRED,
      'ACTION REQUIRED: Teams token expired and automatic refresh failed. You MUST call teams_login to re-authenticate before retrying.',
    ));
  }
  return ok(token);
}

/**
 * Guard for an optional, on-demand token (Assignments, Graph). Returns a cached
 * token while fresh, refreshes it at most once concurrently, and briefly
 * remembers definitive access refusals so non-entitled accounts pay no repeat cost.
 */
function onDemandTokenGuard(
  extract: () => { token: string; expiry: Date } | null,
  refresh: () => Promise<Result<string>>,
) {
  let unavailable: { until: number; error: McpError } | undefined;
  let pending: Promise<Result<string>> | undefined;
  return {
    reset(): void { unavailable = undefined; },
    async require(): Promise<Result<string, McpError>> {
      const current = extract();
      if (current && current.expiry.getTime() - Date.now() >= TOKEN_REFRESH_THRESHOLD_MS) {
        return ok(current.token);
      }
      if (unavailable && Date.now() < unavailable.until) {
        return current ? ok(current.token) : err(unavailable.error);
      }
      if (!pending) {
        pending = refresh().then(result => {
          if (!result.ok && result.error.code === ErrorCode.ACCESS_DENIED) {
            unavailable = { until: Date.now() + ASSIGNMENTS_UNAVAILABLE_TTL_MS, error: result.error };
          }
          return result;
        }).finally(() => { pending = undefined; });
      }
      const result = await pending;
      if (!result.ok && current && current.expiry.getTime() > Date.now()) return ok(current.token);
      return result;
    },
  };
}

const assignmentsGuard = onDemandTokenGuard(extractAssignmentsToken, refreshAssignmentsToken);
const graphGuard = onDemandTokenGuard(extractGraphToken, refreshGraphToken);

/** Forget remembered refusals for optional resources; called on explicit login. */
export function resetAssignmentsAvailability(): void {
  assignmentsGuard.reset();
  graphGuard.reset();
}

/** Require the EDU Assignments token, preserving auth, access and transient errors. */
export function requireAssignmentsTokenAsync(): Promise<Result<string, McpError>> {
  return assignmentsGuard.require();
}

/** Require a Microsoft Graph token (optional; used for file downloads). */
export function requireGraphTokenAsync(): Promise<Result<string, McpError>> {
  return graphGuard.require();
}

/**
 * Requires valid message authentication.
 * Use for chatsvc messaging APIs.
 */
export function requireMessageAuth(): Result<MessageAuthInfo, McpError> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(ErrorCode.AUTH_REQUIRED, 'ACTION REQUIRED: No valid Teams authentication. You MUST call teams_login to authenticate before retrying.'));
  }
  return ok(auth);
}

/**
 * Requires valid CSA authentication (message auth + CSA token).
 * Use for favourites and team list APIs.
 */
export function requireCsaAuth(): Result<CsaAuthInfo, McpError> {
  const auth = extractMessageAuth();
  const csaToken = extractCsaToken();

  if (!auth?.skypeToken || !csaToken) {
    return err(createError(ErrorCode.AUTH_REQUIRED, 'ACTION REQUIRED: No valid authentication for favourites. You MUST call teams_login to authenticate before retrying.'));
  }

  return ok({ auth, csaToken });
}

/** Authentication info for Skype Spaces APIs (calendar, tags, etc.). */
export interface SkypeSpacesAuthInfo {
  skypeToken: string;
  spacesToken: string;
}

/**
 * Requires valid Skype Spaces authentication (Skype token + Spaces token).
 * Use for mt/part APIs (calendar, tags, etc.).
 */
export function requireSkypeSpacesAuth(): Result<SkypeSpacesAuthInfo, McpError> {
  const auth = extractMessageAuth();
  const spacesToken = extractSkypeSpacesToken();

  if (!auth?.skypeToken || !spacesToken) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'API access requires authentication. Please run teams_login.',
      { suggestions: ['Call teams_login to authenticate'] }
    ));
  }

  return ok({ skypeToken: auth.skypeToken, spacesToken });
}

/** Combined Skype Spaces auth + region config for mt/part APIs. */
export interface SkypeSpacesAuthWithConfig {
  skypeToken: string;
  spacesToken: string;
  regionConfig: RegionConfig;
}

/**
 * Requires valid Skype Spaces auth AND region config in one call.
 * Use for mt/part APIs (calendar, tags, profile, etc.).
 *
 * Eliminates the repeated pattern of calling requireSkypeSpacesAuth() and
 * getRegionConfig() separately and null-checking both.
 */
export function requireSkypeSpacesAuthWithConfig(): Result<SkypeSpacesAuthWithConfig, McpError> {
  const authResult = requireSkypeSpacesAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const regionConfig = getRegionConfig();
  if (!regionConfig) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not determine region. Please run teams_login to authenticate.',
      { suggestions: ['Call teams_login to authenticate'] }
    ));
  }
  return ok({ ...authResult.value, regionConfig });
}

// ─────────────────────────────────────────────────────────────────────────────
// Substrate Error Handling
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Handles a failed Substrate/files API response by clearing the token cache
 * when the error indicates an expired token.
 * 
 * Eliminates the repeated pattern:
 * ```
 * if (!response.ok) {
 *   if (response.error.code === ErrorCode.AUTH_EXPIRED) {
 *     clearTokenCache();
 *   }
 *   return response;
 * }
 * ```
 * 
 * @param response - A failed Result (response.ok === false)
 * @returns The same failed Result, unchanged
 */
export function handleSubstrateError<T>(response: Result<T, McpError>): Result<T, McpError> {
  if (!response.ok && response.error.code === ErrorCode.AUTH_EXPIRED) {
    clearTokenCache();
  }
  return response;
}

// ─────────────────────────────────────────────────────────────────────────────
// Region Configuration
// ─────────────────────────────────────────────────────────────────────────────

/** Default region when session config is unavailable. */
const DEFAULT_REGION = 'amer';

/** Cached region config (undefined = not yet extracted, null = extraction failed). */
let cachedRegionConfig: RegionConfig | null | undefined = undefined;

/**
 * Gets the full region config (from DISCOVER-REGION-GTM) including partition
 * and URLs, with caching. Returns null if no valid session.
 */
export function getRegionConfig(): RegionConfig | null {
  if (cachedRegionConfig === undefined) {
    cachedRegionConfig = extractRegionConfig();
  }
  return cachedRegionConfig;
}

/** Gets the user's region (e.g. "amer"), falling back to 'amer'. */
export function getRegion(): string {
  return getRegionConfig()?.region ?? DEFAULT_REGION;
}

/**
 * Gets the Teams base URL (e.g. "https://teams.microsoft.com", or
 * "https://teams.microsoft.us" for GCC), falling back to the commercial default.
 */
export function getTeamsBaseUrl(): string {
  return getRegionConfig()?.teamsBaseUrl ?? DEFAULT_TEAMS_BASE_URL;
}

/**
 * Gets the CSA (chatsvcagg) region from session config.
 *
 * CSA is not always routed like chatsvc: some tenants get a country-level
 * chatsvc region (e.g. "fr") while CSA lives under the wider region ("emea").
 * DISCOVER-REGION-GTM carries the CSA URL explicitly, so prefer it and only
 * fall back to the chatsvc region when it is absent.
 */
export function getCsaRegion(): string {
  const match = getRegionConfig()?.csaServiceUrl.match(/\/api\/csa\/([a-z]+)$/);
  return match?.[1] ?? getRegion();
}

/** API config with region and base URL for constructing API endpoints. */
export interface ApiConfig {
  region: string;
  /** Region segment for CSA URLs; may differ from the chatsvc region. */
  csaRegion: string;
  baseUrl: string;
}

/** Combined message auth + API config for chatsvc operations. */
export interface MessageAuthWithConfig {
  auth: MessageAuthInfo;
  region: string;
  baseUrl: string;
}

/**
 * Requires valid message auth AND returns API config in one call.
 * 
 * Eliminates the repeated 4-line pattern:
 * ```
 * const authResult = requireMessageAuth();
 * if (!authResult.ok) return authResult;
 * const auth = authResult.value;
 * const { region, baseUrl } = getApiConfig();
 * ```
 */
export function requireMessageAuthWithConfig(): Result<MessageAuthWithConfig, McpError> {
  const authResult = requireMessageAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { region, baseUrl } = getApiConfig();
  return ok({ auth: authResult.value, region, baseUrl });
}

/**
 * Gets region and base URL together for API calls.
 * 
 * Shared helper used by all API modules to avoid duplicating
 * the getRegion() + getTeamsBaseUrl() pattern.
 */
export function getApiConfig(): ApiConfig {
  return {
    region: getRegion(),
    csaRegion: getCsaRegion(),
    baseUrl: getTeamsBaseUrl(),
  };
}

/**
 * Clears the cached region config and tenant ID.
 * Call this after login/logout to pick up new session.
 */
export function clearRegionCache(): void {
  cachedRegionConfig = undefined;
  cachedTenantId = undefined;
}

// ─────────────────────────────────────────────────────────────────────────────
// Tenant ID
// ─────────────────────────────────────────────────────────────────────────────

/** Cached tenant ID (undefined = not yet extracted, null = extraction failed). */
let cachedTenantId: string | null | undefined = undefined;

/**
 * Gets the tenant ID from the user's session (JWT tokens).
 * 
 * Required for building reliable Teams deep links.
 * Returns null if no valid session is available.
 */
export function getTenantId(): string | null {
  if (cachedTenantId !== undefined) {
    return cachedTenantId;
  }
  const profile = getUserProfile();
  const tid = profile?.tenantId ?? null;
  cachedTenantId = tid;
  return tid;
}
