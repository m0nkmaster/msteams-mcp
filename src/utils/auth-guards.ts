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
  extractSkypeSpacesToken,
  extractRegionConfig,
  getUserProfile,
  clearTokenCache,
  getValidGraphToken,
  type MessageAuthInfo,
  type RegionConfig,
} from '../auth/token-extractor.js';
import { TOKEN_REFRESH_THRESHOLD_MS } from '../constants.js';
import { refreshTokensViaBrowser } from '../auth/token-refresh.js';

// ─────────────────────────────────────────────────────────────────────────────
// Error Messages
// ─────────────────────────────────────────────────────────────────────────────

const AUTH_ERROR_MESSAGES = {
  messageAuth: 'ACTION REQUIRED: No valid Teams authentication. You MUST call teams_login to authenticate before retrying.',
  csaToken: 'ACTION REQUIRED: No valid authentication for favourites. You MUST call teams_login to authenticate before retrying.',
} as const;

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
 * Checks if the Substrate token needs refresh (expired or approaching expiry).
 * 
 * @returns true if token is expired or will expire within the refresh threshold
 */
function shouldRefreshSubstrateToken(): boolean {
  const substrate = extractSubstrateToken();
  if (!substrate) return false;

  const timeRemaining = substrate.expiry.getTime() - Date.now();
  // Refresh if expired (timeRemaining <= 0) OR approaching expiry
  return timeRemaining < TOKEN_REFRESH_THRESHOLD_MS;
}

/**
 * Requires a valid Substrate token with proactive refresh.
 * 
 * This async version attempts to refresh tokens if they're approaching
 * expiry (within 10 minutes). Use this in tool handlers for better UX.
 */
export async function requireSubstrateTokenAsync(): Promise<Result<string, McpError>> {
  // Check if we need to refresh proactively
  if (shouldRefreshSubstrateToken()) {
    const refreshResult = await refreshTokensViaBrowser();
    if (refreshResult.ok) {
      // Refresh succeeded, get the new token
      const token = getValidSubstrateToken();
      if (token) {
        return ok(token);
      }
    }
    // Refresh failed but token might still be valid, continue
  }

  // Try to get existing token
  const token = getValidSubstrateToken();
  if (!token) {
    // Token expired and refresh not available/failed
    return err(createError(
      ErrorCode.AUTH_EXPIRED,
      'ACTION REQUIRED: Teams token expired and automatic refresh failed. You MUST call teams_login to re-authenticate before retrying.',
    ));
  }

  return ok(token);
}

/**
 * Requires valid message authentication.
 * Use for chatsvc messaging APIs.
 */
export function requireMessageAuth(): Result<MessageAuthInfo, McpError> {
  const auth = extractMessageAuth();
  if (!auth) {
    return err(createError(ErrorCode.AUTH_REQUIRED, AUTH_ERROR_MESSAGES.messageAuth));
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
    return err(createError(ErrorCode.AUTH_REQUIRED, AUTH_ERROR_MESSAGES.csaToken));
  }

  return ok({ auth, csaToken });
}

/** Authentication info for calendar/meetings API. */
export interface CalendarAuthInfo {
  skypeToken: string;
  spacesToken: string;
}

/**
 * Requires valid calendar authentication (Skype token + Spaces token).
 * Use for mt/part calendar APIs.
 */
export function requireCalendarAuth(): Result<CalendarAuthInfo, McpError> {
  const auth = extractMessageAuth();
  const spacesToken = extractSkypeSpacesToken();

  if (!auth?.skypeToken || !spacesToken) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Calendar access requires authentication. Please run teams_login.',
      { suggestions: ['Call teams_login to authenticate'] }
    ));
  }

  return ok({ skypeToken: auth.skypeToken, spacesToken });
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

import { DEFAULT_TEAMS_BASE_URL } from './api-config.js';

/** Default region when session config is unavailable. */
const DEFAULT_REGION = 'amer';

/** Cached region config (undefined = not yet extracted, null = extraction failed). */
let cachedRegionConfig: RegionConfig | null | undefined = undefined;

/**
 * Gets the user's region from session, with caching.
 * 
 * The region is extracted from the DISCOVER-REGION-GTM config in localStorage.
 * Falls back to 'amer' if not available (shouldn't happen with valid session).
 */
export function getRegion(): string {
  if (cachedRegionConfig === undefined) {
    cachedRegionConfig = extractRegionConfig();
  }
  return cachedRegionConfig?.region ?? DEFAULT_REGION;
}

/**
 * Gets the Teams base URL from session config.
 * 
 * Returns the base URL for API calls (e.g., "https://teams.microsoft.com" for
 * commercial cloud, or "https://teams.microsoft.us" for GCC).
 * Falls back to default if config not available.
 */
export function getTeamsBaseUrl(): string {
  if (cachedRegionConfig === undefined) {
    cachedRegionConfig = extractRegionConfig();
  }
  return cachedRegionConfig?.teamsBaseUrl ?? DEFAULT_TEAMS_BASE_URL;
}

/**
 * Gets the full region config including partition and URLs.
 * 
 * Returns null if no valid session - caller should handle auth error.
 */
export function getRegionConfig(): RegionConfig | null {
  if (cachedRegionConfig === undefined) {
    cachedRegionConfig = extractRegionConfig();
  }
  return cachedRegionConfig;
}

/** API config with region and base URL for constructing API endpoints. */
export interface ApiConfig {
  region: string;
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
// Graph API Authentication
// ─────────────────────────────────────────────────────────────────────────────

/** Authentication info for Microsoft Graph API. */
export interface GraphAuthInfo {
  graphToken: string;
}

/**
 * Requires a valid Microsoft Graph API token.
 * Use for Graph API calls (e.g., sending messages via Graph).
 */
export function requireGraphAuth(): Result<GraphAuthInfo, McpError> {
  const graphToken = getValidGraphToken();
  if (!graphToken) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'ACTION REQUIRED: No valid Microsoft Graph token. You MUST call teams_login to authenticate before retrying. The Graph token is acquired during token refresh.',
    ));
  }
  return ok({ graphToken });
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
