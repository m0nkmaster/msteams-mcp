/**
 * Error taxonomy for MCP operations.
 * 
 * Provides machine-readable error codes that help LLMs
 * understand failures and take appropriate action.
 */

/** Enumeration of all error types in the system. */
export enum ErrorCode {
  /** No valid authentication token or session. */
  AUTH_REQUIRED = 'AUTH_REQUIRED',
  /** Token has expired and needs refresh. */
  AUTH_EXPIRED = 'AUTH_EXPIRED',
  /** Rate limited by the API. */
  RATE_LIMITED = 'RATE_LIMITED',
  /** Requested resource was not found. */
  NOT_FOUND = 'NOT_FOUND',
  /** Invalid input parameters. */
  INVALID_INPUT = 'INVALID_INPUT',
  /** API returned an error response. */
  API_ERROR = 'API_ERROR',
  /** Browser automation failed. */
  BROWSER_ERROR = 'BROWSER_ERROR',
  /** Network or connection error. */
  NETWORK_ERROR = 'NETWORK_ERROR',
  /** Operation timed out. */
  TIMEOUT = 'TIMEOUT',
  /** Unknown or unexpected error. */
  UNKNOWN = 'UNKNOWN',
}

/** Structured error with machine-readable information. */
export interface McpError {
  /** Machine-readable error code. */
  code: ErrorCode;
  /** Human-readable error message. */
  message: string;
  /** Whether this error is potentially transient and retryable. */
  retryable: boolean;
  /** Suggested wait time before retry (milliseconds). */
  retryAfterMs?: number;
  /** Suggestions for resolving the error (for LLMs). */
  suggestions: string[];
}

/**
 * Creates a standardised MCP error.
 */
export function createError(
  code: ErrorCode,
  message: string,
  options: {
    retryable?: boolean;
    retryAfterMs?: number;
    suggestions?: string[];
  } = {}
): McpError {
  const defaultSuggestions = getDefaultSuggestions(code);
  
  return {
    code,
    message,
    retryable: options.retryable ?? isRetryableByDefault(code),
    retryAfterMs: options.retryAfterMs,
    suggestions: options.suggestions ?? defaultSuggestions,
  };
}

/**
 * Returns default suggestions for each error code.
 */
function getDefaultSuggestions(code: ErrorCode): string[] {
  switch (code) {
    case ErrorCode.AUTH_REQUIRED:
      return [
        'IMMEDIATELY call teams_login to authenticate',
        'Do NOT skip this step or tell the user Teams is unavailable',
        'After login succeeds, retry the original request',
      ];
    case ErrorCode.AUTH_EXPIRED:
      return [
        'IMMEDIATELY call teams_login to refresh authentication',
        'Do NOT skip this step or tell the user Teams is unavailable',
        'After login succeeds, retry the original request',
      ];
    case ErrorCode.RATE_LIMITED:
      return ['Wait before retrying', 'Reduce request frequency'];
    case ErrorCode.NOT_FOUND:
      return ['Check the ID/query is correct', 'Verify the resource exists'];
    case ErrorCode.INVALID_INPUT:
      return ['Check the input parameters', 'Review the tool documentation'];
    case ErrorCode.API_ERROR:
      return ['Retry the request', 'Check teams_status for system health'];
    case ErrorCode.BROWSER_ERROR:
      return ['Call teams_login to restart browser session'];
    case ErrorCode.NETWORK_ERROR:
      return ['Check network connectivity', 'Retry the request'];
    case ErrorCode.TIMEOUT:
      return ['Retry the request', 'Use smaller page sizes'];
    case ErrorCode.UNKNOWN:
      return ['Check teams_status', 'Try teams_login if authentication issues'];
  }
}

/**
 * Determines if an error code is retryable by default.
 */
function isRetryableByDefault(code: ErrorCode): boolean {
  switch (code) {
    case ErrorCode.RATE_LIMITED:
    case ErrorCode.NETWORK_ERROR:
    case ErrorCode.TIMEOUT:
    case ErrorCode.API_ERROR:
      return true;
    default:
      return false;
  }
}

/**
 * Classifies an HTTP status code into an error code.
 */
export function classifyHttpError(status: number, message?: string): ErrorCode {
  switch (status) {
    case 401:
      return ErrorCode.AUTH_EXPIRED;
    case 403:
      return ErrorCode.AUTH_REQUIRED;
    case 404:
      return ErrorCode.NOT_FOUND;
    case 429:
      return ErrorCode.RATE_LIMITED;
    case 400:
    case 422:
      return ErrorCode.INVALID_INPUT;
    default:
      if (status >= 500) return ErrorCode.API_ERROR;
      if (message?.includes('timeout')) return ErrorCode.TIMEOUT;
      if (message?.includes('network') || message?.includes('ECONNRESET')) {
        return ErrorCode.NETWORK_ERROR;
      }
      return ErrorCode.UNKNOWN;
  }
}

/**
 * Extracts retry-after value from response headers.
 */
export function extractRetryAfter(headers: Headers): number | undefined {
  const retryAfter = headers.get('Retry-After');
  if (!retryAfter) return undefined;
  
  // Could be seconds (number) or HTTP date
  const seconds = parseInt(retryAfter, 10);
  if (!isNaN(seconds)) {
    return seconds * 1000;
  }
  
  // Try parsing as HTTP date
  try {
    const date = new Date(retryAfter);
    return Math.max(0, date.getTime() - Date.now());
  } catch {
    return undefined;
  }
}
