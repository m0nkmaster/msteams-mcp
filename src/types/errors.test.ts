/**
 * Unit tests for error taxonomy functions.
 */

import { describe, it, expect } from 'vitest';
import {
  ErrorCode,
  createError,
  classifyHttpError,
  extractRetryAfter,
} from './errors.js';
import { recommendsForceNewWithoutApproval } from '../__fixtures__/force-new-guard.js';

describe('createError', () => {
  it('creates error with defaults for AUTH_REQUIRED', () => {
    const error = createError(ErrorCode.AUTH_REQUIRED, 'Not authenticated');
    expect(error.code).toBe(ErrorCode.AUTH_REQUIRED);
    expect(error.message).toBe('Not authenticated');
    expect(error.retryable).toBe(false);
    expect(error.suggestions).toEqual([
      'Call teams_login (without forceNew) to authenticate: it tries silent SSO first and opens a browser window for sign-in only if needed',
      'Do NOT skip this step or tell the user Teams is unavailable',
      'After login succeeds, retry the original request',
      'Only if that teams_login fails or signs in the wrong account, ask the user to approve teams_login with { "forceNew": true }: it deletes all local login data (browser profile, cookies, every signed-in account) and forces an interactive sign-in',
    ]);
  });

  it('creates error with defaults for AUTH_EXPIRED', () => {
    const error = createError(ErrorCode.AUTH_EXPIRED, 'Token expired');
    expect(error.code).toBe(ErrorCode.AUTH_EXPIRED);
    expect(error.suggestions).toEqual([
      'Call teams_login (without forceNew) to refresh authentication: it tries silent SSO first and opens a browser window for sign-in only if needed',
      'Do NOT skip this step or tell the user Teams is unavailable',
      'After login succeeds, retry the original request',
      'Only if that teams_login fails or signs in the wrong account, ask the user to approve teams_login with { "forceNew": true }: it deletes all local login data (browser profile, cookies, every signed-in account) and forces an interactive sign-in',
    ]);
  });

  it('creates retryable error for RATE_LIMITED by default', () => {
    const error = createError(ErrorCode.RATE_LIMITED, 'Too many requests');
    expect(error.retryable).toBe(true);
  });

  it('creates retryable error for NETWORK_ERROR by default', () => {
    const error = createError(ErrorCode.NETWORK_ERROR, 'Connection failed');
    expect(error.retryable).toBe(true);
  });

  it('creates retryable error for TIMEOUT by default', () => {
    const error = createError(ErrorCode.TIMEOUT, 'Request timed out');
    expect(error.retryable).toBe(true);
  });

  it('creates retryable error for API_ERROR by default', () => {
    const error = createError(ErrorCode.API_ERROR, 'Server error');
    expect(error.retryable).toBe(true);
  });

  it('creates non-retryable error for INVALID_INPUT by default', () => {
    const error = createError(ErrorCode.INVALID_INPUT, 'Bad input');
    expect(error.retryable).toBe(false);
  });

  it('allows overriding retryable flag', () => {
    const error = createError(ErrorCode.AUTH_REQUIRED, 'Auth needed', { retryable: true });
    expect(error.retryable).toBe(true);
  });

  it('allows custom suggestions', () => {
    const error = createError(ErrorCode.UNKNOWN, 'Something broke', {
      suggestions: ['Try again', 'Contact support'],
    });
    expect(error.suggestions).toEqual(['Try again', 'Contact support']);
  });

  it('includes retryAfterMs when provided', () => {
    const error = createError(ErrorCode.RATE_LIMITED, 'Slow down', {
      retryAfterMs: 5000,
    });
    expect(error.retryAfterMs).toBe(5000);
  });
});

describe('recommendsForceNewWithoutApproval', () => {
  it('ignores "without forceNew" phrasing', () => {
    expect(recommendsForceNewWithoutApproval('Call teams_login (without forceNew) to authenticate')).toBe(false);
  });

  it('flags a bare forceNew recommendation', () => {
    expect(recommendsForceNewWithoutApproval('IMMEDIATELY call teams_login with { "forceNew": true }')).toBe(true);
  });

  it('accepts forceNew advice that requires approval', () => {
    expect(recommendsForceNewWithoutApproval('Only with the user\'s approval, call teams_login with { "forceNew": true }')).toBe(false);
  });

  it('flags a negated "without forceNew" warning', () => {
    expect(recommendsForceNewWithoutApproval('Do NOT call teams_login without forceNew — it will return early')).toBe(true);
  });

  it('requires a whole approval phrase, not just the word approval', () => {
    expect(recommendsForceNewWithoutApproval('No approval needed: call teams_login with forceNew: true')).toBe(true);
  });

  it('flags the CLI --force flag', () => {
    expect(recommendsForceNewWithoutApproval('run npm run cli -- login --force true')).toBe(true);
  });

  it('flags the force_new spelling', () => {
    expect(recommendsForceNewWithoutApproval('set force_new: true')).toBe(true);
  });
});

describe('default suggestions', () => {
  it.each(Object.values(ErrorCode))('%s never recommends forceNew without approval', (code) => {
    const error = createError(code, 'x');
    expect(error.suggestions.filter(recommendsForceNewWithoutApproval)).toEqual([]);
  });
});

describe('classifyHttpError', () => {
  it('classifies 401 as AUTH_EXPIRED', () => {
    expect(classifyHttpError(401)).toBe(ErrorCode.AUTH_EXPIRED);
  });

  it('classifies 403 as AUTH_REQUIRED', () => {
    expect(classifyHttpError(403)).toBe(ErrorCode.AUTH_REQUIRED);
  });

  it('classifies 404 as NOT_FOUND', () => {
    expect(classifyHttpError(404)).toBe(ErrorCode.NOT_FOUND);
  });

  it('classifies 429 as RATE_LIMITED', () => {
    expect(classifyHttpError(429)).toBe(ErrorCode.RATE_LIMITED);
  });

  it('classifies 400 as INVALID_INPUT', () => {
    expect(classifyHttpError(400)).toBe(ErrorCode.INVALID_INPUT);
  });

  it('classifies 422 as INVALID_INPUT', () => {
    expect(classifyHttpError(422)).toBe(ErrorCode.INVALID_INPUT);
  });

  it('classifies 500+ as API_ERROR', () => {
    expect(classifyHttpError(500)).toBe(ErrorCode.API_ERROR);
    expect(classifyHttpError(502)).toBe(ErrorCode.API_ERROR);
    expect(classifyHttpError(503)).toBe(ErrorCode.API_ERROR);
  });

  it('classifies timeout message as TIMEOUT', () => {
    expect(classifyHttpError(0, 'request timeout')).toBe(ErrorCode.TIMEOUT);
  });

  it('classifies network message as NETWORK_ERROR', () => {
    expect(classifyHttpError(0, 'network error')).toBe(ErrorCode.NETWORK_ERROR);
    expect(classifyHttpError(0, 'ECONNRESET')).toBe(ErrorCode.NETWORK_ERROR);
  });

  it('classifies unknown status as UNKNOWN', () => {
    expect(classifyHttpError(418)).toBe(ErrorCode.UNKNOWN);
  });
});

describe('extractRetryAfter', () => {
  it('extracts numeric Retry-After in seconds and converts to ms', () => {
    const headers = new Headers({ 'Retry-After': '30' });
    expect(extractRetryAfter(headers)).toBe(30000);
  });

  it('returns undefined when header is missing', () => {
    const headers = new Headers();
    expect(extractRetryAfter(headers)).toBeUndefined();
  });

  it('handles HTTP date format', () => {
    // Use a future date to get a positive value
    const futureDate = new Date(Date.now() + 60000).toUTCString();
    const headers = new Headers({ 'Retry-After': futureDate });
    const result = extractRetryAfter(headers);
    expect(result).toBeDefined();
    expect(result!).toBeGreaterThan(0);
    expect(result!).toBeLessThanOrEqual(60000);
  });
});
