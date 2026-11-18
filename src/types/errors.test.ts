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

describe('createError', () => {
  it('creates error with defaults for AUTH_REQUIRED', () => {
    const error = createError(ErrorCode.AUTH_REQUIRED, 'Not authenticated');
    expect(error.code).toBe(ErrorCode.AUTH_REQUIRED);
    expect(error.message).toBe('Not authenticated');
    expect(error.retryable).toBe(false);
    expect(error.suggestions).toContain('Call teams_login to authenticate');
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
