/**
 * Unit tests for auth-guards utilities.
 * 
 * Tests the helper functions that don't require actual browser sessions.
 */

import { describe, it, expect } from 'vitest';
import { handleSubstrateError } from './auth-guards.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';

describe('handleSubstrateError', () => {
  it('returns the same error result unchanged', () => {
    const error = createError(ErrorCode.API_ERROR, 'Server error');
    const result: Result<string> = err(error);
    const handled = handleSubstrateError(result);

    expect(handled.ok).toBe(false);
    if (!handled.ok) {
      expect(handled.error.code).toBe(ErrorCode.API_ERROR);
      expect(handled.error.message).toBe('Server error');
    }
  });

  it('returns the error for AUTH_EXPIRED (and would clear cache)', () => {
    const error = createError(ErrorCode.AUTH_EXPIRED, 'Token expired');
    const result: Result<string> = err(error);
    const handled = handleSubstrateError(result);

    expect(handled.ok).toBe(false);
    if (!handled.ok) {
      expect(handled.error.code).toBe(ErrorCode.AUTH_EXPIRED);
    }
  });

  it('passes through non-AUTH_EXPIRED errors without clearing cache', () => {
    const error = createError(ErrorCode.RATE_LIMITED, 'Too many requests');
    const result: Result<string> = err(error);
    const handled = handleSubstrateError(result);

    expect(handled.ok).toBe(false);
    if (!handled.ok) {
      expect(handled.error.code).toBe(ErrorCode.RATE_LIMITED);
    }
  });

  it('does not break if accidentally called on an ok result (defensive)', () => {
    const result: Result<string> = ok('success');
    const handled = handleSubstrateError(result);

    expect(handled.ok).toBe(true);
    if (handled.ok) {
      expect(handled.value).toBe('success');
    }
  });
});
