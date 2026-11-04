/**
 * Unit tests for Result type utilities.
 */

import { describe, it, expect } from 'vitest';
import { ok, err } from './result.js';
import { ErrorCode, createError } from './errors.js';

describe('ok', () => {
  it('creates a successful result', () => {
    const result = ok(42);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toBe(42);
    }
  });

  it('works with complex types', () => {
    const result = ok({ name: 'test', count: 5 });
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.name).toBe('test');
      expect(result.value.count).toBe(5);
    }
  });
});

describe('err', () => {
  it('creates a failed result', () => {
    const error = createError(ErrorCode.NOT_FOUND, 'Not found');
    const result = err(error);
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.NOT_FOUND);
    }
  });
});
