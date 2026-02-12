/**
 * Unit tests for Result type utilities.
 */

import { describe, it, expect } from 'vitest';
import { ok, err, unwrap, unwrapOr, map, andThen } from './result.js';
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

describe('unwrap', () => {
  it('returns value for ok result', () => {
    expect(unwrap(ok('hello'))).toBe('hello');
  });

  it('throws for err result', () => {
    const error = createError(ErrorCode.UNKNOWN, 'Something went wrong');
    expect(() => unwrap(err(error))).toThrow('Something went wrong');
  });
});

describe('unwrapOr', () => {
  it('returns value for ok result', () => {
    expect(unwrapOr(ok(42), 0)).toBe(42);
  });

  it('returns default for err result', () => {
    const error = createError(ErrorCode.UNKNOWN, 'fail');
    expect(unwrapOr(err(error), 0)).toBe(0);
  });
});

describe('map', () => {
  it('transforms ok value', () => {
    const result = map(ok(5), x => x * 2);
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toBe(10);
    }
  });

  it('passes through err unchanged', () => {
    const error = createError(ErrorCode.API_ERROR, 'fail');
    const result = map(err(error), () => 'should not run');
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.API_ERROR);
    }
  });
});

describe('andThen', () => {
  it('chains ok results', () => {
    const result = andThen(ok(5), x => ok(x * 2));
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value).toBe(10);
    }
  });

  it('short-circuits on err', () => {
    const error = createError(ErrorCode.NOT_FOUND, 'missing');
    const result = andThen(err(error), () => ok('should not run'));
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.NOT_FOUND);
    }
  });

  it('propagates inner err', () => {
    const innerError = createError(ErrorCode.INVALID_INPUT, 'bad value');
    const result = andThen(ok(5), () => err(innerError));
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.INVALID_INPUT);
    }
  });
});
