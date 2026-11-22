/**
 * Unit tests for numeric option clamping.
 */

import { describe, it, expect } from 'vitest';
import { clampInt } from './numbers.js';

describe('clampInt', () => {
  it('returns the fallback for absent or non-finite values', () => {
    expect(clampInt(undefined, 30, 1, 200)).toBe(30);
    expect(clampInt(Number.NaN, 30, 1, 200)).toBe(30);
    expect(clampInt(Number.POSITIVE_INFINITY, 30, 1, 200)).toBe(30);
  });

  it('floors fractions and clamps into range', () => {
    expect(clampInt(12.9, 30, 1, 200)).toBe(12);
    expect(clampInt(0, 30, 1, 200)).toBe(1);
    expect(clampInt(500, 30, 1, 200)).toBe(200);
  });
});
