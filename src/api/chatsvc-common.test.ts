/**
 * Unit tests for chatsvc-common shared utilities.
 */

import { describe, it, expect } from 'vitest';
import { formatHumanReadableDate } from './chatsvc-common.js';

describe('formatHumanReadableDate', () => {
  it('formats a valid ISO timestamp with day of week', () => {
    // 2026-01-20 is a Tuesday
    const result = formatHumanReadableDate('2026-01-20T14:30:00.000Z');
    expect(result).toContain('Tuesday');
    expect(result).toContain('January');
    expect(result).toContain('20');
    expect(result).toContain('2026');
    expect(result).toContain('UTC');
  });

  it('formats another date correctly', () => {
    // 2026-01-23 is a Friday
    const result = formatHumanReadableDate('2026-01-23T09:00:00.000Z');
    expect(result).toContain('Friday');
    expect(result).toContain('January');
    expect(result).toContain('23');
  });

  it('returns empty string for invalid timestamp', () => {
    expect(formatHumanReadableDate('not-a-date')).toBe('');
  });

  it('returns empty string for empty string', () => {
    expect(formatHumanReadableDate('')).toBe('');
  });

  it('handles epoch timestamp string', () => {
    // 1970-01-01 is a Thursday
    const result = formatHumanReadableDate('1970-01-01T00:00:00.000Z');
    expect(result).toContain('Thursday');
    expect(result).toContain('January');
    expect(result).toContain('1970');
  });
});
