/**
 * Unit tests for calendar-api helpers.
 *
 * Covers toUtcIso, which normalises Microsoft Graph dateTimeTimeZone values into
 * unambiguous ISO 8601 strings (Graph omits the timezone marker on dateTime).
 */

import { describe, it, expect } from 'vitest';
import { toUtcIso } from './calendar-api.js';

describe('toUtcIso', () => {
  it('appends Z to a UTC value that lacks a timezone marker', () => {
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00.0000000', timeZone: 'UTC' }))
      .toBe('2026-06-05T10:00:00.0000000Z');
  });

  it('appends Z when the timeZone field is missing (Graph defaults to UTC via Prefer)', () => {
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00.0000000' }))
      .toBe('2026-06-05T10:00:00.0000000Z');
  });

  it('treats timeZone casing leniently', () => {
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00', timeZone: 'utc' }))
      .toBe('2026-06-05T10:00:00Z');
  });

  it('leaves a value that already ends in Z untouched', () => {
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00Z', timeZone: 'UTC' }))
      .toBe('2026-06-05T10:00:00Z');
  });

  it('leaves a value with an explicit offset untouched', () => {
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00+00:00' }))
      .toBe('2026-06-05T10:00:00+00:00');
    expect(toUtcIso({ dateTime: '2026-06-05T15:30:00+05:30' }))
      .toBe('2026-06-05T15:30:00+05:30');
  });

  it('does not append Z for a non-UTC timezone (avoids mislabelling)', () => {
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00', timeZone: 'Pacific Standard Time' }))
      .toBe('2026-06-05T10:00:00');
  });

  it('returns an empty string for missing input', () => {
    expect(toUtcIso(undefined)).toBe('');
    expect(toUtcIso({})).toBe('');
  });
});
