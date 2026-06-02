/**
 * Unit tests for token-extractor helpers.
 *
 * Covers jwtGrantsScope, used to ensure a Graph token actually carries the
 * Calendars.ReadWrite scope before it is used for calendar operations.
 */

import { describe, it, expect } from 'vitest';
import { jwtGrantsScope } from './token-extractor.js';

describe('jwtGrantsScope', () => {
  it('returns true when a delegated scope is present in scp', () => {
    const payload = { scp: 'Calendars.Read Calendars.ReadWrite Mail.Read' };
    expect(jwtGrantsScope(payload, 'Calendars.ReadWrite')).toBe(true);
  });

  it('returns false when the scope is absent from scp', () => {
    const payload = { scp: 'Calendars.Read Mail.Read' };
    expect(jwtGrantsScope(payload, 'Calendars.ReadWrite')).toBe(false);
  });

  it('matches scopes case-insensitively', () => {
    const payload = { scp: 'calendars.readwrite' };
    expect(jwtGrantsScope(payload, 'Calendars.ReadWrite')).toBe(true);
  });

  it('checks app roles when scp is absent', () => {
    const payload = { roles: ['Calendars.ReadWrite', 'Mail.Read'] };
    expect(jwtGrantsScope(payload, 'Calendars.ReadWrite')).toBe(true);
    expect(jwtGrantsScope(payload, 'User.Read.All')).toBe(false);
  });

  it('returns null when scopes cannot be determined', () => {
    expect(jwtGrantsScope({ aud: 'https://graph.microsoft.com' }, 'Calendars.ReadWrite')).toBeNull();
    expect(jwtGrantsScope(null, 'Calendars.ReadWrite')).toBeNull();
  });
});
