/**
 * Tests for messaging cookie selection across Teams hosts — the authtoken
 * moved to teams.cloud.microsoft while skypetoken_asm stayed on
 * .asyncgw.teams.microsoft.com, so domain filtering must match TEAMS_HOSTS
 * (incl. subdomains) and pick the freshest JWT.
 */
import { describe, it, expect } from 'vitest';
import { extractMessageAuth, getMessageAuthStatus } from './token-extractor.js';
import type { SessionState } from './session-store.js';

function fakeJwt(payload: Record<string, unknown>): string {
  const body = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url');
  return `eyJhbGciOiJub25lIn0.${body}.sig`;
}

const EXP_FRESH = Math.floor(Date.now() / 1000) + 60 * 60;
const EXP_STALE = Math.floor(Date.now() / 1000) - 60 * 60 * 24 * 14;

function skypeJwt(exp: number): string {
  return fakeJwt({ exp, skypeid: 'orgid:00000000-0000-0000-0000-000000000000' });
}
function authJwt(exp: number): string {
  return fakeJwt({ exp, oid: '00000000-0000-0000-0000-000000000000' });
}

function cookie(name: string, value: string, domain: string): SessionState['cookies'][number] {
  return { name, value, domain };
}

describe('extractMessageAuth', () => {
  it('finds authtoken on teams.cloud.microsoft and skypetoken on a Teams subdomain', () => {
    const state: SessionState = {
      cookies: [
        cookie('skypetoken_asm', skypeJwt(EXP_FRESH), '.asyncgw.teams.microsoft.com'),
        cookie('authtoken', encodeURIComponent(authJwt(EXP_FRESH)), 'teams.cloud.microsoft'),
      ],
      origins: [],
    };
    const auth = extractMessageAuth(state);
    expect(auth).not.toBeNull();
    expect(auth?.userMri).toContain('8:orgid:');
  });

  it('picks the freshest authtoken when the same name exists on two Teams hosts', () => {
    const stale = authJwt(EXP_STALE);
    const fresh = authJwt(EXP_FRESH);
    const state: SessionState = {
      cookies: [
        cookie('skypetoken_asm', skypeJwt(EXP_FRESH), '.asyncgw.teams.microsoft.com'),
        cookie('authtoken', encodeURIComponent(stale), 'teams.microsoft.com'),
        cookie('authtoken', encodeURIComponent(fresh), 'teams.cloud.microsoft'),
      ],
      origins: [],
    };
    const auth = extractMessageAuth(state);
    expect(auth).not.toBeNull();
    expect(auth?.authToken).toBe(fresh);
  });

  it('returns null and status hasToken=false when no authtoken exists anywhere', () => {
    const state: SessionState = {
      cookies: [
        cookie('skypetoken_asm', skypeJwt(EXP_FRESH), '.asyncgw.teams.microsoft.com'),
      ],
      origins: [],
    };
    expect(extractMessageAuth(state)).toBeNull();
    expect(getMessageAuthStatus(state).hasToken).toBe(false);
  });
});
