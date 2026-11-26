/**
 * Tests for getTeamsOrigin — especially New Teams (teams.cloud.microsoft)
 * SubstrateSearch token selection when multiple origins are present.
 */

import { describe, it, expect, vi } from 'vitest';
import { getTeamsOrigin, getSessionEpoch, clearAllLocalData, clearSessionState, type SessionState } from './session-store.js';

// clearAllLocalData/clearSessionState delete under the real CONFIG_DIR;
// stub the destructive fs calls so the epoch tests never touch it.
vi.mock('fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('fs')>();
  return { ...actual, rmSync: vi.fn(), existsSync: vi.fn(() => false), unlinkSync: vi.fn() };
});

function makeOrigin(
  origin: string,
  localStorage: SessionState['origins'][number]['localStorage'] = []
): SessionState['origins'][number] {
  return { origin, localStorage };
}

/** Fake JWT: unsigned header.payload.sig with a base64url JSON payload. */
function fakeJwt(payload: Record<string, unknown>): string {
  const body = Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url');
  return `eyJhbGciOiJub25lIn0.${body}.sig`;
}

function makeSubstrateEntry(
  target = 'https://substrate.office.com/SubstrateSearch-Internal.ReadWrite',
  secret = 'eyJhbGciOiJub25lIn0.eyJzdWIiOiJ0ZXN0In0.sig'
) {
  return {
    name: 'msal-substrate-token',
    value: JSON.stringify({
      credentialType: 'AccessToken',
      target,
      // JWT-shaped secret (header.payload.sig) — extractor only checks prefix
      secret,
    }),
  };
}

function makeChatSvcEntry() {
  return {
    name: 'msal-chatsvc-token',
    value: JSON.stringify({
      credentialType: 'AccessToken',
      target: 'https://chatsvcagg.teams.microsoft.com/.default',
      secret: 'eyJhbGciOiJub25lIn0.eyJzdWIiOiJjaGF0In0.sig',
    }),
  };
}

describe('getTeamsOrigin', () => {
  it('prefers the origin that holds a SubstrateSearch token', () => {
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://teams.microsoft.com', [makeChatSvcEntry()]),
        makeOrigin('https://teams.cloud.microsoft', [makeSubstrateEntry()]),
      ],
    };

    const chosen = getTeamsOrigin(state);
    expect(chosen?.origin).toBe('https://teams.cloud.microsoft');
  });

  it('falls back to teams.cloud.microsoft when no Substrate token exists', () => {
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://teams.microsoft.com', [makeChatSvcEntry()]),
        makeOrigin('https://teams.cloud.microsoft', [makeChatSvcEntry()]),
      ],
    };

    const chosen = getTeamsOrigin(state);
    expect(chosen?.origin).toBe('https://teams.cloud.microsoft');
  });

  it('returns teams.microsoft.com when it is the only known origin', () => {
    const state: SessionState = {
      cookies: [],
      origins: [makeOrigin('https://teams.microsoft.com', [makeChatSvcEntry()])],
    };

    const chosen = getTeamsOrigin(state);
    expect(chosen?.origin).toBe('https://teams.microsoft.com');
  });

  it('prefers the origin with the latest-expiring UNEXPIRED JWT over a stale one listed first', () => {
    const expired = Math.floor(Date.now() / 1000) - 60 * 60 * 24 * 14; // ~2 weeks ago
    const fresh = Math.floor(Date.now() / 1000) + 60 * 60;
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://teams.microsoft.com', [makeSubstrateEntry(undefined, fakeJwt({ exp: expired }))]),
        makeOrigin('https://teams.cloud.microsoft', [makeSubstrateEntry(undefined, fakeJwt({ exp: fresh }))]),
      ],
    };
    expect(getTeamsOrigin(state)?.origin).toBe('https://teams.cloud.microsoft');
  });

  it('picks the fresh origin even when the stale one is listed second', () => {
    const expired = Math.floor(Date.now() / 1000) - 60 * 60 * 24 * 14;
    const fresh = Math.floor(Date.now() / 1000) + 60 * 60;
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://teams.cloud.microsoft', [makeSubstrateEntry(undefined, fakeJwt({ exp: fresh }))]),
        makeOrigin('https://teams.microsoft.com', [makeSubstrateEntry(undefined, fakeJwt({ exp: expired }))]),
      ],
    };
    expect(getTeamsOrigin(state)?.origin).toBe('https://teams.cloud.microsoft');
  });

  it('falls back to the previous behaviour when every JWT is expired', () => {
    const expired = Math.floor(Date.now() / 1000) - 60;
    const state: SessionState = {
      cookies: [],
      origins: [
        makeOrigin('https://teams.microsoft.com', [makeSubstrateEntry(undefined, fakeJwt({ exp: expired }))]),
        makeOrigin('https://teams.cloud.microsoft', [makeChatSvcEntry()]),
      ],
    };
    // Upstream logic: first origin with a SubstrateSearch token.
    expect(getTeamsOrigin(state)?.origin).toBe('https://teams.microsoft.com');
  });

  it('returns null when no Teams origins are present', () => {
    const state: SessionState = {
      cookies: [],
      origins: [makeOrigin('https://login.microsoftonline.com', [])],
    };

    expect(getTeamsOrigin(state)).toBeNull();
  });
});

describe('session epoch', () => {
  it('bumps when local data is deleted so an in-flight refresh can detect the wipe', () => {
    const before = getSessionEpoch();
    clearAllLocalData();
    expect(getSessionEpoch()).toBe(before + 1);
    clearSessionState();
    expect(getSessionEpoch()).toBe(before + 2);
  });
});
