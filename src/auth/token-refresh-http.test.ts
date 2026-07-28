/**
 * Tests for browserless HTTP token refresh.
 * 
 * Tests the MSAL cache extraction and session state update logic.
 * Network calls are mocked to avoid hitting real Azure AD endpoints.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// We need to mock the session-store and token-extractor modules
// before importing the module under test.
vi.mock('./session-store.js', () => ({
  readSessionState: vi.fn(),
  writeSessionState: vi.fn(),
  getTeamsOrigin: vi.fn(),
}));

vi.mock('./token-extractor.js', () => ({
  clearTokenCache: vi.fn(),
}));

import { refreshTokensViaHttp } from './token-refresh-http.js';
import { readSessionState, writeSessionState, getTeamsOrigin } from './session-store.js';
import { clearTokenCache } from './token-extractor.js';
import type { SessionState } from './session-store.js';

// ============================================================================
// Test Fixtures
// ============================================================================

const MOCK_HOME_ACCOUNT_ID = 'user-guid.tenant-guid';
const MOCK_CLIENT_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const MOCK_TENANT_ID = 'tenant-guid';
const MOCK_ENVIRONMENT = 'login.windows.net';

/** Creates a minimal MSAL refresh token localStorage entry. */
function makeRefreshTokenEntry() {
  return {
    name: `${MOCK_HOME_ACCOUNT_ID}-${MOCK_ENVIRONMENT}-refreshtoken-${MOCK_CLIENT_ID}----`,
    value: JSON.stringify({
      credentialType: 'RefreshToken',
      homeAccountId: MOCK_HOME_ACCOUNT_ID,
      environment: MOCK_ENVIRONMENT,
      clientId: MOCK_CLIENT_ID,
      secret: 'mock-refresh-token-secret',
      expiresOn: String(Math.floor(Date.now() / 1000) + 86400),
      lastUpdatedAt: String(Date.now()),
    }),
  };
}

/** Creates a minimal MSAL access token localStorage entry. */
function makeAccessTokenEntry(resource: string, target: string, expiresInSeconds = 3600) {
  const now = Math.floor(Date.now() / 1000);
  return {
    name: `${MOCK_HOME_ACCOUNT_ID}-${MOCK_ENVIRONMENT}-accesstoken-${MOCK_CLIENT_ID}-${MOCK_TENANT_ID}-${target.toLowerCase()}`,
    value: JSON.stringify({
      credentialType: 'AccessToken',
      homeAccountId: MOCK_HOME_ACCOUNT_ID,
      environment: MOCK_ENVIRONMENT,
      clientId: MOCK_CLIENT_ID,
      realm: MOCK_TENANT_ID,
      target,
      tokenType: 'Bearer',
      secret: 'old-access-token',
      expiresOn: String(now + expiresInSeconds),
      extendedExpiresOn: String(now + expiresInSeconds + 3600),
      cachedAt: String(now),
    }),
  };
}

/** Creates a mock session state with MSAL cache entries. */
function makeMockSessionState(): SessionState {
  return {
    cookies: [
      {
        name: 'skypetoken_asm',
        value: 'old-skype-token',
        domain: '.asyncgw.teams.microsoft.com',
        path: '/',
        expires: Date.now() / 1000 + 3600,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
      {
        name: 'skypetoken_asm',
        value: 'old-skype-token',
        domain: '.asm.skype.com',
        path: '/',
        expires: Date.now() / 1000 + 3600,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
      {
        name: 'authtoken',
        value: 'Bearer%3Dold-auth-token',
        domain: 'teams.microsoft.com',
        path: '/',
        expires: Date.now() / 1000 + 3600,
        httpOnly: false,
        secure: true,
        sameSite: 'None',
      },
    ],
    origins: [
      {
        origin: 'https://teams.microsoft.com',
        localStorage: [
          makeRefreshTokenEntry(),
          makeAccessTokenEntry('substrate.office.com', 'https://substrate.office.com/.default'),
          makeAccessTokenEntry('api.spaces.skype.com', 'https://api.spaces.skype.com/.default'),
          makeAccessTokenEntry('chatsvcagg.teams.microsoft.com', 'https://chatsvcagg.teams.microsoft.com/.default'),
        ],
      },
    ],
  };
}

/** Creates a mock Azure AD token response. */
function makeTokenResponse(scope: string, expiresIn = 3600) {
  return {
    access_token: `new-access-token-for-${scope}`,
    refresh_token: 'new-refresh-token',
    token_type: 'Bearer',
    expires_in: expiresIn,
    scope,
    ext_expires_in: expiresIn + 3600,
  };
}

// ============================================================================
// Tests
// ============================================================================

describe('refreshTokensViaHttp', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal('fetch', vi.fn());
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('returns AUTH_REQUIRED when no session state exists', async () => {
    vi.mocked(readSessionState).mockReturnValue(null);

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_REQUIRED');
      expect(result.error.message).toContain('No session state found');
    }
  });

  it('returns AUTH_REQUIRED when no refresh token in session', async () => {
    const state = makeMockSessionState();
    // Remove the refresh token entry
    const origin = state.origins[0];
    origin.localStorage = origin.localStorage.filter(item => {
      try {
        const val = JSON.parse(item.value);
        return val.credentialType !== 'RefreshToken';
      } catch { return true; }
    });

    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_REQUIRED');
      expect(result.error.message).toContain('No MSAL refresh token');
    }
  });

  it('refreshes all tokens successfully', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    // Mock Azure AD token endpoint responses
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('mocked-scope')), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'new-skype-token', expiresIn: 86400 },
        }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('unexpected URL', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.tokensRefreshed).toBe(4);
      expect(result.value.skypeTokenRefreshed).toBe(true);
      expect(result.value.refreshTokenRotated).toBe(true);
    }

    // Verify session state was written back
    expect(writeSessionState).toHaveBeenCalledOnce();

    // Verify token cache was cleared
    expect(clearTokenCache).toHaveBeenCalledOnce();

    // Verify fetch was called 5 times (4 token refreshes + 1 skype exchange)
    expect(vi.mocked(fetch)).toHaveBeenCalledTimes(5);
  });

  it('updates skypetoken_asm cookies in session state', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'brand-new-skype-token', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    await refreshTokensViaHttp();

    // Check the session state that was written
    const writtenState = vi.mocked(writeSessionState).mock.calls[0][0];
    const skypeCookies = writtenState.cookies.filter(
      (c: { name: string }) => c.name === 'skypetoken_asm'
    );
    expect(skypeCookies.length).toBe(2);
    for (const cookie of skypeCookies) {
      expect(cookie.value).toBe('brand-new-skype-token');
    }
  });

  it('handles Azure AD error (expired refresh token)', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async () => {
      return new Response(JSON.stringify({
        error: 'invalid_grant',
        error_description: 'AADSTS700082: The refresh token has expired.',
      }), { status: 400 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_EXPIRED');
      expect(result.error.message).toContain('refresh token has expired');
    }
  });

  it('continues with remaining scopes if one fails with network error', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    let callCount = 0;
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        callCount++;
        if (callCount === 1) {
          // First scope fails with network error
          throw new Error('ECONNRESET');
        }
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'new-skype-token', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) {
      // 3 of 4 scopes succeeded (first one failed with network error)
      expect(result.value.tokensRefreshed).toBe(3);
    }
  });

  it('handles skype token exchange failure gracefully', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response('Forbidden', { status: 403 });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    // Should still succeed — skype token failure is non-fatal
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.tokensRefreshed).toBe(4);
      expect(result.value.skypeTokenRefreshed).toBe(false);
    }
  });

  it('updates MSAL access token cache entries in localStorage', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify({
          access_token: 'fresh-substrate-token',
          refresh_token: 'new-rt',
          token_type: 'Bearer',
          expires_in: 7200,
          scope: 'https://substrate.office.com/.default',
          ext_expires_in: 10800,
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'st', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    await refreshTokensViaHttp();

    // Check that the Substrate access token was updated in localStorage
    const writtenState = vi.mocked(writeSessionState).mock.calls[0][0];
    const origin = writtenState.origins[0];
    const substrateEntry = origin.localStorage.find((item: { name: string; value: string }) => {
      try {
        const val = JSON.parse(item.value);
        return val.credentialType === 'AccessToken' && val.target?.includes('substrate.office.com');
      } catch { return false; }
    });

    expect(substrateEntry).toBeDefined();
    const parsed = JSON.parse(substrateEntry!.value);
    expect(parsed.secret).toBe('fresh-substrate-token');
  });

  it('updates refresh token when Azure AD rotates it', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify({
          access_token: 'new-at',
          refresh_token: 'rotated-refresh-token',
          token_type: 'Bearer',
          expires_in: 3600,
          scope: 'scope',
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'st', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    await refreshTokensViaHttp();

    const writtenState = vi.mocked(writeSessionState).mock.calls[0][0];
    const origin = writtenState.origins[0];
    const rtEntry = origin.localStorage.find((item: { name: string; value: string }) => {
      try {
        const val = JSON.parse(item.value);
        return val.credentialType === 'RefreshToken';
      } catch { return false; }
    });

    expect(rtEntry).toBeDefined();
    const parsed = JSON.parse(rtEntry!.value);
    expect(parsed.secret).toBe('rotated-refresh-token');
  });
});
