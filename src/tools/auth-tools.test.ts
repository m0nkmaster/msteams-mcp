import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('../auth/session-store.js', () => ({
  hasSessionState: vi.fn(() => true),
  isSessionLikelyExpired: vi.fn(() => false),
  clearAllLocalData: vi.fn(),
}));

vi.mock('../auth/token-extractor.js', () => ({
  getSubstrateTokenStatus: vi.fn(),
  getMessageAuthStatus: vi.fn(),
  extractMessageAuth: vi.fn(),
  extractCsaToken: vi.fn(),
  extractSubstrateToken: vi.fn(),
  extractSkypeSpacesToken: vi.fn(),
}));

vi.mock('../auth/token-refresh-http.js', () => ({ refreshTokensViaHttp: vi.fn() }));
vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ resetAssignmentsAvailability: vi.fn(), clearRegionCache: vi.fn() }));

import { loginTool } from './auth-tools.js';
import { clearAllLocalData } from '../auth/session-store.js';
import {
  getSubstrateTokenStatus,
  getMessageAuthStatus,
  extractMessageAuth,
  extractCsaToken,
  extractSubstrateToken,
  extractSkypeSpacesToken,
} from '../auth/token-extractor.js';
import { refreshTokensViaHttp } from '../auth/token-refresh-http.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';

/** A session holding every core credential except the CSA (favourites) token. */
function sessionMissingCsaToken(): void {
  vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: true, minutesRemaining: 60 });
  vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: true, minutesRemaining: 600 });
  vi.mocked(extractSubstrateToken).mockReturnValue({ token: 'substrate', expiry: new Date(Date.now() + 3.6e6) });
  vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
  vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
  vi.mocked(extractCsaToken).mockReturnValue(null);
}

describe('teams_login core credential top-up', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
  });

  it('mints the missing CSA token when the session is otherwise valid', async () => {
    sessionMissingCsaToken();

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
  });

  it('skips the refresh when every core credential is present', async () => {
    sessionMissingCsaToken();
    vi.mocked(extractCsaToken).mockReturnValue('csa');

    await loginTool.handler({ forceNew: false });

    expect(refreshTokensViaHttp).not.toHaveBeenCalled();
  });

  it('tops up after a forced browser login leaves the CSA token behind', async () => {
    sessionMissingCsaToken();
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await loginTool.handler({ forceNew: true });

    expect(result.success).toBe(true);
    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
  });

  it('forceNew wipes all previous login data and signs in from scratch in a visible browser', async () => {
    sessionMissingCsaToken();
    vi.mocked(extractCsaToken).mockReturnValue('csa');
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    await loginTool.handler({ forceNew: true });

    expect(clearAllLocalData).toHaveBeenCalledOnce();
    expect(vi.mocked(clearAllLocalData).mock.invocationCallOrder[0])
      .toBeLessThan(vi.mocked(createBrowserContext).mock.invocationCallOrder[0]);
    // Valid-looking tokens from the old session must not short-circuit the login,
    // and a wiped profile has nothing for headless SSO to use.
    expect(createBrowserContext).toHaveBeenCalledOnce();
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: false });
  });

  it('still reports success when the top-up fails', async () => {
    sessionMissingCsaToken();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
  });
});

describe('teams_login fast path requires usable messaging', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: true, minutesRemaining: 60 });
    vi.mocked(extractSubstrateToken).mockReturnValue({ token: 'substrate', expiry: new Date(Date.now() + 3.6e6) });
    vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
    vi.mocked(extractCsaToken).mockReturnValue('csa');
  });

  it('falls through to headless SSO when messaging cookies are missing and the top-up fails', async () => {
    vi.mocked(extractMessageAuth).mockReturnValue(null);
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false });
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
    // Exactly one HTTP exchange before any browser launches (the post-headless
    // top-up may legitimately refresh again after SSO).
    const refreshOrder = vi.mocked(refreshTokensViaHttp).mock.invocationCallOrder;
    const browserOrder = vi.mocked(createBrowserContext).mock.invocationCallOrder[0];
    expect(refreshOrder.filter(order => order < browserOrder)).toHaveLength(1);
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
    if (result.success) {
      expect(String(result.data.message)).not.toContain('Already authenticated');
    }
  });

  it('refreshes an expired messaging cookie over HTTP without a browser', async () => {
    vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
    vi.mocked(getMessageAuthStatus)
      .mockReturnValueOnce({ hasToken: false, expiresAt: '2026-01-01T00:00:00.000Z', minutesRemaining: 0 })
      .mockReturnValue({ hasToken: true, expiresAt: '2026-11-26T00:00:00.000Z', minutesRemaining: 600 });

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
    if (result.success) {
      expect(String(result.data.message)).toContain('Already authenticated');
    }
  });

  it.each([
    ['has expired', { hasToken: false, expiresAt: '2026-01-01T00:00:00.000Z', minutesRemaining: 0 }],
    ['is about to expire', { hasToken: true, expiresAt: '2026-01-01T00:00:00.000Z', minutesRemaining: 5 }],
  ])('falls through to headless SSO when the messaging cookie %s', async (_label, status) => {
    vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
    vi.mocked(getMessageAuthStatus).mockReturnValue(status);
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
    if (result.success) {
      expect(String(result.data.message)).not.toContain('Already authenticated');
    }
  });
});

describe('teams_login tool definition', () => {
  it('warns that forceNew is destructive and needs approval', () => {
    const props = loginTool.definition.inputSchema.properties as Record<string, { description?: string }>;
    expect(props.forceNew.description).toMatch(/destructive/i);
    expect(props.forceNew.description).toMatch(/approv/i);
  });

  it('describes plain login as the default path', () => {
    expect(loginTool.definition.description).toMatch(/without forceNew/i);
  });
});
