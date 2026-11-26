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
vi.mock('../auth/token-refresh.js', () => ({ refreshCoreTokensViaHttp: vi.fn(), awaitRefreshQueue: vi.fn() }));
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
import { awaitRefreshQueue, refreshCoreTokensViaHttp } from '../auth/token-refresh.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';
import { ErrorCode } from '../types/errors.js';

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
    vi.mocked(refreshCoreTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
    vi.mocked(awaitRefreshQueue).mockResolvedValue(undefined);
  });

  it('mints the missing CSA token when the session is otherwise valid', async () => {
    sessionMissingCsaToken();

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
    expect(refreshCoreTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
  });

  it('skips the refresh when every core credential is present', async () => {
    sessionMissingCsaToken();
    vi.mocked(extractCsaToken).mockReturnValue('csa');

    await loginTool.handler({ forceNew: false });

    expect(refreshCoreTokensViaHttp).not.toHaveBeenCalled();
  });

  it('tops up after a forced browser login leaves the CSA token behind', async () => {
    sessionMissingCsaToken();
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await loginTool.handler({ forceNew: true });

    expect(result.success).toBe(true);
    expect(refreshCoreTokensViaHttp).toHaveBeenCalledOnce();
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
    vi.mocked(refreshCoreTokensViaHttp).mockResolvedValue({
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
    vi.mocked(refreshCoreTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
    vi.mocked(awaitRefreshQueue).mockResolvedValue(undefined);
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: true, minutesRemaining: 60 });
    vi.mocked(extractSubstrateToken).mockReturnValue({ token: 'substrate', expiry: new Date(Date.now() + 3.6e6) });
    vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
    vi.mocked(extractCsaToken).mockReturnValue('csa');
  });

  it('fails with API_ERROR when headless SSO cannot mint messaging credentials', async () => {
    vi.mocked(extractMessageAuth).mockReturnValue(null);
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false });
    vi.mocked(refreshCoreTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe(ErrorCode.API_ERROR);
      expect(result.error.retryable).toBe(false);
      expect(result.error.message).toContain('skypetoken_asm');
      expect(result.error.message).not.toContain('Already authenticated');
      expect(result.error.suggestions?.length).toBeGreaterThan(0);
    }
    // Headless SSO ran once; the visible browser must not launch — a visible
    // sign-in cannot mint the cookie the HTTP exchange failed to mint.
    expect(createBrowserContext).toHaveBeenCalledTimes(1);
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
  });

  it('reports success when the post-sign-in refresh mints messaging credentials', async () => {
    let fixed = false;
    let refreshCalls = 0;
    vi.mocked(refreshCoreTokensViaHttp).mockImplementation(async () => {
      refreshCalls++;
      // The fast-path refresh runs before headless SSO and cannot fix the
      // session; the post-sign-in check runs it a second time, which does.
      if (refreshCalls >= 2) fixed = true;
      return { ok: true, value: undefined };
    });
    vi.mocked(extractMessageAuth).mockImplementation(() =>
      fixed ? { skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' } : null);
    vi.mocked(getMessageAuthStatus).mockImplementation(() =>
      fixed ? { hasToken: true, minutesRemaining: 600 } : { hasToken: false });

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
    if (result.success) {
      expect(String(result.data.message)).toContain('silently');
    }
    expect(createBrowserContext).toHaveBeenCalledTimes(1);
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
  });

  it('fails when the messaging cookie is present but expired and the refresh fails', async () => {
    // Presence alone is not enough: an expired skypetoken_asm leaves messaging
    // unusable, and the old presence-only top-up check reported success here.
    vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false, expiresAt: '2026-01-01T00:00:00.000Z', minutesRemaining: 0 });
    vi.mocked(refreshCoreTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe(ErrorCode.API_ERROR);
      expect(result.error.message).toContain('skypetoken_asm');
    }
    expect(createBrowserContext).toHaveBeenCalledTimes(1);
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
  });

  it('fails after a visible sign-in when messaging credentials stay unusable', async () => {
    vi.mocked(extractMessageAuth).mockReturnValue(null);
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false });
    vi.mocked(refreshCoreTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);
    vi.mocked(ensureAuthenticated)
      .mockRejectedValueOnce(new Error('headless SSO needs interaction'))
      .mockResolvedValue(undefined);

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe(ErrorCode.API_ERROR);
      expect(result.error.retryable).toBe(false);
    }
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: false });
  });

  it('does not open the headless browser until an in-flight refresh settles', async () => {
    // No valid Substrate token: straight to the browser path, which must
    // first wait out the serialized refresh queue (Chromium profile lock).
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false });
    vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: true, minutesRemaining: 600 });
    let release!: () => void;
    const gate = new Promise<void>(resolve => { release = resolve; });
    vi.mocked(awaitRefreshQueue).mockReturnValue(gate);

    const pending = loginTool.handler({ forceNew: false });
    await new Promise(resolve => setTimeout(resolve, 10));
    expect(createBrowserContext).not.toHaveBeenCalled();

    release();
    const result = await pending;
    expect(result.success).toBe(true);
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
  });

  it('refreshes an expired messaging cookie over HTTP without a browser', async () => {
    vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
    vi.mocked(getMessageAuthStatus)
      .mockReturnValueOnce({ hasToken: false, expiresAt: '2026-01-01T00:00:00.000Z', minutesRemaining: 0 })
      .mockReturnValue({ hasToken: true, expiresAt: '2026-11-26T00:00:00.000Z', minutesRemaining: 600 });

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
    expect(refreshCoreTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
    if (result.success) {
      expect(String(result.data.message)).toContain('Already authenticated');
    }
  });

  it('reports the refreshed token lifetime after the HTTP refresh', async () => {
    vi.mocked(getSubstrateTokenStatus)
      .mockReturnValueOnce({ hasToken: true, expiresAt: '2026-11-26T00:00:00.000Z', minutesRemaining: 15 })
      .mockReturnValue({ hasToken: true, expiresAt: '2026-11-26T01:00:00.000Z', minutesRemaining: 70 });
    vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
    vi.mocked(getMessageAuthStatus)
      .mockReturnValueOnce({ hasToken: false, expiresAt: '2026-01-01T00:00:00.000Z', minutesRemaining: 0 })
      .mockReturnValue({ hasToken: true, expiresAt: '2026-11-26T00:00:00.000Z', minutesRemaining: 600 });

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(true);
    if (result.success) {
      expect(String(result.data.message)).toContain('70');
    }
  });

  it('fails when the HTTP refresh succeeds but messaging is still unusable after headless SSO', async () => {
    vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false, expiresAt: '2026-01-01T00:00:00.000Z', minutesRemaining: 0 });

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe(ErrorCode.API_ERROR);
    }
    // One fast-path refresh plus one post-sign-in check; headless only.
    expect(refreshCoreTokensViaHttp).toHaveBeenCalledTimes(2);
    expect(createBrowserContext).toHaveBeenCalledTimes(1);
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
  });

  it.each([
    ['has expired', { hasToken: false, expiresAt: '2026-01-01T00:00:00.000Z', minutesRemaining: 0 }],
    ['is about to expire', { hasToken: true, expiresAt: '2026-01-01T00:00:00.000Z', minutesRemaining: 5 }],
  ])('fails after headless SSO when the messaging cookie %s and refresh fails', async (_label, status) => {
    vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
    vi.mocked(getMessageAuthStatus).mockReturnValue(status);
    vi.mocked(refreshCoreTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);

    const result = await loginTool.handler({ forceNew: false });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe(ErrorCode.API_ERROR);
      expect(result.error.message).not.toContain('Already authenticated');
    }
    expect(createBrowserContext).toHaveBeenCalledTimes(1);
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: true });
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
