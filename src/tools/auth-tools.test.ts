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

  extractSkypeSpacesToken: vi.fn(),
}));

vi.mock('../auth/token-refresh-http.js', () => ({ refreshTokensViaHttp: vi.fn() }));
vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', async importOriginal => ({ ...await importOriginal<typeof import('../browser/auth.js')>(), ensureAuthenticated: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ resetAssignmentsAvailability: vi.fn(), clearRegionCache: vi.fn() }));

import { ensureAuthenticated, StaleTeamsSession, TeamsLoginTimeout } from '../browser/auth.js';
import { loginTool } from './auth-tools.js';
import { clearAllLocalData } from '../auth/session-store.js';
import {
  getSubstrateTokenStatus,
  extractMessageAuth,
  extractCsaToken,
  getMessageAuthStatus,
  extractSkypeSpacesToken,
} from '../auth/token-extractor.js';
import { refreshTokensViaHttp } from '../auth/token-refresh-http.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';

/** A session holding every core credential except the CSA (favourites) token. */
function sessionMissingCsaToken(): void {
  vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: true, minutesRemaining: 60 });
  vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: true });
  vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
  vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
  vi.mocked(extractCsaToken).mockReturnValue(null);
}

describe('teams_login core credential top-up', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
  });

  it('recovers expired credentials over HTTP without opening a browser', async () => {
    sessionMissingCsaToken();
    vi.mocked(getSubstrateTokenStatus).mockReturnValueOnce({ hasToken: false });
    const result = await loginTool.handler({ forceNew: false });
    expect(result.success).toBe(true);
    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
  });

  it('uses a clean visible sign-in after stale SSO and closes both windows on failure', async () => {
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false });
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(ensureAuthenticated).mockRejectedValueOnce(new StaleTeamsSession())
      .mockRejectedValueOnce(new TeamsLoginTimeout('sign-in timed out'));
    expect(await loginTool.handler({ forceNew: false })).toMatchObject({
      success: false, error: { code: 'TIMEOUT', message: 'sign-in timed out' },
    });
    expect(createBrowserContext).toHaveBeenLastCalledWith({ headless: false, freshProfile: true });
    expect(closeBrowser).toHaveBeenCalledTimes(2);
    expect(vi.mocked(ensureAuthenticated).mock.calls[1][5]).toBe(vi.mocked(ensureAuthenticated).mock.calls[0][5]);
    expect(clearAllLocalData).not.toHaveBeenCalled();
  });

  it('mints the missing CSA token when the session is otherwise valid', async () => {
    sessionMissingCsaToken();

    vi.mocked(refreshTokensViaHttp).mockImplementationOnce(async () => {
      vi.mocked(extractCsaToken).mockReturnValue('csa');
      return { ok: true, value: undefined };
    });
    const result = await loginTool.handler({ forceNew: false });

    expect(result).toMatchObject({ success: true, data: { capabilities: { favorites: { available: true } } } });
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
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: false, freshProfile: false });
  });

  it('reports usable capabilities and the reason a top-up failed', async () => {
    sessionMissingCsaToken();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);

    const result = await loginTool.handler({ forceNew: false });

    expect(result).toMatchObject({ success: true, data: {
      warning: 'Unavailable capabilities: favorites. Refresh failed: offline',
      capabilities: { messaging: { available: true }, favorites: { available: false } },
    } });
  });

  it('fails login when a saved session still has no usable capabilities after refresh', async () => {
    sessionMissingCsaToken();
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false });
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false });
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);

    const result = await loginTool.handler({ forceNew: true });

    expect(result).toMatchObject({ success: false, error: {
      code: 'AUTH_REQUIRED', message: expect.stringContaining('directApi, messaging, favorites, calendar'),
    } });
  });
});
