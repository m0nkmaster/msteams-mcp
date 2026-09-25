import { beforeEach, describe, expect, it, vi } from 'vitest';
import { refreshAssignmentsToken } from './token-refresh.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import { extractSubstrateToken, getValidAssignmentsToken } from './token-extractor.js';
import { err, ok } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import { createBrowserContext } from '../browser/context.js';

vi.mock('./token-refresh-http.js', () => ({ refreshTokensViaHttp: vi.fn() }));
vi.mock('./token-extractor.js', () => ({ extractSubstrateToken: vi.fn(), getValidAssignmentsToken: vi.fn() }));
vi.mock('./session-store.js', () => ({ clearTokenCache: vi.fn() }));
vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(async () => ({ page: {}, context: {} })), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn() }));
const refreshed = ok(undefined);

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(extractSubstrateToken).mockReturnValue({ token: 'substrate', expiry: new Date(Date.now() + 3600000) });
});
describe('on-demand Assignments refresh', () => {
  it('never refreshes core credentials or launches a browser for Assignments', async () => {
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(err(createError(ErrorCode.AUTH_EXPIRED, 'Expired refresh token')));
    expect(await refreshAssignmentsToken()).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_INTERACTION_REQUIRED, retryable: false } });
    expect(vi.mocked(refreshTokensViaHttp).mock.calls).toEqual([['assignments']]);
    expect(createBrowserContext).not.toHaveBeenCalled();
  });
  it.each([ErrorCode.AUTH_EXPIRED, ErrorCode.AUTH_REQUIRED])('never returns %s, which would trigger Teams auto-login', async code => {
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(err(createError(code, 'Auth failure')));
    const result = await refreshAssignmentsToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect([ErrorCode.AUTH_EXPIRED, ErrorCode.AUTH_REQUIRED]).not.toContain(result.error.code);
  });
  it('does not interpret successful exchange without a valid audience as feature unavailability', async () => {
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(refreshed);
    vi.mocked(getValidAssignmentsToken).mockReturnValue(null);
    expect(await refreshAssignmentsToken()).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR } });
  });
  it('does not run core refresh on an access refusal', async () => {
    const denied = err(createError(ErrorCode.ACCESS_DENIED, 'Consent required'));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(denied);
    expect(await refreshAssignmentsToken()).toEqual(denied);
    expect(refreshTokensViaHttp).toHaveBeenCalledTimes(1);
  });

});
