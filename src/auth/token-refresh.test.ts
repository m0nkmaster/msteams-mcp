import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('./token-extractor.js', () => ({
  extractSubstrateToken: vi.fn(),
  clearTokenCache: vi.fn(),
}));

vi.mock('./token-refresh-http.js', () => ({
  refreshTokensViaHttp: vi.fn(),
}));

import { refreshTokensViaBrowser } from './token-refresh.js';
import { extractSubstrateToken } from './token-extractor.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';

describe('refreshTokensViaBrowser', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('uses HTTP refresh after the cached access token has expired', async () => {
    const refreshedExpiry = new Date(Date.now() + 60 * 60 * 1000);

    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(null)
      .mockReturnValueOnce({ token: 'refreshed-access-token', expiry: refreshedExpiry });
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: true,
      value: {
        tokensRefreshed: 3,
        skypeTokenRefreshed: true,
        refreshTokenRotated: true,
      },
    });

    const result = await refreshTokensViaBrowser();

    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.method).toBe('http');
      expect(result.value.newExpiry).toEqual(refreshedExpiry);
      expect(result.value.previousExpiry).toBeNull();
      expect(result.value.minutesGained).toBeNull();
      expect(result.value.refreshNeeded).toBe(true);
    }
  });
});
