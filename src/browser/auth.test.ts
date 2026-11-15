import { describe, expect, it, vi } from 'vitest';
import type { Page } from 'playwright';

vi.mock('./context.js', () => ({ saveSessionState: vi.fn() }));
vi.mock('../auth/token-extractor.js', () => ({ extractSubstrateToken: vi.fn() }));
vi.mock('../auth/msal-cache-crypto.js', () => ({ decryptStorageState: vi.fn() }));

import { ensureAuthenticated, getAuthStatus, isTeamsUrl } from './auth.js';
import { extractSubstrateToken } from '../auth/token-extractor.js';
import { decryptStorageState } from '../auth/msal-cache-crypto.js';

/** A page on `url` with no Teams DOM markers and `tokenMinutes` left on MSAL. */
function fakePage(url: string, tokenMinutes: number): Page {
  return {
    url: () => url,
    locator: () => ({ count: async () => 0 }),
    evaluate: async () => tokenMinutes,
  } as unknown as Page;
}

describe('isTeamsUrl', () => {
  it('covers every host Microsoft redirects Teams users to', () => {
    // New Teams settles on the bare host, with no path.
    for (const url of [
      'https://teams.microsoft.com/',
      'https://teams.cloud.microsoft/',
      'https://teams.chat.microsoft.com/',
      'https://teams.microsoft.us/',
      'https://dod.teams.microsoft.us/',
    ]) {
      expect(isTeamsUrl(url), url).toBe(true);
    }
    expect(isTeamsUrl('https://outlook.office.com/')).toBe(false);
  });

  it('matches on the host, whatever path the app has routed to', () => {
    expect(isTeamsUrl('https://teams.microsoft.com/v2/#/conversations/19:abc@thread.tacv2')).toBe(true);
  });
});

describe('getAuthStatus', () => {
  it('detects a completed login from the MSAL token alone', async () => {
    const status = await getAuthStatus(fakePage('https://teams.cloud.microsoft/', 55));

    expect(status.isAuthenticated).toBe(true);
    expect(status.isOnLoginPage).toBe(false);
  });

  it('stays unauthenticated on a Teams host with no token', async () => {
    const status = await getAuthStatus(fakePage('https://teams.cloud.microsoft/', -1));

    expect(status.isAuthenticated).toBe(false);
  });

  it('detects a login whose MSAL v4 tokens are encrypted in localStorage', async () => {
    // In-page scan finds no plaintext token (-1); the Node-side decrypt does.
    const results: unknown[] = [-1, [{ name: 'k', value: '{"id":"a","nonce":"b","data":"c"}' }]];
    const page = {
      url: () => 'https://teams.cloud.microsoft/',
      evaluate: async () => results.shift(),
      context: () => ({ cookies: async () => [] }),
    } as unknown as Page;
    vi.mocked(decryptStorageState).mockReturnValueOnce({ encrypted: 1, decrypted: 1 });
    vi.mocked(extractSubstrateToken).mockReturnValueOnce({ token: 'fake', expiry: new Date(Date.now() + 30 * 60000) });

    const status = await getAuthStatus(page);

    expect(status.isAuthenticated).toBe(true);
  });

  it('reports the login page', async () => {
    const status = await getAuthStatus(fakePage('https://login.microsoftonline.com/common/oauth2', 99));

    expect(status).toMatchObject({ isAuthenticated: false, isOnLoginPage: true });
  });
});

describe('ensureAuthenticated', () => {
  it('keeps a visible browser open for sign-in when Teams loads with expired tokens', async () => {
    // Teams can render without redirecting to login while the Microsoft SSO
    // session is gone; only a fresh MSAL token proves the user has signed in.
    const tokenMinutes = [-1, -1, 55];
    const page = {
      goto: async () => null,
      waitForURL: async () => { throw new Error('timeout'); },
      url: () => 'https://teams.cloud.microsoft/',
      evaluate: async () => tokenMinutes.length > 1 ? tokenMinutes.shift()! : tokenMinutes[0],
      waitForTimeout: async () => {},
    } as unknown as Page;
    const messages: string[] = [];

    await ensureAuthenticated(page, {} as never, msg => messages.push(msg), false, false);

    expect(messages).toContain('Waiting for manual login...');
    expect(messages).toContain('Authentication successful!');
    expect(tokenMinutes).toEqual([55]);
  });
});
