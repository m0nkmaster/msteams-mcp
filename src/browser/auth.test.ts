import { describe, expect, it, vi } from 'vitest';
import type { BrowserContext, Page } from 'playwright';
import { ensureAuthenticated, isTeamsUrl, StaleTeamsSession } from './auth.js';
import { saveAuthenticatedSession } from './context.js';

vi.mock('./context.js', () => ({ saveAuthenticatedSession: vi.fn() }));

describe('isTeamsUrl', () => {
  it('covers every host Microsoft redirects Teams users to', () => {
    // New Teams settles on the bare host, with no path.
    for (const url of [
      'https://teams.microsoft.com/',
      'https://teams.cloud.microsoft/',
      'https://teams.chat.microsoft.com/',
      'https://teams.live.com/v2/',
      'https://teams.microsoft.us/',
      'https://dod.teams.microsoft.us/',
    ]) {
      expect(isTeamsUrl(url), url).toBe(true);
    }
    expect(isTeamsUrl('https://outlook.office.com/')).toBe(false);
    expect(isTeamsUrl('https://example.org/?next=https://teams.live.com')).toBe(false);
  });

  it('matches on the host, whatever path the app has routed to', () => {
    expect(isTeamsUrl('https://teams.microsoft.com/v2/#/conversations/19:abc@thread.tacv2')).toBe(true);
  });
});

describe('ensureAuthenticated', () => {
  function setup(url = 'https://teams.cloud.microsoft') {
    const page = {
      url: () => url,
      goto: vi.fn(),
      waitForURL: vi.fn().mockRejectedValue(new Error('no redirect')),
      evaluate: vi.fn(),
      waitForTimeout: vi.fn(),
    };
    const context = {} as BrowserContext;
    return { page, login: (headless: boolean, deadline?: number) =>
      ensureAuthenticated(page as unknown as Page, context, vi.fn(), !headless, headless, deadline) };
  }

  it('shows waiting progress before success, including on a silent visible sign-in', async () => {
    const { page, login } = setup();
    vi.mocked(saveAuthenticatedSession).mockResolvedValueOnce(false).mockResolvedValueOnce(true);
    await login(false);
    expect(page.evaluate.mock.calls.map(call => call[1].message)).toEqual([
      'Waiting for Teams credentials', 'Waiting for Teams credentials', 'All done!',
    ]);
    expect(saveAuthenticatedSession).toHaveBeenLastCalledWith({}, true);
  });

  it('reports a stale profile after silent SSO expires', async () => {
    const { page, login } = setup();
    await expect(login(true, Date.now() - 1)).rejects.toBeInstanceOf(StaleTeamsSession);
    expect(page.evaluate).not.toHaveBeenCalled();
  });

  it('returns a sign-in timeout without covering the identity provider page', async () => {
    const { page, login } = setup('https://login.microsoftonline.com/common');
    await expect(login(false, Date.now() - 1)).rejects.toThrow('before Teams provided usable credentials');
    expect(page.evaluate).not.toHaveBeenCalled();
  });
});
