import { afterAll, describe, expect, it, vi } from 'vitest';
import { existsSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { CONFIG_DIR, type SessionState } from '../auth/session-store.js';
import { chromium } from 'playwright';
import { closeBrowser, createBrowserContext, saveAuthenticatedSession } from './context.js';

vi.mock('playwright', () => ({ chromium: { launchPersistentContext: vi.fn() }, firefox: {} }));
vi.mock('../auth/session-store.js', async importOriginal => ({
  ...await importOriginal<typeof import('../auth/session-store.js')>(),
  CONFIG_DIR: mkdtempSync(join(tmpdir(), 'msteams-login-')),
  ensureConfigDir: vi.fn(),
  clearTokenCache: vi.fn(),
  readSessionState: () => saved,
  writeSessionState: (state: SessionState) => { saved = state; },
}));

let saved: SessionState | undefined;
afterAll(() => rmSync(CONFIG_DIR, { recursive: true, force: true }));

describe('interactive recovery profile', () => {
  it('preserves the previous session on failure and reuses the verified profile after success', async () => {
    vi.stubEnv('TEAMS_BROWSER', 'chrome');
    const snapshot: SessionState = { origins: [], cookies: [] };
    const context = { pages: () => [{}], close: vi.fn(), storageState: async () => snapshot };
    vi.mocked(chromium.launchPersistentContext).mockResolvedValue(context as never);
    saved = { origins: [], cookies: [] };
    const previous = saved;
    try {
      const failed = await createBrowserContext({ headless: false, freshProfile: true });
      const failedDir = vi.mocked(chromium.launchPersistentContext).mock.lastCall![0];
      expect(await saveAuthenticatedSession(failed.context, true)).toBe(false);
      await closeBrowser(failed);
      expect(existsSync(failedDir)).toBe(false);
      expect(saved).toBe(previous);

      const recovered = await createBrowserContext({ headless: false, freshProfile: true });
      const recoveredDir = vi.mocked(chromium.launchPersistentContext).mock.lastCall![0];
      const payload = Buffer.from(JSON.stringify({ exp: Date.now() / 1000 + 3600 })).toString('base64url');
      snapshot.origins.push({ origin: 'https://teams.cloud.microsoft', localStorage: [{
        name: 'test', value: JSON.stringify({ target: 'api.spaces.skype.com', secret: `eyJ0.${payload}.sig` }),
      }] });
      expect(await saveAuthenticatedSession(recovered.context, false)).toBe(false);
      expect(await saveAuthenticatedSession(recovered.context, true)).toBe(true);
      await closeBrowser(recovered);
      expect(existsSync(recoveredDir)).toBe(true);
      const next = await createBrowserContext();
      expect(vi.mocked(chromium.launchPersistentContext).mock.lastCall![0]).toBe(recoveredDir);
      await closeBrowser(next);
    } finally {
      vi.unstubAllEnvs();
    }
  });
});
