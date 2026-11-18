import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { expect, it, vi } from 'vitest';
import { createServer } from './server.js';

vi.mock('./auth/session-store.js', async importOriginal => ({
  ...await importOriginal<typeof import('./auth/session-store.js')>(),
  readSessionState: () => ({
    cookies: [
      { domain: '.teams.microsoft.com', name: 'skypetoken_asm', value: 'skype-token' },
      { domain: '.teams.microsoft.com', name: 'authtoken', value: 'header.eyJvaWQiOiJ1c2VyIn0.signature' },
    ],
    origins: [{ origin: 'https://teams.microsoft.com', localStorage: [{
      name: 'chatsvcagg.teams.microsoft.com',
      value: JSON.stringify({ secret: 'expired-csa-token', expiresOn: '1' }),
    }] }],
  }),
}));
vi.mock('./browser/context.js', () => ({
  createBrowserContext: async () => { throw new Error('Browser unavailable'); },
}));

it('returns the tool, HTTP refresh and browser errors to the MCP client', async () => {
  const server = await createServer();
  const client = new Client({ name: 'test', version: '1' });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  vi.stubGlobal('fetch', vi.fn().mockRejectedValue(new Error('Unexpected network request')));
  try {
    await server.connect(serverTransport);
    await client.connect(clientTransport);
    const response = await client.callTool({ name: 'teams_get_favorites', arguments: {} });
    expect(response.isError).toBe(true);
    const [content] = response.content as Array<{ text: string }>;
    const failure = JSON.parse(content.text);
    expect(failure.errorCode).toBe('AUTH_REQUIRED');
    expect(failure.error).toContain('CSA token required for favourites is missing or expired');
    expect(failure.error).toContain('refreshToken: false');
    expect(failure.error).toContain('Browser unavailable');
    expect(content.text).not.toContain('forceNew');
  } finally {
    vi.unstubAllGlobals();
    await client.close();
    await server.close();
  }
});
