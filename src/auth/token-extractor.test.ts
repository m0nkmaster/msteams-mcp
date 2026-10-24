import { describe, expect, it } from 'vitest';
import { extractGraphToken } from './token-extractor.js';
import type { SessionState } from './session-store.js';

function jwt(id: string, aud: string, exp: number): string {
  const header = Buffer.from(JSON.stringify({ alg: 'none' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify({ id, aud, exp })).toString('base64url');
  return `${header}.${payload}.signature`;
}

describe('extractGraphToken', () => {
  it('selects a valid Graph audience over expired and foreign tokens', () => {
    const now = Math.floor(Date.now() / 1000);
    const valid = jwt('valid', 'https://graph.microsoft.com', now + 3600);
    const state: SessionState = {
      cookies: [],
      origins: [{
        origin: 'https://teams.microsoft.com',
        localStorage: [
          { name: 'expired', value: JSON.stringify({ credentialType: 'AccessToken', target: 'https://graph.microsoft.com/.default', secret: jwt('expired', 'https://graph.microsoft.com', now - 1) }) },
          { name: 'foreign', value: JSON.stringify({ credentialType: 'AccessToken', target: 'https://notgraph.microsoft.com/.default', secret: jwt('foreign', 'https://notgraph.microsoft.com', now + 7200) }) },
          { name: 'graph', value: JSON.stringify({ credentialType: 'AccessToken', target: 'https://graph.microsoft.com/.default', secret: valid }) },
        ],
      }],
    };

    expect(extractGraphToken(state)?.token).toBe(valid);
  });
});
