import { describe, expect, it } from 'vitest';
import { webcrypto } from 'node:crypto';
import {
  MSAL_ENCRYPTION_COOKIE,
  decryptMsalValue,
  decryptStorageState,
  findEncryptionKeys,
  parseEncryptedEnvelope,
} from './msal-cache-crypto.js';
import { extractSubstrateToken } from './token-extractor.js';
import { getTeamsOrigin, type SessionState } from './session-store.js';

const CLIENT_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const TENANT = '11111111-2222-3333-4444-555555555555';
const subtle = webcrypto.subtle;
const b64url = (bytes: Uint8Array) => Buffer.from(bytes).toString('base64url');

/** Fake, unsigned JWT with the given expiry. */
function fakeJwt(expSeconds: number): string {
  const enc = (o: object) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${enc({ alg: 'none', typ: 'JWT' })}.${enc({ exp: expSeconds, aud: 'https://substrate.office.com' })}.fakesig`;
}

/** Encrypts like msal-browser v4 BrowserCrypto.encrypt (WebCrypto, HKDF + AES-GCM, zero IV). */
async function msalEncrypt(baseKey: Uint8Array, plain: string, context: string) {
  const nonce = webcrypto.getRandomValues(new Uint8Array(16));
  const hkdf = await subtle.importKey('raw', baseKey, 'HKDF', false, ['deriveKey']);
  const key = await subtle.deriveKey(
    { name: 'HKDF', salt: nonce, hash: 'SHA-256', info: new TextEncoder().encode(context) },
    hkdf, { name: 'AES-GCM', length: 256 }, false, ['encrypt'],
  );
  const data = await subtle.encrypt({ name: 'AES-GCM', iv: new Uint8Array(12) }, key, new TextEncoder().encode(plain));
  return { nonce: b64url(nonce), data: b64url(new Uint8Array(data)) };
}

async function fixture() {
  const baseKey = webcrypto.getRandomValues(new Uint8Array(32));
  const id = 'fake-key-id';
  const cookieValue = encodeURIComponent(JSON.stringify({ id, key: b64url(baseKey) }));
  const tokenKey = `msal.2|uid.utid|login.windows.net|accesstoken|${CLIENT_ID}|${TENANT}|https://substrate.office.com/substratesearch-internal.readwrite|`;
  const plaintext = JSON.stringify({
    credentialType: 'AccessToken',
    clientId: CLIENT_ID,
    realm: TENANT,
    target: 'https://substrate.office.com/substratesearch-internal.readwrite',
    secret: fakeJwt(Math.floor(Date.now() / 1000) + 3600),
  });
  const enc = await msalEncrypt(baseKey, plaintext, CLIENT_ID);
  const encryptedValue = JSON.stringify({ id, ...enc, lastUpdatedAt: String(Date.now()) });
  return { baseKey, id, cookieValue, tokenKey, plaintext, encryptedValue };
}

describe('MSAL v4 cache decryption', () => {
  it('recognises the encrypted envelope and not plaintext entries', async () => {
    const f = await fixture();
    expect(parseEncryptedEnvelope(f.encryptedValue)).not.toBeNull();
    expect(parseEncryptedEnvelope(f.plaintext)).toBeNull();
    expect(parseEncryptedEnvelope('not json')).toBeNull();
  });

  it('decrypts a value encrypted the way msal-browser does', async () => {
    const f = await fixture();
    const keys = findEncryptionKeys([{ name: MSAL_ENCRYPTION_COOKIE, value: f.cookieValue }]);
    expect(keys).toHaveLength(1);
    expect(decryptMsalValue(f.tokenKey, f.encryptedValue, keys)).toBe(f.plaintext);
  });

  it('decrypts entries encrypted with an empty context', async () => {
    const f = await fixture();
    const enc = await msalEncrypt(f.baseKey, '{"x":1}', '');
    const keys = findEncryptionKeys([{ name: MSAL_ENCRYPTION_COOKIE, value: f.cookieValue }]);
    expect(decryptMsalValue('msal.account.keys', JSON.stringify({ id: f.id, ...enc }), keys)).toBe('{"x":1}');
  });

  it('keeps entries it cannot decrypt (wrong or missing key)', async () => {
    const f = await fixture();
    const wrongKey = encodeURIComponent(JSON.stringify({ id: f.id, key: b64url(new Uint8Array(32).fill(7)) }));
    for (const cookies of [[], [{ name: MSAL_ENCRYPTION_COOKIE, value: wrongKey }]]) {
      const state = { cookies, origins: [{ origin: 'https://teams.cloud.microsoft', localStorage: [{ name: f.tokenKey, value: f.encryptedValue }] }] };
      expect(decryptStorageState(state)).toEqual({ encrypted: 1, decrypted: 0 });
      expect(state.origins[0].localStorage[0].value).toBe(f.encryptedValue);
    }
  });

  it('makes a lowercase-scope encrypted Substrate token extractable', async () => {
    const f = await fixture();
    const state: SessionState = {
      cookies: [{ name: MSAL_ENCRYPTION_COOKIE, value: f.cookieValue, domain: 'teams.cloud.microsoft' }],
      origins: [
        { origin: 'https://teams.microsoft.com', localStorage: [] },
        { origin: 'https://teams.cloud.microsoft', localStorage: [
          { name: 'unrelated', value: 'plain-string' },
          { name: f.tokenKey, value: f.encryptedValue },
        ] },
      ],
    };
    expect(extractSubstrateToken(state)).toBeNull();

    expect(decryptStorageState(state)).toEqual({ encrypted: 1, decrypted: 1 });
    expect(getTeamsOrigin(state)?.origin).toBe('https://teams.cloud.microsoft');
    expect(extractSubstrateToken(state)?.token).toBe(JSON.parse(f.plaintext).secret);
  });
});
