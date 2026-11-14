/**
 * Decryption of MSAL.js v4+ encrypted localStorage cache entries.
 *
 * MSAL.js v4 encrypts every cached credential it writes to localStorage. The
 * stored value becomes an envelope `{ id, nonce, data, lastUpdatedAt }` instead
 * of the plaintext `{ credentialType, target, secret, ... }` entry, so code
 * that JSON-parses localStorage for tokens finds nothing.
 *
 * The scheme (msal-browser `LocalStorage` + `BrowserCrypto`):
 * - A session cookie `msal.cache.encryption` holds URI-encoded JSON
 *   `{ id, key }`, where `key` is a base64url 256-bit base key.
 * - Per entry: HKDF-SHA256(ikm = base key, salt = nonce, info = context)
 *   derives an AES-256-GCM key; the IV is 12 zero bytes (the key is unique per
 *   nonce); `data` is base64url ciphertext with the 16-byte GCM tag appended.
 * - `context` is the client ID when the cache key contains it, else "".
 * - Entries whose `id` differs from the cookie's id were written under an
 *   older key and cannot be decrypted.
 *
 * Decryption happens in Node on the saved storage state; entries that cannot
 * be decrypted are left untouched.
 */

import { createDecipheriv, hkdfSync } from 'node:crypto';

/** Name of the MSAL cookie holding the localStorage encryption key. */
export const MSAL_ENCRYPTION_COOKIE = 'msal.cache.encryption';

/** Base key parsed from the `msal.cache.encryption` cookie. */
export interface MsalEncryptionKey {
  id: string;
  key: Buffer;
}

/** An encrypted MSAL localStorage value. */
interface EncryptedEnvelope {
  id: string;
  nonce: string;
  data: string;
}

type Cookie = { name: string; value: string };
type LocalStorageItem = { name: string; value: string };
type StorageStateLike = {
  cookies?: Cookie[];
  origins?: Array<{ origin: string; localStorage?: LocalStorageItem[] }>;
};

const GUID_PATTERN = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi;

/** Decodes base64 or base64url, padded or not. */
function decodeBase64(value: string): Buffer {
  return Buffer.from(value.replace(/-/g, '+').replace(/_/g, '/'), 'base64');
}

/** Parses an `msal.cache.encryption` cookie value, or null if malformed. */
export function parseEncryptionCookie(value: string): MsalEncryptionKey | null {
  for (const candidate of [value, safeDecodeUri(value)]) {
    if (!candidate) continue;
    try {
      const parsed = JSON.parse(candidate) as { id?: unknown; key?: unknown };
      if (typeof parsed.id === 'string' && typeof parsed.key === 'string' && parsed.key) {
        const key = decodeBase64(parsed.key);
        if (key.length > 0) return { id: parsed.id, key };
      }
    } catch {
      // try next candidate
    }
  }
  return null;
}

function safeDecodeUri(value: string): string | null {
  try {
    return decodeURIComponent(value);
  } catch {
    return null;
  }
}

/** All distinct encryption keys present in a cookie list. */
export function findEncryptionKeys(cookies: Cookie[] | undefined): MsalEncryptionKey[] {
  const keys: MsalEncryptionKey[] = [];
  for (const cookie of cookies ?? []) {
    if (cookie.name !== MSAL_ENCRYPTION_COOKIE) continue;
    const parsed = parseEncryptionCookie(cookie.value);
    if (parsed && !keys.some(k => k.id === parsed.id && k.key.equals(parsed.key))) keys.push(parsed);
  }
  return keys;
}

/** Returns the envelope if `value` is an MSAL v4 encrypted entry. */
export function parseEncryptedEnvelope(value: string): EncryptedEnvelope | null {
  if (!value || value[0] !== '{') return null;
  try {
    const parsed = JSON.parse(value) as Record<string, unknown>;
    if (
      typeof parsed.id === 'string' &&
      typeof parsed.nonce === 'string' &&
      typeof parsed.data === 'string' &&
      parsed.secret === undefined &&
      parsed.credentialType === undefined
    ) {
      return { id: parsed.id, nonce: parsed.nonce, data: parsed.data };
    }
  } catch {
    // not JSON
  }
  return null;
}

function tryDecrypt(baseKey: Buffer, envelope: EncryptedEnvelope, context: string): string | null {
  try {
    const salt = decodeBase64(envelope.nonce);
    const raw = decodeBase64(envelope.data);
    if (raw.length <= 16) return null;
    const derived = Buffer.from(hkdfSync('sha256', baseKey, salt, Buffer.from(context, 'utf8'), 32));
    const decipher = createDecipheriv('aes-256-gcm', derived, Buffer.alloc(12));
    decipher.setAuthTag(raw.subarray(raw.length - 16));
    return Buffer.concat([decipher.update(raw.subarray(0, raw.length - 16)), decipher.final()]).toString('utf8');
  } catch {
    return null;
  }
}

/**
 * Decrypts one localStorage value. Returns plaintext, or null when the value
 * is not encrypted or no key/context combination authenticates.
 */
export function decryptMsalValue(
  cacheKey: string,
  value: string,
  keys: MsalEncryptionKey[],
): string | null {
  const envelope = parseEncryptedEnvelope(value);
  if (!envelope || keys.length === 0) return null;

  // Context is the client ID if the cache key contains it, else ''. Try every
  // GUID in the key (client ID, tenant, object IDs) plus ''; GCM authentication
  // rejects wrong guesses.
  const contexts = [...new Set([...(cacheKey.match(GUID_PATTERN) ?? []), ''])];
  const ordered = [...keys.filter(k => k.id === envelope.id), ...keys.filter(k => k.id !== envelope.id)];

  for (const key of ordered) {
    for (const context of contexts) {
      const plain = tryDecrypt(key.key, envelope, context);
      if (plain !== null) return plain;
    }
  }
  return null;
}

/** Counts from {@link decryptStorageState}; no secret material. */
export interface DecryptStats {
  encrypted: number;
  decrypted: number;
}

/**
 * Replaces MSAL v4 encrypted localStorage values with their plaintext, in
 * place. Values that cannot be decrypted are kept as they are, so nothing is
 * lost. Returns counts only.
 */
export function decryptStorageState<T extends StorageStateLike>(state: T): DecryptStats {
  const stats: DecryptStats = { encrypted: 0, decrypted: 0 };
  const keys = findEncryptionKeys(state.cookies);

  for (const origin of state.origins ?? []) {
    for (const item of origin.localStorage ?? []) {
      if (!parseEncryptedEnvelope(item.value)) continue;
      stats.encrypted++;
      const plain = decryptMsalValue(item.name, item.value, keys);
      if (plain !== null) {
        item.value = plain;
        stats.decrypted++;
      }
    }
  }
  return stats;
}
