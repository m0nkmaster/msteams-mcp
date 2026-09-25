/**
 * Unit tests for encryption utilities.
 */

import { describe, it, expect } from 'vitest';
import { encrypt, decrypt, type EncryptedData } from './crypto.js';

describe('encrypt/decrypt round-trip', () => {
  it('encrypts and decrypts a simple string', () => {
    const plaintext = 'hello world';
    const encrypted = encrypt(plaintext);
    const decrypted = decrypt(encrypted);
    expect(decrypted).toBe(plaintext);
  });

  it('encrypts and decrypts JSON data', () => {
    const data = JSON.stringify({ token: 'abc123', expiry: 1705850000 });
    const encrypted = encrypt(data);
    const decrypted = decrypt(encrypted);
    expect(decrypted).toBe(data);
  });

  it('encrypts and decrypts empty string', () => {
    const encrypted = encrypt('');
    const decrypted = decrypt(encrypted);
    expect(decrypted).toBe('');
  });

  it('encrypts and decrypts unicode content', () => {
    const text = 'Hello 🌍 — "quotes" & <tags>';
    const encrypted = encrypt(text);
    const decrypted = decrypt(encrypted);
    expect(decrypted).toBe(text);
  });

  it('produces different ciphertext for same plaintext (unique IV)', () => {
    const plaintext = 'same input';
    const encrypted1 = encrypt(plaintext);
    const encrypted2 = encrypt(plaintext);
    // IVs should differ, so ciphertext differs
    expect(encrypted1.iv).not.toBe(encrypted2.iv);
    expect(encrypted1.content).not.toBe(encrypted2.content);
  });
});

describe('decrypt error handling', () => {
  it('throws on tampered ciphertext', () => {
    const encrypted = encrypt('secret');
    const tampered: EncryptedData = {
      ...encrypted,
      content: encrypted.content.slice(0, -2) + 'xx',
    };
    expect(() => decrypt(tampered)).toThrow();
  });

  it('throws on tampered auth tag', () => {
    const encrypted = encrypt('secret');
    const tampered: EncryptedData = {
      ...encrypted,
      tag: 'a'.repeat(encrypted.tag.length),
    };
    expect(() => decrypt(tampered)).toThrow();
  });
});
