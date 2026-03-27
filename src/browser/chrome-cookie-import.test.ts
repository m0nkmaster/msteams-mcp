import { describe, it, expect } from 'vitest';
import * as crypto from 'crypto';

/**
 * Tests for the cookie import logic.
 *
 * The public importMicrosoftCookies() function depends on filesystem, Keychain,
 * and sqlite3 — so we test the pure cryptographic and conversion logic directly.
 */

describe('chrome-cookie-import logic', () => {
  describe('cookie decryption (v10 AES-128-CBC)', () => {
    // Replicate Chrome macOS encryption: PBKDF2-SHA1, salt='saltysalt', 1003 iters
    const password = 'test-password';
    const key = crypto.pbkdf2Sync(password, 'saltysalt', 1003, 16, 'sha1');
    const iv = Buffer.alloc(16, 0x20); // 16 spaces

    function encrypt(plaintext: string): string {
      const cipher = crypto.createCipheriv('aes-128-cbc', key, iv);
      const encrypted = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
      return Buffer.concat([Buffer.from('v10'), encrypted]).toString('hex');
    }

    function decrypt(hexValue: string): string | null {
      const buf = Buffer.from(hexValue, 'hex');
      if (buf.length < 4 || buf[0] !== 0x76 || buf[1] !== 0x31 || buf[2] !== 0x30) {
        return buf.toString('utf8');
      }
      try {
        const decipher = crypto.createDecipheriv('aes-128-cbc', key, iv);
        decipher.setAutoPadding(true);
        return Buffer.concat([decipher.update(buf.subarray(3)), decipher.final()]).toString('utf8');
      } catch {
        return null;
      }
    }

    it('round-trips a short cookie value', () => {
      const value = 'ESTSAUTHPERSISTENT_VALUE_HERE';
      expect(decrypt(encrypt(value))).toBe(value);
    });

    it('round-trips an empty string', () => {
      expect(decrypt(encrypt(''))).toBe('');
    });

    it('round-trips a long value (>1 AES block)', () => {
      const value = 'A'.repeat(200);
      expect(decrypt(encrypt(value))).toBe(value);
    });

    it('returns raw string for non-v10 prefix', () => {
      const plain = Buffer.from('hello');
      expect(decrypt(plain.toString('hex'))).toBe('hello');
    });

    it('returns null for corrupted ciphertext', () => {
      // Valid v10 prefix but garbage data
      const garbage = Buffer.concat([Buffer.from('v10'), Buffer.from('not-valid-ciphertext-at-all!!')]);
      expect(decrypt(garbage.toString('hex'))).toBeNull();
    });
  });

  describe('Chrome epoch conversion', () => {
    function chromeEpochToUnix(chromeTimestamp: number): number {
      return Math.floor(chromeTimestamp / 1_000_000) - 11644473600;
    }

    it('converts a known Chrome timestamp to Unix', () => {
      // 2025-01-01T00:00:00Z in Unix = 1735689600
      // In Chrome epoch = (1735689600 + 11644473600) * 1_000_000 = 13380163200000000
      const chromeTs = 13380163200000000;
      expect(chromeEpochToUnix(chromeTs)).toBe(1735689600);
    });

    it('returns 0 for the Unix epoch in Chrome time', () => {
      const chromeTs = 11644473600 * 1_000_000;
      expect(chromeEpochToUnix(chromeTs)).toBe(0);
    });
  });

  describe('samesite mapping', () => {
    function chromeSameSiteToPlaywright(samesite: number): 'Strict' | 'Lax' | 'None' {
      switch (samesite) {
        case 2: return 'Strict';
        case 1: return 'Lax';
        default: return 'None';
      }
    }

    it('maps Chrome samesite values', () => {
      expect(chromeSameSiteToPlaywright(-1)).toBe('None');
      expect(chromeSameSiteToPlaywright(0)).toBe('None');
      expect(chromeSameSiteToPlaywright(1)).toBe('Lax');
      expect(chromeSameSiteToPlaywright(2)).toBe('Strict');
    });
  });

  describe('profile detection heuristics', () => {
    const profiles = [
      { dirName: 'Default', name: 'Person 1', gaiaName: 'Test' },
      { dirName: 'Profile 1', name: 'corp.example.com', gaiaName: 'Jane Smith' },
      { dirName: 'Profile 2', name: 'Jane', gaiaName: 'Jane Smith' },
      { dirName: 'Profile 4', name: 'Test', gaiaName: '' },
    ];

    function selectWorkProfile(profiles: Array<{ dirName: string; name: string }>) {
      return profiles.find(p =>
        /\.[a-z]{2,}$/i.test(p.name) ||
        p.name.toLowerCase().includes('work') ||
        p.name.toLowerCase().includes('corp')
      ) ?? null;
    }

    it('selects profile with domain-like name', () => {
      expect(selectWorkProfile(profiles)?.dirName).toBe('Profile 1');
    });

    it('selects profile with "work" in name', () => {
      const custom = [
        { dirName: 'Default', name: 'Personal' },
        { dirName: 'Profile 1', name: 'Work Account' },
      ];
      expect(selectWorkProfile(custom)?.dirName).toBe('Profile 1');
    });

    it('selects profile with "corp" in name', () => {
      const custom = [
        { dirName: 'Default', name: 'Me' },
        { dirName: 'Profile 1', name: 'CorpNet' },
      ];
      expect(selectWorkProfile(custom)?.dirName).toBe('Profile 1');
    });

    it('returns null when no work profile found', () => {
      const custom = [
        { dirName: 'Default', name: 'Person 1' },
        { dirName: 'Profile 2', name: 'Gaming' },
      ];
      expect(selectWorkProfile(custom)).toBeNull();
    });
  });
});
