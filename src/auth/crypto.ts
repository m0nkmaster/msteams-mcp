/**
 * Encryption utilities for credential storage.
 * 
 * Uses machine-specific key derivation to encrypt sensitive data at rest.
 * This is not foolproof security, but significantly raises the bar compared
 * to plaintext storage.
 */

import * as crypto from 'crypto';
import * as os from 'os';

/** Algorithm for encryption. */
const ALGORITHM = 'aes-256-gcm';

/** Salt for key derivation. */
const SALT = 'teams-mcp-credential-salt-v1';

/** Derives an encryption key from machine-specific values. */
function deriveKey(): Buffer {
  // Combine hostname and username for machine-specific key
  const machineId = `${os.hostname()}:${os.userInfo().username}`;
  
  return crypto.scryptSync(machineId, SALT, 32);
}

/** Encrypted data format. */
export interface EncryptedData {
  /** Initialisation vector (hex). */
  iv: string;
  /** Encrypted content (hex). */
  content: string;
  /** Authentication tag (hex). */
  tag: string;
  /** Version marker for format compatibility. */
  version: number;
}

/**
 * Encrypts a string value.
 */
export function encrypt(plaintext: string): EncryptedData {
  const key = deriveKey();
  const iv = crypto.randomBytes(16);
  
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);
  
  let encrypted = cipher.update(plaintext, 'utf8', 'hex');
  encrypted += cipher.final('hex');
  
  const tag = cipher.getAuthTag();
  
  return {
    iv: iv.toString('hex'),
    content: encrypted,
    tag: tag.toString('hex'),
    version: 1,
  };
}

/**
 * Decrypts encrypted data.
 * 
 * @throws Error if decryption fails (wrong machine, corrupted data, etc.)
 */
export function decrypt(data: EncryptedData): string {
  if (data.version !== 1) {
    throw new Error(`Unsupported encryption version: ${data.version}`);
  }
  
  const key = deriveKey();
  const iv = Buffer.from(data.iv, 'hex');
  const tag = Buffer.from(data.tag, 'hex');
  
  const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
  decipher.setAuthTag(tag);
  
  let decrypted = decipher.update(data.content, 'hex', 'utf8');
  decrypted += decipher.final('utf8');
  
  return decrypted;
}

/**
 * Checks if data looks like encrypted format.
 */
export function isEncrypted(data: unknown): data is EncryptedData {
  if (!data || typeof data !== 'object') return false;
  
  const obj = data as Record<string, unknown>;
  return (
    typeof obj.iv === 'string' &&
    typeof obj.content === 'string' &&
    typeof obj.tag === 'string' &&
    typeof obj.version === 'number'
  );
}
