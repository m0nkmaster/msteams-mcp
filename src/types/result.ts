/**
 * Result type for operations that can fail.
 * 
 * Provides a discriminated union for success/failure that
 * forces callers to handle errors explicitly.
 */

import type { McpError } from './errors.js';

/** Successful result with a value. */
export interface Ok<T> {
  ok: true;
  value: T;
}

/** Failed result with an error. */
export interface Err<E = McpError> {
  ok: false;
  error: E;
}

/** A result that can be either a success or failure. */
export type Result<T, E = McpError> = Ok<T> | Err<E>;

/**
 * Creates a successful result.
 */
export function ok<T>(value: T): Ok<T> {
  return { ok: true, value };
}

/**
 * Creates a failed result.
 */
export function err<E = McpError>(error: E): Err<E> {
  return { ok: false, error };
}
