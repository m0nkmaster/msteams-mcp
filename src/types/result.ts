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

/**
 * Unwraps a result, throwing if it's an error.
 */
export function unwrap<T>(result: Result<T>): T {
  if (result.ok) {
    return result.value;
  }
  throw new Error(result.error.message);
}

/**
 * Unwraps a result with a default value for errors.
 */
export function unwrapOr<T>(result: Result<T>, defaultValue: T): T {
  return result.ok ? result.value : defaultValue;
}

/**
 * Maps a successful result to a new value.
 */
export function map<T, U>(result: Result<T>, fn: (value: T) => U): Result<U> {
  if (result.ok) {
    return ok(fn(result.value));
  }
  return result;
}

/**
 * Maps a successful result to a new result (flatMap).
 */
export function andThen<T, U>(
  result: Result<T>,
  fn: (value: T) => Result<U>
): Result<U> {
  if (result.ok) {
    return fn(result.value);
  }
  return result;
}
