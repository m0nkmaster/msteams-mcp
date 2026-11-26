/**
 * Unit tests for the opaque event cursors.
 */

import { describe, it, expect } from 'vitest';
import { decodeNewMessagesCursor, encodeNewMessagesCursor } from '../api/chatsvc-new-messages.js';
import {
  advanceScanCursor,
  createScanCursor,
  decodeActivityCursor,
  encodeActivityCursor,
  newestScanCursorMs,
  truncateScanBatch,
} from './cursor.js';

const NOW = Date.parse('2026-01-01T12:00:00.000Z');

describe('activity cursor', () => {
  it('round-trips the sync state and issue time', () => {
    const state = { version: 1 as const, syncState: 'token/abc==', at: NOW };
    const encoded = encodeActivityCursor(state);
    expect(encoded).toMatch(/^[A-Za-z0-9_-]+$/);
    expect(decodeActivityCursor(encoded)).toEqual(state);
  });

  it('rejects anything that is not an activity cursor', () => {
    expect(decodeActivityCursor('')).toBeNull();
    expect(decodeActivityCursor('not-json')).toBeNull();
    expect(decodeActivityCursor(Buffer.from('{"v":2,"s":"x","t":1}').toString('base64url'))).toBeNull();
    expect(decodeActivityCursor(Buffer.from('{"v":1,"s":"","t":1}').toString('base64url'))).toBeNull();
    expect(decodeActivityCursor(Buffer.from('{"v":1,"s":"x","t":"nope"}').toString('base64url'))).toBeNull();
  });
});

describe('scan cursor helpers', () => {
  it('creates a cursor that starts at a given instant', () => {
    const decoded = decodeNewMessagesCursor(createScanCursor(NOW));
    expect(decoded).toEqual({ version: 1, since: NOW, chats: {} });
  });

  it('reports the newest point reached, floor or per-chat watermark', () => {
    expect(newestScanCursorMs({ version: 1, since: NOW, chats: {} })).toBe(NOW);
    expect(newestScanCursorMs({
      version: 1,
      since: NOW - 60_000,
      chats: { a: String(NOW + 1000), b: 'not-a-number' },
    })).toBe(NOW + 1000);
  });

  it('advances only the event conversation watermark', () => {
    const base = encodeNewMessagesCursor({
      version: 1,
      since: NOW - 60_000,
      chats: { other: String(NOW - 30_000) },
    });

    const next = decodeNewMessagesCursor(advanceScanCursor(base, 'chat', String(NOW)));
    expect(next?.chats.chat).toBe(String(NOW));
    expect(next?.chats.other).toBe(String(NOW - 30_000));
    expect(next?.since).toBe(NOW - 60_000);
  });

  it('never moves a watermark backwards and ignores invalid input', () => {
    const base = encodeNewMessagesCursor({
      version: 1,
      since: NOW - 60_000,
      chats: { chat: String(NOW) },
    });

    const older = decodeNewMessagesCursor(advanceScanCursor(base, 'chat', String(NOW - 1000)));
    expect(older?.chats.chat).toBe(String(NOW));
    expect(advanceScanCursor('nope', 'chat', String(NOW))).toBe('nope');
    expect(advanceScanCursor(base, 'chat', 'not-a-message-id')).toBe(base);
  });

  it('builds a partial-batch cursor from the messages actually returned', () => {
    const base = encodeNewMessagesCursor({
      version: 1,
      since: NOW - 60_000,
      chats: { a: String(NOW - 30_000) },
    });

    const next = decodeNewMessagesCursor(truncateScanBatch(base, [
      { conversationId: 'a', messageId: String(NOW - 20_000) },
      { conversationId: 'b', messageId: String(NOW - 10_000) },
    ]));

    expect(next?.since).toBe(NOW - 60_000);
    expect(next?.chats.a).toBe(String(NOW - 20_000));
    expect(next?.chats.b).toBe(String(NOW - 10_000));
  });

  it('returns the input unchanged when it is not a scan cursor', () => {
    expect(truncateScanBatch('nope', [{ conversationId: 'a', messageId: '1' }])).toBe('nope');
  });
});
