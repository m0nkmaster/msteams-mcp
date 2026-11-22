/**
 * Opaque cursors for MCP Events.
 *
 * `teams.message.new` reuses the incremental scan cursor from
 * `api/chatsvc-new-messages.ts`. `teams.mention` and `teams.reply` are served
 * by the activity feed, whose `syncState` is already an opaque watermark, so
 * it is wrapped with the time it was issued to support `maxAgeMs`.
 */

import {
  decodeNewMessagesCursor,
  encodeNewMessagesCursor,
  type NewMessagesCursor,
} from '../api/chatsvc-new-messages.js';

/** Decoded activity-feed cursor. */
export interface ActivityCursorState {
  version: 1;
  /** Opaque `syncState` from the activity feed. */
  syncState: string;
  /** When the cursor was issued (epoch ms), for `maxAgeMs`. */
  at: number;
}

/** A message a cursor batch actually returned. */
export interface CursorEventRef {
  conversationId: string;
  messageId: string;
}

/** Encodes an activity cursor as opaque base64url. */
export function encodeActivityCursor(state: ActivityCursorState): string {
  return Buffer.from(
    JSON.stringify({ v: state.version, s: state.syncState, t: state.at }),
    'utf8'
  ).toString('base64url');
}

/** Decodes an activity cursor, returning null for anything else. */
export function decodeActivityCursor(value: string): ActivityCursorState | null {
  if (!value) return null;

  let parsed: unknown;
  try {
    parsed = JSON.parse(Buffer.from(value, 'base64url').toString('utf8'));
  } catch {
    return null;
  }

  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return null;
  const raw = parsed as Record<string, unknown>;
  if (raw.v !== 1) return null;
  if (typeof raw.s !== 'string' || raw.s.length === 0) return null;
  if (typeof raw.t !== 'number' || !Number.isFinite(raw.t) || raw.t < 0) return null;

  return { version: 1, syncState: raw.s, at: raw.t };
}

/** A cursor that starts the scan at `sinceMs` ("start from now"). */
export function createScanCursor(sinceMs: number): string {
  return encodeNewMessagesCursor({ version: 1, since: sinceMs, chats: {} });
}

/**
 * The newest point a scan cursor has reached, in epoch ms. The poll floor is
 * only where unscanned conversations start; per-chat watermarks are message
 * IDs, which Teams also uses as epoch-ms arrival times.
 */
export function newestScanCursorMs(cursor: NewMessagesCursor): number {
  let newest = cursor.since;
  for (const messageId of Object.values(cursor.chats)) {
    const parsed = Number(messageId);
    if (Number.isFinite(parsed) && parsed > newest) newest = parsed;
  }
  return newest;
}

/**
 * Advances one conversation's watermark to an event's message ID.
 *
 * Push streams attach a cursor to every event. Only the event's own
 * conversation moves, so reconnecting from a mid-batch cursor may replay other
 * conversations (deduplicated by `eventId`) but never skips them.
 */
export function advanceScanCursor(
  cursor: string,
  conversationId: string,
  messageId: string
): string {
  const state = decodeNewMessagesCursor(cursor);
  const next = Number(messageId);
  if (!state || !Number.isFinite(next) || next <= 0) return cursor;

  const current = Number(state.chats[conversationId] ?? 0);
  if (!Number.isFinite(current) || next > current) {
    state.chats[conversationId] = String(next);
  }
  return encodeNewMessagesCursor(state);
}

/**
 * Builds the cursor for a batch that was cut short by `maxEvents`.
 *
 * Only conversations whose events were returned advance; withheld events stay
 * after their conversation's watermark and arrive on the next poll. The poll
 * floor is left alone so conversations not in the batch are not skipped.
 */
export function truncateScanBatch(
  cursor: string,
  returned: readonly CursorEventRef[]
): string {
  const state = decodeNewMessagesCursor(cursor);
  if (!state) return cursor;

  for (const event of returned) {
    const next = Number(event.messageId);
    if (!Number.isFinite(next) || next <= 0) continue;
    const current = Number(state.chats[event.conversationId] ?? 0);
    if (!Number.isFinite(current) || next > current) {
      state.chats[event.conversationId] = String(next);
    }
  }
  return encodeNewMessagesCursor(state);
}
