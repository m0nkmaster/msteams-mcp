/**
 * MCP Events delivery (SEP-3415, `io.modelcontextprotocol/events`).
 *
 * Two sources back the three event types:
 * - `teams.message.new` uses the incremental scan from #61
 *   (`getNewMessages`), whose opaque cursor tracks per-conversation message-ID
 *   watermarks.
 * - `teams.mention` and `teams.reply` use the Teams activity feed, which is the
 *   only source that sees replies inside channel threads. Its `syncState` is
 *   wrapped in an opaque cursor for the client.
 *
 * This service owns Teams-specific scheduling: backoff, heartbeats, filtering
 * and cursor construction.
 */

import { McpError, ErrorCode as JsonRpcErrorCode } from '@modelcontextprotocol/sdk/types.js';
import {
  decodeNewMessagesCursor,
  getNewMessages,
  type GetNewMessagesOptions,
} from '../api/chatsvc-new-messages.js';
import { getActivityFeed, type ActivityItem } from '../api/chatsvc-activity.js';
import { ErrorCode } from '../types/errors.js';
import * as log from '../utils/logger.js';
import { clampInt } from '../utils/numbers.js';
import {
  CURSOR_CLOCK_SKEW_MS,
  DEFAULT_ACTIVITY_LIMIT,
  DEFAULT_EVENTS_HEARTBEAT_MS,
  DEFAULT_EVENTS_MAX_EVENTS,
  DEFAULT_NEW_MESSAGES_MAX_CHATS,
  DEFAULT_NEW_MESSAGES_PER_CHAT,
  MAX_EVENTS_BACKOFF_DOUBLINGS,
  MAX_EVENTS_BACKOFF_MS,
  MAX_EVENTS_HEARTBEAT_MS,
  MAX_EVENTS_PER_POLL,
  MIN_EVENTS_HEARTBEAT_MS,
  NEW_MESSAGES_FLOOR_SKEW_MS,
} from '../constants.js';
import type { EventsConfig } from './config.js';
import {
  advanceScanCursor,
  createScanCursor,
  decodeActivityCursor,
  encodeActivityCursor,
  newestScanCursorMs,
  truncateScanBatch,
  type ActivityCursorState,
} from './cursor.js';
import {
  EVENT_TYPES,
  MENTION_EVENT,
  MESSAGE_NEW_EVENT,
  REPLY_EVENT,
  activityToOccurrence,
  matchesScanSubscription,
  newMessageToOccurrence,
  resolveSubscription,
  type EventOccurrence,
  type EventSubscription,
  type EventTypeDefinition,
} from './event-types.js';

/** Request id type used by JSON-RPC (the stream's subscription id). */
export type RequestId = string | number;

/** Parameters of `events/poll`. */
export interface PollParams {
  name: string;
  arguments?: unknown;
  cursor?: string | null;
  maxAgeMs?: number;
  maxEvents?: number;
}

/** Response of `events/poll`. */
export interface PollResult {
  resultType: 'complete';
  events: EventOccurrence[];
  cursor: string;
  truncated: boolean;
  hasMore: boolean;
  nextPollMs: number;
}

/** Parameters of `events/stream`. */
export interface StreamParams {
  name: string;
  arguments?: unknown;
  cursor?: string | null;
  maxAgeMs?: number;
  heartbeatMs?: number;
}

/** Final frame of `events/stream`. */
export interface StreamResult {
  resultType: 'complete';
  _meta: { 'io.modelcontextprotocol/subscriptionId': RequestId };
}

/** Sends one JSON-RPC notification for a stream. */
export type EventNotifier = (method: string, params: Record<string, unknown>) => Promise<void>;

/** Injectable dependencies, so poll/stream behaviour is testable without timers. */
export interface EventsServiceDeps {
  fetchNewMessages?: typeof getNewMessages;
  fetchActivity?: typeof getActivityFeed;
  now?: () => number;
  sleep?: (ms: number, signal: AbortSignal) => Promise<void>;
}

/** Resolves after `ms`, or immediately when the signal aborts. */
export function sleepUnlessAborted(ms: number, signal: AbortSignal): Promise<void> {
  if (signal.aborted || ms <= 0) return Promise.resolve();
  return new Promise(resolve => {
    const done = (): void => {
      clearTimeout(timer);
      signal.removeEventListener('abort', done);
      resolve();
    };
    const timer = setTimeout(done, ms);
    signal.addEventListener('abort', done, { once: true });
  });
}

/** Clamps a client's heartbeat suggestion into the server's bounds. */
export function clampHeartbeatMs(value: number | undefined): number {
  if (value !== undefined && (!Number.isFinite(value) || value <= 0)) {
    return DEFAULT_EVENTS_HEARTBEAT_MS;
  }
  return clampInt(value, DEFAULT_EVENTS_HEARTBEAT_MS, MIN_EVENTS_HEARTBEAT_MS, MAX_EVENTS_HEARTBEAT_MS);
}

/** Returns true for auth failures that no amount of backoff will fix. */
function isAuthError(code: ErrorCode): boolean {
  return code === ErrorCode.AUTH_REQUIRED || code === ErrorCode.AUTH_EXPIRED;
}

/** Teams event source shared by poll requests and push streams. */
export class EventsService {
  private failures = 0;
  private retryAfterMs = 0;

  private readonly fetchNewMessages: typeof getNewMessages;
  private readonly fetchActivity: typeof getActivityFeed;
  private readonly now: () => number;
  private readonly sleep: (ms: number, signal: AbortSignal) => Promise<void>;

  constructor(
    private readonly config: EventsConfig,
    deps: EventsServiceDeps = {}
  ) {
    this.fetchNewMessages = deps.fetchNewMessages ?? getNewMessages;
    this.fetchActivity = deps.fetchActivity ?? getActivityFeed;
    this.now = deps.now ?? (() => Date.now());
    this.sleep = deps.sleep ?? sleepUnlessAborted;
  }

  /** `events/list` response. */
  list(): { resultType: 'complete'; events: readonly EventTypeDefinition[] } {
    return { resultType: 'complete', events: EVENT_TYPES };
  }

  /** Recommended delay before the next `events/poll`, including backoff. */
  nextPollMs(): number {
    const backoff = this.config.intervalMs * 2 ** this.failures;
    return Math.min(Math.max(backoff, this.retryAfterMs), MAX_EVENTS_BACKOFF_MS);
  }

  /** Resets the error backoff after a successful upstream poll. */
  recordSuccess(): void {
    this.failures = 0;
    this.retryAfterMs = 0;
  }

  /** Lengthens the next poll delay after an upstream failure or throttle. */
  recordFailure(retryAfterMs?: number): void {
    this.failures = Math.min(this.failures + 1, MAX_EVENTS_BACKOFF_DOUBLINGS);
    if (retryAfterMs !== undefined && retryAfterMs > 0) {
      this.retryAfterMs = Math.max(this.retryAfterMs, retryAfterMs);
    }
  }

  /** Handles one `events/poll` request. */
  async poll(params: PollParams): Promise<PollResult> {
    const subscription = resolveSubscription(params.name, params.arguments);
    return subscription.name === MESSAGE_NEW_EVENT
      ? this.pollNewMessages(subscription, params)
      : this.pollActivity(subscription, params);
  }

  /**
   * Serves one `events/stream` until the client cancels it.
   *
   * The stream polls Teams itself, so events arrive without the client asking;
   * heartbeats advance the client's cursor during quiet periods.
   */
  async stream(
    params: StreamParams,
    requestId: RequestId,
    notify: EventNotifier,
    signal: AbortSignal
  ): Promise<StreamResult> {
    const subscription = resolveSubscription(params.name, params.arguments);
    const meta = { 'io.modelcontextprotocol/subscriptionId': requestId };

    if (subscription.name === MESSAGE_NEW_EVENT) {
      await this.streamNewMessages(subscription, params, notify, signal, meta);
    } else {
      await this.streamActivity(subscription, params, notify, signal, meta);
    }

    return { resultType: 'complete', _meta: meta };
  }

  // ───────────────────────────────────────────────────────────────────────────
  // teams.message.new
  // ───────────────────────────────────────────────────────────────────────────

  /** Handles one poll against the incremental new-message scan. */
  private async pollNewMessages(
    subscription: EventSubscription,
    params: PollParams
  ): Promise<PollResult> {
    const prepared = this.prepareScanCursor(params.cursor, params.maxAgeMs);

    if (prepared.bootstrapped) {
      return {
        resultType: 'complete',
        events: [],
        cursor: prepared.cursor,
        truncated: prepared.truncated,
        hasMore: false,
        nextPollMs: this.nextPollMs(),
      };
    }

    const result = await this.fetchNewMessages(this.scanOptions(subscription, prepared.cursor));

    if (!result.ok) {
      this.recordFailure(result.error.retryAfterMs);
      log.warn('events', `Upstream poll failed (${result.error.code}): ${result.error.message}`);
      if (isAuthError(result.error.code)) {
        throw new McpError(
          JsonRpcErrorCode.InternalError,
          `Teams authentication failed while polling events: ${result.error.message}`
        );
      }
      return {
        resultType: 'complete',
        events: [],
        cursor: prepared.cursor,
        truncated: false,
        hasMore: false,
        nextPollMs: this.nextPollMs(),
      };
    }

    this.recordSuccess();
    this.warnAboutScanErrors(result.value.errors);

    const matching = result.value.items.filter(item => matchesScanSubscription(subscription, item));
    const maxEvents = clampInt(
      params.maxEvents,
      DEFAULT_EVENTS_MAX_EVENTS,
      1,
      MAX_EVENTS_PER_POLL
    );
    const returned = matching.slice(0, maxEvents);
    const withheld = matching.length > returned.length;

    return {
      resultType: 'complete',
      events: returned.map(item => newMessageToOccurrence(item)),
      // A cut batch resumes after the last returned message. Otherwise the
      // scan's cursor already points past everything it covered.
      cursor: withheld
        ? truncateScanBatch(prepared.cursor, returned)
        : result.value.nextAfter,
      // Per-conversation failures mean events from those conversations were
      // skipped; a full page does not (its cursor resumes there).
      truncated: result.value.errors.length > 0,
      hasMore: withheld || result.value.truncated,
      nextPollMs: this.nextPollMs(),
    };
  }

  /** Pushes new-message scan results until cancelled. */
  private async streamNewMessages(
    subscription: EventSubscription,
    params: StreamParams,
    notify: EventNotifier,
    signal: AbortSignal,
    meta: Record<string, string | RequestId>
  ): Promise<void> {
    const heartbeatMs = clampHeartbeatMs(params.heartbeatMs);
    const prepared = this.prepareScanCursor(params.cursor, params.maxAgeMs);
    let cursor = prepared.cursor;

    await notify('notifications/events/active', {
      cursor,
      truncated: prepared.truncated,
      heartbeatMs,
      _meta: meta,
    });
    let lastActivity = this.now();

    while (!signal.aborted) {
      const result = await this.fetchNewMessages(this.scanOptions(subscription, cursor));

      if (!result.ok) {
        this.recordFailure(result.error.retryAfterMs);
        log.warn('events', `Stream poll failed (${result.error.code}): ${result.error.message}`);
        await notify('notifications/events/error', {
          error: { code: JsonRpcErrorCode.InternalError, message: result.error.message },
          _meta: meta,
        });
        await this.sleep(this.nextPollMs(), signal);
        continue;
      }

      this.recordSuccess();
      if (result.value.errors.length > 0) {
        await this.notifyScanErrors(notify, result.value.errors, meta);
      }

      for (const item of result.value.items) {
        if (signal.aborted) break;
        if (!matchesScanSubscription(subscription, item)) continue;

        cursor = advanceScanCursor(cursor, item.conversationId, item.messageId);
        await notify('notifications/events/event', {
          ...newMessageToOccurrence(item, cursor),
          _meta: meta,
        });
        lastActivity = this.now();
      }

      cursor = result.value.nextAfter;
      if (signal.aborted) break;

      if (this.now() - lastActivity >= heartbeatMs) {
        await this.sendHeartbeat(notify, cursor, meta);
        lastActivity = this.now();
      }

      await this.waitForNextPoll(this.nextPollMs(), heartbeatMs, async () => {
        await this.sendHeartbeat(notify, cursor, meta);
        lastActivity = this.now();
      }, signal);
    }
  }

  /**
   * Validates an incoming scan cursor. A malformed cursor is `InvalidParams`; a
   * cursor older than `maxAgeMs` (or ahead of the clock) restarts from now and
   * reports `truncated`, because the events in between were skipped.
   */
  private prepareScanCursor(
    cursor: string | null | undefined,
    maxAgeMs?: number
  ): { cursor: string; truncated: boolean; bootstrapped: boolean } {
    // Keep the scan's floor skew so a message arriving just after bootstrap
    // with a slightly-behind server clock is not skipped.
    const fresh = (): string => createScanCursor(this.now() - NEW_MESSAGES_FLOOR_SKEW_MS);

    if (!cursor) {
      return { cursor: fresh(), truncated: false, bootstrapped: true };
    }

    const decoded = decodeNewMessagesCursor(cursor);
    if (!decoded) {
      throw new McpError(
        JsonRpcErrorCode.InvalidParams,
        "'cursor' is not a cursor this server issued. Omit it to start from now."
      );
    }

    const now = this.now();
    const newestMs = newestScanCursorMs(decoded);
    const tooOld = maxAgeMs !== undefined && maxAgeMs > 0 && now - newestMs > maxAgeMs;
    const inTheFuture = decoded.since > now + CURSOR_CLOCK_SKEW_MS;

    if (tooOld || inTheFuture) {
      return { cursor: fresh(), truncated: true, bootstrapped: true };
    }
    return { cursor, truncated: false, bootstrapped: false };
  }

  /** Scan options shared by poll requests and push streams. */
  private scanOptions(
    subscription: EventSubscription,
    after: string
  ): GetNewMessagesOptions {
    return {
      after,
      maxChats: DEFAULT_NEW_MESSAGES_MAX_CHATS,
      limitPerChat: DEFAULT_NEW_MESSAGES_PER_CHAT,
      includeOwn: subscription.includeOwn,
      now: this.now(),
    };
  }

  // ───────────────────────────────────────────────────────────────────────────
  // teams.mention / teams.reply (activity feed)
  // ───────────────────────────────────────────────────────────────────────────

  /** Handles one poll against the Teams activity feed. */
  private async pollActivity(
    subscription: EventSubscription,
    params: PollParams
  ): Promise<PollResult> {
    if (!params.cursor) {
      const state = await this.startActivityCursor();
      return {
        resultType: 'complete',
        events: [],
        cursor: encodeActivityCursor(state),
        truncated: false,
        hasMore: false,
        nextPollMs: this.nextPollMs(),
      };
    }

    let state = decodeActivityCursor(params.cursor);
    if (!state) {
      throw new McpError(
        JsonRpcErrorCode.InvalidParams,
        "'cursor' is not a cursor this server issued. Omit it to start from now."
      );
    }

    if (params.maxAgeMs !== undefined && params.maxAgeMs > 0 && this.now() - state.at > params.maxAgeMs) {
      state = await this.startActivityCursor();
      return {
        resultType: 'complete',
        events: [],
        cursor: encodeActivityCursor(state),
        truncated: true,
        hasMore: false,
        nextPollMs: this.nextPollMs(),
      };
    }

    const before = encodeActivityCursor(state);
    const result = await this.fetchActivity({
      limit: DEFAULT_ACTIVITY_LIMIT,
      syncState: state.syncState,
    });

    if (!result.ok) {
      this.recordFailure(result.error.retryAfterMs);
      log.warn('events', `Activity poll failed (${result.error.code}): ${result.error.message}`);
      if (isAuthError(result.error.code)) {
        throw new McpError(
          JsonRpcErrorCode.InternalError,
          `Teams authentication failed while polling events: ${result.error.message}`
        );
      }
      return {
        resultType: 'complete',
        events: [],
        cursor: before,
        truncated: false,
        hasMore: false,
        nextPollMs: this.nextPollMs(),
      };
    }

    this.recordSuccess();

    const activities = result.value.activities;
    const next = this.nextActivityState(state, result.value.syncState);

    return {
      resultType: 'complete',
      // Per-event cursors stay at the batch start: reconnecting mid-batch can
      // replay events, which the client deduplicates by eventId, but never
      // skips them.
      events: this.mapActivityEvents(subscription.name, activities, before),
      cursor: encodeActivityCursor(next),
      // The feed is a watermark, not a page cursor: a full page may have left
      // older items behind.
      truncated: activities.length >= DEFAULT_ACTIVITY_LIMIT,
      hasMore: false,
      nextPollMs: this.nextPollMs(),
    };
  }

  /** Pushes activity-feed events until cancelled. */
  private async streamActivity(
    subscription: EventSubscription,
    params: StreamParams,
    notify: EventNotifier,
    signal: AbortSignal,
    meta: Record<string, string | RequestId>
  ): Promise<void> {
    const heartbeatMs = clampHeartbeatMs(params.heartbeatMs);
    let state: ActivityCursorState | null = null;
    let truncated = false;

    if (params.cursor) {
      state = decodeActivityCursor(params.cursor);
      if (!state) {
        throw new McpError(
          JsonRpcErrorCode.InvalidParams,
          "'cursor' is not a cursor this server issued. Omit it to start from now."
        );
      }
      if (params.maxAgeMs !== undefined && params.maxAgeMs > 0 && this.now() - state.at > params.maxAgeMs) {
        state = null;
        truncated = true;
      }
    }

    if (!state) {
      state = await this.tryActivityState();
    }

    await notify('notifications/events/active', {
      cursor: state ? encodeActivityCursor(state) : null,
      truncated,
      heartbeatMs,
      _meta: meta,
    });
    let lastActivity = this.now();

    while (!signal.aborted) {
      if (!state) {
        const bootstrapped = await this.tryActivityState();
        if (!bootstrapped) {
          await this.sleep(this.nextPollMs(), signal);
          continue;
        }
        state = bootstrapped;
        await notify('notifications/events/active', {
          cursor: encodeActivityCursor(state),
          truncated: false,
          heartbeatMs,
          _meta: meta,
        });
        lastActivity = this.now();
      }

      const before = encodeActivityCursor(state);
      const result = await this.fetchActivity({
        limit: DEFAULT_ACTIVITY_LIMIT,
        syncState: state.syncState,
      });

      if (!result.ok) {
        this.recordFailure(result.error.retryAfterMs);
        log.warn('events', `Activity stream failed (${result.error.code}): ${result.error.message}`);
        await notify('notifications/events/error', {
          error: { code: JsonRpcErrorCode.InternalError, message: result.error.message },
          _meta: meta,
        });
        await this.sleep(this.nextPollMs(), signal);
        continue;
      }

      this.recordSuccess();

      for (const event of this.mapActivityEvents(subscription.name, result.value.activities, before)) {
        if (signal.aborted) break;
        await notify('notifications/events/event', { ...event, _meta: meta });
        lastActivity = this.now();
      }
      if (signal.aborted) break;

      const fullPage = result.value.activities.length >= DEFAULT_ACTIVITY_LIMIT;
      state = this.nextActivityState(state, result.value.syncState);

      if (fullPage) {
        // Older activity may have been skipped; tell the client where we are.
        await notify('notifications/events/active', {
          cursor: encodeActivityCursor(state),
          truncated: true,
          heartbeatMs,
          _meta: meta,
        });
        lastActivity = this.now();
      }

      if (this.now() - lastActivity >= heartbeatMs) {
        await this.sendHeartbeat(notify, encodeActivityCursor(state), meta);
        lastActivity = this.now();
      }

      await this.waitForNextPoll(this.nextPollMs(), heartbeatMs, async () => {
        await this.sendHeartbeat(notify, encodeActivityCursor(state as ActivityCursorState), meta);
        lastActivity = this.now();
      }, signal);
    }
  }

  /** Whether the activity feed returned an opaque sync state. */
  private mapActivityEvents(
    name: EventSubscription['name'],
    activities: readonly ActivityItem[],
    cursor: string
  ): EventOccurrence[] {
    const wanted = name === MENTION_EVENT ? 'mention' : 'reply';
    return activities
      .filter(item => item.type === wanted)
      .map(item => activityToOccurrence(name as typeof MENTION_EVENT | typeof REPLY_EVENT, item, cursor))
      .filter((event): event is EventOccurrence => event !== null);
  }

  /** Builds the next activity cursor, keeping the old syncState when absent. */
  private nextActivityState(
    previous: ActivityCursorState,
    syncState: string | undefined
  ): ActivityCursorState {
    if (!syncState) {
      log.warn('events', 'Teams activity feed returned no syncState; keeping the previous cursor.');
    }
    return {
      version: 1,
      syncState: syncState ?? previous.syncState,
      at: this.now(),
    };
  }

  /** Starts an activity subscription at "now" without emitting old items. */
  private async startActivityCursor(): Promise<ActivityCursorState> {
    const result = await this.fetchActivity({ limit: DEFAULT_ACTIVITY_LIMIT });

    if (!result.ok) {
      this.recordFailure(result.error.retryAfterMs);
      log.warn('events', `Activity bootstrap failed (${result.error.code}): ${result.error.message}`);
      throw new McpError(
        JsonRpcErrorCode.InternalError,
        `Could not start the Teams activity feed: ${result.error.message}`
      );
    }
    if (!result.value.syncState) {
      throw new McpError(
        JsonRpcErrorCode.InternalError,
        'Teams did not return an activity sync state, so mentions and replies cannot be resumed.'
      );
    }

    this.recordSuccess();
    return { version: 1, syncState: result.value.syncState, at: this.now() };
  }

  /** Best-effort bootstrap for a stream, which retries instead of failing. */
  private async tryActivityState(): Promise<ActivityCursorState | null> {
    try {
      return await this.startActivityCursor();
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log.warn('events', `Activity bootstrap failed: ${message}`);
      return null;
    }
  }

  // ───────────────────────────────────────────────────────────────────────────
  // Shared delivery helpers
  // ───────────────────────────────────────────────────────────────────────────

  /** Logs conversations that could not be scanned in a poll. */
  private warnAboutScanErrors(errors: readonly { conversationId: string }[]): void {
    if (errors.length === 0) return;
    log.warn(
      'events',
      `${errors.length} conversation(s) could not be scanned: ` +
      errors.map(error => error.conversationId).join(', ')
    );
  }

  /** Reports conversations that could not be scanned on a stream. */
  private async notifyScanErrors(
    notify: EventNotifier,
    errors: readonly { conversationId: string; code: string; message: string }[],
    meta: Record<string, string | RequestId>
  ): Promise<void> {
    this.warnAboutScanErrors(errors);
    await notify('notifications/events/error', {
      error: {
        code: JsonRpcErrorCode.InternalError,
        message: `${errors.length} conversation(s) could not be scanned`,
        data: { conversations: errors },
      },
      _meta: meta,
    });
  }

  /** Sends a heartbeat carrying the stream's latest cursor. */
  private async sendHeartbeat(
    notify: EventNotifier,
    cursor: string,
    meta: Record<string, string | RequestId>
  ): Promise<void> {
    await notify('notifications/events/heartbeat', { cursor, _meta: meta });
  }

  /**
   * Sleeps until the next poll, emitting a heartbeat at least every
   * `heartbeatMs` so the client can tell "quiet" from "dead".
   */
  private async waitForNextPoll(
    totalMs: number,
    heartbeatMs: number,
    sendHeartbeat: () => Promise<void>,
    signal: AbortSignal
  ): Promise<void> {
    let remaining = totalMs;
    while (remaining > 0 && !signal.aborted) {
      const slice = Math.min(remaining, heartbeatMs);
      await this.sleep(slice, signal);
      remaining -= slice;
      if (!signal.aborted && remaining > 0) {
        await sendHeartbeat();
      }
    }
  }
}
