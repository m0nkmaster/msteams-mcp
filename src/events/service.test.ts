/**
 * Unit tests for MCP Events polling, filtering, backoff and push delivery.
 */

import { describe, it, expect, vi } from 'vitest';
import { McpError } from '@modelcontextprotocol/sdk/types.js';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import {
  decodeNewMessagesCursor,
  encodeNewMessagesCursor,
  type GetNewMessagesOptions,
  type GetNewMessagesResult,
  type NewMessageItem,
} from '../api/chatsvc-new-messages.js';
import type { ActivityItem, GetActivityResult } from '../api/chatsvc-activity.js';
import {
  DEFAULT_ACTIVITY_LIMIT,
  DEFAULT_EVENTS_INTERVAL_MS,
  MAX_EVENTS_BACKOFF_MS,
  MIN_EVENTS_INTERVAL_MS,
} from '../constants.js';
import type { EventsConfig } from './config.js';
import {
  EVENT_TYPES,
  MENTION_EVENT,
  MESSAGE_NEW_EVENT,
  REPLY_EVENT,
} from './event-types.js';
import { advanceScanCursor, decodeActivityCursor, encodeActivityCursor } from './cursor.js';
import { EventsService, clampHeartbeatMs, sleepUnlessAborted } from './service.js';

const NOW = Date.parse('2026-01-01T12:00:00.000Z');
const SCAN_CURSOR = encodeNewMessagesCursor({
  version: 1,
  since: NOW - 60_000,
  chats: { '19:other@thread.v2': String(NOW - 30_000) },
});
const NEXT_CURSOR = encodeNewMessagesCursor({
  version: 1,
  since: NOW,
  chats: { '19:chat@thread.v2': '1770000000000' },
});
const ACTIVITY_SYNC = 'activity-sync-token';
const ACTIVITY_CURSOR = encodeActivityCursor({
  version: 1,
  syncState: ACTIVITY_SYNC,
  at: NOW - 1000,
});

/** Events config used by most tests. */
function config(overrides: Partial<EventsConfig> = {}): EventsConfig {
  return { enabled: true, intervalMs: DEFAULT_EVENTS_INTERVAL_MS, ...overrides };
}

/** A scan item with sensible defaults. */
function item(overrides: Partial<NewMessageItem> = {}): NewMessageItem {
  return {
    conversationId: '19:chat@thread.v2',
    topic: 'Chat',
    chatType: 'group',
    messageId: '100',
    senderName: 'Bob',
    sentAt: '2026-01-01T11:58:00.000Z',
    snippet: 'hello',
    ...overrides,
  };
}

/** A successful scan result. */
function scan(
  items: NewMessageItem[] = [],
  overrides: Partial<GetNewMessagesResult> = {}
): GetNewMessagesResult {
  return { items, nextAfter: NEXT_CURSOR, truncated: false, errors: [], ...overrides };
}

/** An activity item with sensible defaults. */
function activity(overrides: Partial<ActivityItem> = {}): ActivityItem {
  return {
    id: '200',
    type: 'mention',
    content: 'ping',
    contentType: 'RichText/Html',
    sender: { mri: '8:orgid:someone', displayName: 'Alice' },
    timestamp: '2026-01-01T11:59:00.000Z',
    conversationId: '19:c@thread.v2',
    ...overrides,
  };
}

/** A successful activity page. */
function activityPage(
  activities: ActivityItem[] = [],
  syncState: string | undefined = 'next-sync-token'
): GetActivityResult {
  return { activities, syncState };
}

const noSleep = async (): Promise<void> => {};

describe('EventsService.list', () => {
  it('publishes the Teams event types', () => {
    const service = new EventsService(config());
    expect(service.list()).toEqual({ resultType: 'complete', events: EVENT_TYPES });
  });
});

describe('backoff', () => {
  it('starts at the configured interval and doubles on failures', () => {
    const service = new EventsService(config({ intervalMs: 20_000 }), { now: () => NOW });
    expect(service.nextPollMs()).toBe(20_000);
    service.recordFailure();
    expect(service.nextPollMs()).toBe(40_000);
    service.recordFailure();
    expect(service.nextPollMs()).toBe(80_000);
    service.recordSuccess();
    expect(service.nextPollMs()).toBe(20_000);
  });

  it('honours a retry-after hint and caps the delay', () => {
    const service = new EventsService(config(), { now: () => NOW });
    service.recordFailure(120_000);
    expect(service.nextPollMs()).toBe(120_000);

    for (let i = 0; i < 20; i++) service.recordFailure();
    expect(service.nextPollMs()).toBe(MAX_EVENTS_BACKOFF_MS);
  });

  it('never polls below the configured interval', () => {
    const service = new EventsService(config({ intervalMs: MIN_EVENTS_INTERVAL_MS }), { now: () => NOW });
    expect(service.nextPollMs()).toBe(MIN_EVENTS_INTERVAL_MS);
  });
});

describe('poll: teams.message.new', () => {
  it('bootstraps from now without touching Teams', async () => {
    const fetchNewMessages = vi.fn();
    const service = new EventsService(config(), { fetchNewMessages, now: () => NOW });

    const result = await service.poll({ name: MESSAGE_NEW_EVENT, cursor: null });

    expect(fetchNewMessages).not.toHaveBeenCalled();
    expect(result.events).toEqual([]);
    expect(result.truncated).toBe(false);
    expect(result.nextPollMs).toBe(DEFAULT_EVENTS_INTERVAL_MS);
    expect(decodeNewMessagesCursor(result.cursor)?.since).toBe(NOW - 60_000);
  });

  it('fetches from the supplied cursor with the subscription options', async () => {
    const fetchNewMessages = vi.fn(async () => ok(scan([item()])));
    const service = new EventsService(config(), { fetchNewMessages, now: () => NOW });

    const result = await service.poll({ name: MESSAGE_NEW_EVENT, cursor: SCAN_CURSOR });

    expect(fetchNewMessages).toHaveBeenCalledWith({
      after: SCAN_CURSOR,
      maxChats: 30,
      limitPerChat: 50,
      includeOwn: false,
      now: NOW,
    });
    expect(result.events).toHaveLength(1);
    expect(result.events[0]).toMatchObject({
      eventId: '19:chat@thread.v2:100',
      name: MESSAGE_NEW_EVENT,
      data: { conversationId: '19:chat@thread.v2', messageId: '100', snippet: 'hello' },
    });
    expect(result.cursor).toBe(NEXT_CURSOR);
    expect(result.hasMore).toBe(false);
  });

  it('passes includeOwn through and narrows by conversation', async () => {
    const fetchNewMessages = vi.fn(async (_options?: GetNewMessagesOptions) => ok(scan([
      item({ messageId: '1', conversationId: '19:a@thread.v2' }),
      item({ messageId: '2', conversationId: '19:b@thread.v2' }),
    ])));
    const service = new EventsService(config(), { fetchNewMessages, now: () => NOW });

    const result = await service.poll({
      name: MESSAGE_NEW_EVENT,
      arguments: { conversationIds: ['19:a@thread.v2'], includeOwn: true },
      cursor: SCAN_CURSOR,
    });

    expect(fetchNewMessages).toHaveBeenCalledWith(expect.objectContaining({ includeOwn: true }));
    expect(result.events.map(e => e.data.messageId)).toEqual(['1']);
  });

  it('reports hasMore for a truncated scan and truncated for scan failures', async () => {
    const fetchNewMessages = vi.fn(async () => ok(scan([item()], {
      truncated: true,
      errors: [{ conversationId: '19:bad@thread.v2', code: ErrorCode.API_ERROR, message: 'boom' }],
    })));
    const service = new EventsService(config(), { fetchNewMessages, now: () => NOW });

    const result = await service.poll({ name: MESSAGE_NEW_EVENT, cursor: SCAN_CURSOR });

    expect(result.hasMore).toBe(true);
    expect(result.truncated).toBe(true);
    expect(result.events).toHaveLength(1);
  });

  it('cuts a batch at maxEvents and resumes without losing the rest', async () => {
    const fetchNewMessages = vi.fn(async () => ok(scan([
      item({ messageId: String(NOW - 20_000), conversationId: '19:a@thread.v2' }),
      item({ messageId: String(NOW - 10_000), conversationId: '19:b@thread.v2' }),
    ])));
    const service = new EventsService(config(), { fetchNewMessages, now: () => NOW });

    const result = await service.poll({
      name: MESSAGE_NEW_EVENT,
      cursor: SCAN_CURSOR,
      maxEvents: 1,
    });

    expect(result.events.map(e => e.data.conversationId)).toEqual(['19:a@thread.v2']);
    expect(result.hasMore).toBe(true);
    const cursor = decodeNewMessagesCursor(result.cursor);
    expect(cursor?.since).toBe(NOW - 60_000);
    expect(cursor?.chats['19:a@thread.v2']).toBe(String(NOW - 20_000));
    expect(cursor?.chats['19:b@thread.v2']).toBeUndefined();
  });

  it('backs off without failing the call on a throttled upstream', async () => {
    const fetchNewMessages = vi.fn(async () => err(createError(ErrorCode.RATE_LIMITED, 'slow down', {
      retryAfterMs: 90_000,
    })));
    const service = new EventsService(config(), { fetchNewMessages, now: () => NOW });

    const result = await service.poll({ name: MESSAGE_NEW_EVENT, cursor: SCAN_CURSOR });

    expect(result.events).toEqual([]);
    expect(result.cursor).toBe(SCAN_CURSOR);
    expect(result.nextPollMs).toBe(120_000);
  });

  it('surfaces an authentication failure so the client can re-login', async () => {
    const fetchNewMessages = vi.fn(async () => err(createError(ErrorCode.AUTH_EXPIRED, 'expired')));
    const service = new EventsService(config(), { fetchNewMessages, now: () => NOW });

    await expect(service.poll({ name: MESSAGE_NEW_EVENT, cursor: SCAN_CURSOR }))
      .rejects.toMatchObject({ message: expect.stringContaining('authentication failed') });
  });

  it('restarts from now with truncated=true when the cursor is older than maxAgeMs', async () => {
    const staleCursor = encodeNewMessagesCursor({
      version: 1,
      since: NOW - 10 * 60_000,
      chats: { '19:chat@thread.v2': String(NOW - 9 * 60_000) },
    });
    const fetchNewMessages = vi.fn();
    const service = new EventsService(config(), { fetchNewMessages, now: () => NOW });

    const result = await service.poll({
      name: MESSAGE_NEW_EVENT,
      cursor: staleCursor,
      maxAgeMs: 60_000,
    });

    expect(fetchNewMessages).not.toHaveBeenCalled();
    expect(result.truncated).toBe(true);
    expect(result.events).toEqual([]);
  });

  it('treats a current per-chat watermark as fresh even when the floor is old', async () => {
    const perEventCursor = encodeNewMessagesCursor({
      version: 1,
      since: NOW - 4 * 60 * 60_000,
      chats: { '19:chat@thread.v2': String(NOW - 30_000) },
    });
    const fetchNewMessages = vi.fn(async () => ok(scan()));
    const service = new EventsService(config(), { fetchNewMessages, now: () => NOW });

    const result = await service.poll({
      name: MESSAGE_NEW_EVENT,
      cursor: perEventCursor,
      maxAgeMs: 10 * 60_000,
    });

    expect(fetchNewMessages).toHaveBeenCalled();
    expect(result.truncated).toBe(false);
  });

  it('rejects a cursor from somebody else', async () => {
    const service = new EventsService(config(), { now: () => NOW });
    await expect(service.poll({ name: MESSAGE_NEW_EVENT, cursor: 'nope' }))
      .rejects.toBeInstanceOf(McpError);
    // An activity cursor is not a scan cursor.
    await expect(service.poll({ name: MESSAGE_NEW_EVENT, cursor: ACTIVITY_CURSOR }))
      .rejects.toBeInstanceOf(McpError);
  });
});

describe('poll: teams.mention and teams.reply', () => {
  it('bootstraps the activity feed without emitting existing items', async () => {
    const fetchActivity = vi.fn(async () => ok(activityPage([activity()], 'start-sync')));
    const service = new EventsService(config(), { fetchActivity, now: () => NOW });

    const result = await service.poll({ name: MENTION_EVENT, cursor: null });

    expect(fetchActivity).toHaveBeenCalledWith({ limit: DEFAULT_ACTIVITY_LIMIT });
    expect(result.events).toEqual([]);
    expect(result.truncated).toBe(false);
    expect(decodeActivityCursor(result.cursor)).toEqual({
      version: 1,
      syncState: 'start-sync',
      at: NOW,
    });
  });

  it('delivers only mentions for teams.mention', async () => {
    const fetchActivity = vi.fn(async () => ok(activityPage([
      activity({ id: '1', type: 'mention' }),
      activity({ id: '2', type: 'reply' }),
      activity({ id: '3', type: 'reaction' }),
    ])));
    const service = new EventsService(config(), { fetchActivity, now: () => NOW });

    const result = await service.poll({ name: MENTION_EVENT, cursor: ACTIVITY_CURSOR });

    expect(fetchActivity).toHaveBeenCalledWith({
      limit: DEFAULT_ACTIVITY_LIMIT,
      syncState: ACTIVITY_SYNC,
    });
    expect(result.events.map(e => e.data.messageId)).toEqual(['1']);
    expect(result.events[0].name).toBe(MENTION_EVENT);
    expect(decodeActivityCursor(result.cursor)?.syncState).toBe('next-sync-token');
    expect(result.truncated).toBe(false);
  });

  it('delivers only replies for teams.reply', async () => {
    const fetchActivity = vi.fn(async () => ok(activityPage([
      activity({ id: '1', type: 'mention' }),
      activity({ id: '2', type: 'reply' }),
    ])));
    const service = new EventsService(config(), { fetchActivity, now: () => NOW });

    const result = await service.poll({ name: REPLY_EVENT, cursor: ACTIVITY_CURSOR });

    expect(result.events.map(e => e.data.messageId)).toEqual(['2']);
    expect(result.events[0].name).toBe(REPLY_EVENT);
  });

  it('keeps the batch-start cursor on events so a reconnect can replay them', async () => {
    const fetchActivity = vi.fn(async () => ok(activityPage([activity()])));
    const service = new EventsService(config(), { fetchActivity, now: () => NOW });

    const result = await service.poll({ name: MENTION_EVENT, cursor: ACTIVITY_CURSOR });

    expect(result.events[0].cursor).toBe(ACTIVITY_CURSOR);
  });

  it('marks a full activity page as truncated', async () => {
    const fullPage = Array.from({ length: DEFAULT_ACTIVITY_LIMIT }, (_, index) =>
      activity({ id: String(index + 1), type: 'mention' })
    );
    const fetchActivity = vi.fn(async () => ok(activityPage(fullPage)));
    const service = new EventsService(config(), { fetchActivity, now: () => NOW });

    const result = await service.poll({ name: MENTION_EVENT, cursor: ACTIVITY_CURSOR });

    expect(result.events).toHaveLength(DEFAULT_ACTIVITY_LIMIT);
    expect(result.truncated).toBe(true);
  });

  it('backs off on an activity failure without losing the cursor', async () => {
    const fetchActivity = vi.fn(async () => err(createError(ErrorCode.API_ERROR, 'boom')));
    const service = new EventsService(config(), { fetchActivity, now: () => NOW });

    const result = await service.poll({ name: MENTION_EVENT, cursor: ACTIVITY_CURSOR });

    expect(result.events).toEqual([]);
    expect(result.cursor).toBe(ACTIVITY_CURSOR);
    expect(result.nextPollMs).toBe(120_000);
  });

  it('restarts a stale activity cursor', async () => {
    const staleCursor = encodeActivityCursor({
      version: 1,
      syncState: ACTIVITY_SYNC,
      at: NOW - 10 * 60_000,
    });
    const fetchActivity = vi.fn(async () => ok(activityPage([], 'fresh-sync')));
    const service = new EventsService(config(), { fetchActivity, now: () => NOW });

    const result = await service.poll({
      name: MENTION_EVENT,
      cursor: staleCursor,
      maxAgeMs: 60_000,
    });

    expect(fetchActivity).toHaveBeenCalledWith({ limit: DEFAULT_ACTIVITY_LIMIT });
    expect(result.truncated).toBe(true);
    expect(decodeActivityCursor(result.cursor)?.syncState).toBe('fresh-sync');
  });

  it('rejects a scan cursor for activity events', async () => {
    const service = new EventsService(config(), { now: () => NOW });
    await expect(service.poll({ name: MENTION_EVENT, cursor: SCAN_CURSOR }))
      .rejects.toBeInstanceOf(McpError);
  });
});

describe('stream: teams.message.new', () => {
  it('confirms the subscription, pushes events and advances a per-event cursor', async () => {
    const controller = new AbortController();
    const fetchNewMessages = vi.fn(async () => ok(scan([item()])));
    const service = new EventsService(config(), {
      fetchNewMessages,
      now: () => NOW,
      sleep: noSleep,
    });
    const notifications: Array<{ method: string; params: Record<string, unknown> }> = [];

    const result = await service.stream(
      { name: MESSAGE_NEW_EVENT, cursor: SCAN_CURSOR },
      7,
      async (method, params) => {
        notifications.push({ method, params });
        if (method === 'notifications/events/event') controller.abort();
      },
      controller.signal
    );

    expect(notifications[0]).toMatchObject({
      method: 'notifications/events/active',
      params: {
        cursor: SCAN_CURSOR,
        truncated: false,
        heartbeatMs: 30_000,
        _meta: { 'io.modelcontextprotocol/subscriptionId': 7 },
      },
    });

    const event = notifications.find(n => n.method === 'notifications/events/event');
    expect(event?.params).toMatchObject({
      eventId: '19:chat@thread.v2:100',
      _meta: { 'io.modelcontextprotocol/subscriptionId': 7 },
    });
    const cursor = decodeNewMessagesCursor(event?.params.cursor as string);
    expect(cursor?.chats['19:chat@thread.v2']).toBe('100');
    expect(cursor?.chats['19:other@thread.v2']).toBe(String(NOW - 30_000));

    expect(result).toEqual({
      resultType: 'complete',
      _meta: { 'io.modelcontextprotocol/subscriptionId': 7 },
    });
  });

  it('bootstraps from now when the client opens a stream without a cursor', async () => {
    const controller = new AbortController();
    const fetchNewMessages = vi.fn(async () => ok(scan()));
    const service = new EventsService(config(), {
      fetchNewMessages,
      now: () => NOW,
      sleep: noSleep,
    });
    const notifications: Array<{ method: string; params: Record<string, unknown> }> = [];

    await service.stream(
      { name: MESSAGE_NEW_EVENT, cursor: null, heartbeatMs: 1 },
      1,
      async (method, params) => {
        notifications.push({ method, params });
        controller.abort();
      },
      controller.signal
    );

    expect(notifications[0].method).toBe('notifications/events/active');
    expect(notifications[0].params.heartbeatMs).toBe(5_000); // clamped to the floor
  });

  it('reports upstream and scan failures without closing the stream', async () => {
    const controller = new AbortController();
    const fetchNewMessages = vi.fn(async () => ok(scan([], {
      errors: [{ conversationId: '19:bad@thread.v2', code: ErrorCode.API_ERROR, message: 'boom' }],
    })));
    const service = new EventsService(config(), {
      fetchNewMessages,
      now: () => NOW,
      sleep: noSleep,
    });
    const notifications: Array<{ method: string; params: Record<string, unknown> }> = [];

    const result = await service.stream(
      { name: MESSAGE_NEW_EVENT, cursor: SCAN_CURSOR },
      3,
      async (method, params) => {
        notifications.push({ method, params });
        if (method === 'notifications/events/error') controller.abort();
      },
      controller.signal
    );

    const error = notifications.find(n => n.method === 'notifications/events/error');
    expect(error?.params).toMatchObject({
      error: {
        message: '1 conversation(s) could not be scanned',
        data: { conversations: [{ conversationId: '19:bad@thread.v2' }] },
      },
    });
    expect(result.resultType).toBe('complete');
  });

  it('waits the configured interval after a full batch instead of draining back to back', async () => {
    const controller = new AbortController();
    const fetchNewMessages = vi.fn(async () => ok(scan([item()], { truncated: true })));
    const sleeps: number[] = [];
    const service = new EventsService(config(), {
      fetchNewMessages,
      now: () => NOW,
      sleep: async ms => {
        sleeps.push(ms);
        controller.abort();
      },
    });

    await service.stream(
      { name: MESSAGE_NEW_EVENT, cursor: SCAN_CURSOR },
      6,
      async () => {},
      controller.signal
    );

    expect(fetchNewMessages).toHaveBeenCalledTimes(1);
    expect(sleeps).toEqual([30_000]);
  });

  it('sends heartbeats while the stream is quiet', async () => {
    const controller = new AbortController();
    const fetchNewMessages = vi.fn(async () => ok(scan()));
    let clock = NOW;
    const service = new EventsService(config(), {
      fetchNewMessages,
      now: () => (clock += 10_000),
      sleep: noSleep,
    });
    const notifications: Array<{ method: string; params: Record<string, unknown> }> = [];

    await service.stream(
      { name: MESSAGE_NEW_EVENT, cursor: SCAN_CURSOR, heartbeatMs: 5_000 },
      4,
      async (method, params) => {
        notifications.push({ method, params });
        if (method === 'notifications/events/heartbeat') controller.abort();
      },
      controller.signal
    );

    const heartbeat = notifications.find(n => n.method === 'notifications/events/heartbeat');
    expect(heartbeat?.params).toMatchObject({
      cursor: NEXT_CURSOR,
      _meta: { 'io.modelcontextprotocol/subscriptionId': 4 },
    });
  });
});

describe('stream: teams.mention and teams.reply', () => {
  it('confirms from an activity cursor and pushes matching items', async () => {
    const controller = new AbortController();
    const fetchActivity = vi.fn(async () => ok(activityPage([
      activity({ id: '1', type: 'reply' }),
      activity({ id: '2', type: 'mention' }),
    ])));
    const service = new EventsService(config(), {
      fetchActivity,
      now: () => NOW,
      sleep: noSleep,
    });
    const notifications: Array<{ method: string; params: Record<string, unknown> }> = [];

    const result = await service.stream(
      { name: REPLY_EVENT, cursor: ACTIVITY_CURSOR },
      8,
      async (method, params) => {
        notifications.push({ method, params });
        if (method === 'notifications/events/event') controller.abort();
      },
      controller.signal
    );

    expect(notifications[0]).toMatchObject({
      method: 'notifications/events/active',
      params: { cursor: ACTIVITY_CURSOR, truncated: false, heartbeatMs: 30_000 },
    });
    const event = notifications.find(n => n.method === 'notifications/events/event');
    expect(event?.params).toMatchObject({
      eventId: '19:c@thread.v2:1',
      name: REPLY_EVENT,
      // The per-event cursor is the batch start, so a reconnect replays.
      cursor: ACTIVITY_CURSOR,
    });
    expect(result.resultType).toBe('complete');
  });

  it('bootstraps the activity feed before confirming when no cursor is given', async () => {
    const controller = new AbortController();
    const fetchActivity = vi.fn(async () => ok(activityPage([], 'boot-sync')));
    const service = new EventsService(config(), {
      fetchActivity,
      now: () => NOW,
      sleep: noSleep,
    });
    const notifications: Array<{ method: string; params: Record<string, unknown> }> = [];

    await service.stream(
      { name: MENTION_EVENT, cursor: null },
      2,
      async (method, params) => {
        notifications.push({ method, params });
        if (method === 'notifications/events/active') controller.abort();
      },
      controller.signal
    );

    expect(fetchActivity).toHaveBeenCalledWith({ limit: DEFAULT_ACTIVITY_LIMIT });
    expect(decodeActivityCursor(notifications[0].params.cursor as string)?.syncState).toBe('boot-sync');
  });

  it('keeps the stream open and retries when the activity bootstrap fails', async () => {
    const controller = new AbortController();
    const fetchActivity = vi.fn(async () => err(createError(ErrorCode.API_ERROR, 'down')));
    let attempts = 0;
    const service = new EventsService(config(), {
      fetchActivity,
      now: () => NOW,
      sleep: async () => {
        attempts++;
        if (attempts >= 2) controller.abort();
      },
    });
    const notifications: Array<{ method: string; params: Record<string, unknown> }> = [];

    await service.stream(
      { name: MENTION_EVENT, cursor: null },
      5,
      async (method, params) => {
        notifications.push({ method, params });
      },
      controller.signal
    );

    // Confirmation goes out immediately with a null cursor, then retries.
    expect(notifications[0]).toMatchObject({
      method: 'notifications/events/active',
      params: { cursor: null, truncated: false },
    });
    expect(fetchActivity.mock.calls.length).toBeGreaterThan(1);
  });
});

describe('cursor helpers', () => {
  it('advances only the event conversation watermark', () => {
    const next = advanceScanCursor(SCAN_CURSOR, '19:chat@thread.v2', String(NOW));
    const decoded = decodeNewMessagesCursor(next);
    expect(decoded?.chats['19:chat@thread.v2']).toBe(String(NOW));
    expect(decoded?.chats['19:other@thread.v2']).toBe(String(NOW - 30_000));
  });

  it('clamps heartbeat suggestions', () => {
    expect(clampHeartbeatMs(undefined)).toBe(30_000);
    expect(clampHeartbeatMs(1)).toBe(5_000);
    expect(clampHeartbeatMs(10 * 60_000)).toBe(60_000);
    expect(clampHeartbeatMs(15_000)).toBe(15_000);
  });

  it('sleeps zero even when the signal is already aborted', async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(sleepUnlessAborted(1000, controller.signal)).resolves.toBeUndefined();
  });
});
