/**
 * Unit tests for event type descriptors, subscriptions and payload mapping.
 */

import { describe, it, expect } from 'vitest';
import { McpError } from '@modelcontextprotocol/sdk/types.js';
import type { NewMessageItem } from '../api/chatsvc-new-messages.js';
import type { ActivityItem } from '../api/chatsvc-activity.js';
import {
  EVENT_TYPES,
  MENTION_EVENT,
  MESSAGE_NEW_EVENT,
  REPLY_EVENT,
  activityToOccurrence,
  getEventType,
  matchesScanSubscription,
  newMessageToOccurrence,
  resolveSubscription,
} from './event-types.js';

/** A scan item with sensible defaults. */
function item(overrides: Partial<NewMessageItem> = {}): NewMessageItem {
  return {
    conversationId: '19:chat@thread.v2',
    topic: 'Chat',
    chatType: 'group',
    messageId: '100',
    senderName: 'Bob',
    sentAt: '2026-01-01T10:00:00.000Z',
    snippet: 'hello',
    ...overrides,
  };
}

/** An activity item with sensible defaults. */
function activity(overrides: Partial<ActivityItem> = {}): ActivityItem {
  return {
    id: '200',
    type: 'mention',
    content: 'ping',
    contentType: 'RichText/Html',
    sender: { mri: '8:orgid:someone', displayName: 'Alice' },
    timestamp: '2026-01-01T11:00:00.000Z',
    conversationId: '19:c@thread.v2',
    ...overrides,
  };
}

describe('EVENT_TYPES', () => {
  it('offers the three Teams events with poll and push delivery', () => {
    expect(EVENT_TYPES.map(e => e.name)).toEqual([MESSAGE_NEW_EVENT, MENTION_EVENT, REPLY_EVENT]);
    for (const event of EVENT_TYPES) {
      expect(event.delivery).toEqual(['poll', 'push']);
      expect(event.inputSchema).toMatchObject({ type: 'object' });
      expect(event.payloadSchema).toMatchObject({
        type: 'object',
        required: ['conversationId', 'messageId', 'senderName', 'sentAt', 'snippet'],
      });
    }
  });

  it('documents the message.new filters', () => {
    const properties = (getEventType(MESSAGE_NEW_EVENT)?.inputSchema as {
      properties: Record<string, unknown>;
    }).properties;
    expect(properties).toHaveProperty('conversationIds');
    expect(properties).toHaveProperty('includeOwn');
    expect(getEventType('teams.nope')).toBeUndefined();
  });
});

describe('resolveSubscription', () => {
  it('defaults includeOwn to false', () => {
    expect(resolveSubscription(MESSAGE_NEW_EVENT, undefined)).toEqual({
      name: MESSAGE_NEW_EVENT,
      includeOwn: false,
      conversationIds: undefined,
    });
  });

  it('reads filters for message.new', () => {
    const subscription = resolveSubscription(MESSAGE_NEW_EVENT, {
      conversationIds: ['19:a@thread.v2'],
      includeOwn: true,
    });
    expect(subscription.includeOwn).toBe(true);
    expect(subscription.conversationIds?.has('19:a@thread.v2')).toBe(true);
  });

  it('accepts no arguments for mention and reply', () => {
    expect(resolveSubscription(MENTION_EVENT, undefined).name).toBe(MENTION_EVENT);
    expect(resolveSubscription(REPLY_EVENT, {}).name).toBe(REPLY_EVENT);
  });

  it('rejects unknown event names with InvalidParams', () => {
    expect(() => resolveSubscription('teams.nope', {})).toThrowError(McpError);
    try {
      resolveSubscription('teams.nope', {});
    } catch (error) {
      expect((error as McpError).code).toBe(-32602);
    }
  });

  it('rejects malformed arguments and arguments to no-argument events', () => {
    expect(() => resolveSubscription(MESSAGE_NEW_EVENT, { conversationIds: 'nope' })).toThrowError(McpError);
    expect(() => resolveSubscription(MENTION_EVENT, { includeOwn: true })).toThrowError(McpError);
  });
});

describe('matchesScanSubscription', () => {
  it('narrows message.new by conversation', () => {
    const subscription = resolveSubscription(MESSAGE_NEW_EVENT, {
      conversationIds: ['19:a@thread.v2'],
    });
    expect(matchesScanSubscription(subscription, item({ conversationId: '19:a@thread.v2' }))).toBe(true);
    expect(matchesScanSubscription(subscription, item({ conversationId: '19:b@thread.v2' }))).toBe(false);
  });

  it('accepts every conversation when no filter is set', () => {
    const subscription = resolveSubscription(MESSAGE_NEW_EVENT, undefined);
    expect(matchesScanSubscription(subscription, item({ conversationId: '19:any@thread.v2' }))).toBe(true);
  });
});

describe('newMessageToOccurrence', () => {
  it('maps a scan item to the documented payload', () => {
    const occurrence = newMessageToOccurrence(item({
      conversationId: '19:c@thread.v2',
      messageId: '42',
      senderName: 'Alice',
      sentAt: '2026-01-02T00:00:00.000Z',
      snippet: 'ping',
    }), 'cursor-1');

    expect(occurrence).toEqual({
      eventId: '19:c@thread.v2:42',
      name: MESSAGE_NEW_EVENT,
      timestamp: '2026-01-02T00:00:00.000Z',
      data: {
        conversationId: '19:c@thread.v2',
        messageId: '42',
        senderName: 'Alice',
        sentAt: '2026-01-02T00:00:00.000Z',
        snippet: 'ping',
      },
      cursor: 'cursor-1',
    });
  });
});

describe('activityToOccurrence', () => {
  it('maps a mention activity to the shared payload', () => {
    const occurrence = activityToOccurrence(MENTION_EVENT, activity(), 'cursor-2');
    expect(occurrence).toEqual({
      eventId: '19:c@thread.v2:200',
      name: MENTION_EVENT,
      timestamp: '2026-01-01T11:00:00.000Z',
      data: {
        conversationId: '19:c@thread.v2',
        messageId: '200',
        senderName: 'Alice',
        sentAt: '2026-01-01T11:00:00.000Z',
        snippet: 'ping',
      },
      cursor: 'cursor-2',
    });
  });

  it('normalises timestamps and caps the snippet', () => {
    const occurrence = activityToOccurrence(REPLY_EVENT, activity({
      timestamp: '2026-01-01T12:00:00+01:00',
      content: 'a'.repeat(300),
    }));
    expect(occurrence?.timestamp).toBe('2026-01-01T11:00:00.000Z');
    expect(occurrence?.data.snippet).toHaveLength(200);
  });

  it('returns null when the activity has no source conversation', () => {
    expect(activityToOccurrence(MENTION_EVENT, activity({ conversationId: undefined }))).toBeNull();
    expect(activityToOccurrence(MENTION_EVENT, activity({ id: '' }))).toBeNull();
  });
});
