/**
 * Event type registry for the `io.modelcontextprotocol/events` extension.
 *
 * The names and schemas here are the contract advertised through
 * `events/list`; `EventsService` maps new-message scans and activity-feed
 * items onto them.
 */

import { z } from 'zod';
import { McpError, ErrorCode as JsonRpcErrorCode } from '@modelcontextprotocol/sdk/types.js';
import type { NewMessageItem } from '../api/chatsvc-new-messages.js';
import type { ActivityItem } from '../api/chatsvc-activity.js';
import { NEW_MESSAGES_SNIPPET_MAX_LENGTH } from '../constants.js';

/** Event fired for every new message (optionally limited to conversations). */
export const MESSAGE_NEW_EVENT = 'teams.message.new';

/** Event fired when a message @mentions the authenticated user. */
export const MENTION_EVENT = 'teams.mention';

/** Event fired when a reply arrives in a conversation you are part of. */
export const REPLY_EVENT = 'teams.reply';

/** All event type names this server offers. */
export type EventTypeName = typeof MESSAGE_NEW_EVENT | typeof MENTION_EVENT | typeof REPLY_EVENT;

/** A delivery mode from SEP-3415. Webhook is not offered by a stdio server. */
export type EventDeliveryMode = 'poll' | 'push' | 'webhook';

/** An event type as returned by `events/list`. */
export interface EventTypeDefinition {
  name: EventTypeName;
  description: string;
  delivery: EventDeliveryMode[];
  inputSchema: Record<string, unknown>;
  payloadSchema: Record<string, unknown>;
}

/** JSON Schema for the payload shared by all three Teams events. */
export const MESSAGE_PAYLOAD_SCHEMA: Record<string, unknown> = {
  type: 'object',
  properties: {
    conversationId: { type: 'string', description: 'Conversation the message belongs to.' },
    messageId: { type: 'string', description: 'Stable Teams message ID.' },
    senderName: { type: 'string', description: 'Display name of the sender.' },
    sentAt: { type: 'string', format: 'date-time', description: 'When the message arrived (ISO 8601).' },
    snippet: {
      type: 'string',
      maxLength: NEW_MESSAGES_SNIPPET_MAX_LENGTH,
      description: 'Plain-text preview of the message, at most 200 characters.',
    },
  },
  required: ['conversationId', 'messageId', 'senderName', 'sentAt', 'snippet'],
};

/** Arguments accepted by {@link MESSAGE_NEW_EVENT}. */
export const MessageNewArgumentsSchema = z.object({
  conversationIds: z.array(z.string().min(1)).min(1).optional(),
  includeOwn: z.boolean().optional(),
});

/** Arguments for the mention and reply events: none. */
export const NoArgumentsSchema = z.object({}).strict();

/** Event types in the order `events/list` publishes them. */
export const EVENT_TYPES: readonly EventTypeDefinition[] = [
  {
    name: MESSAGE_NEW_EVENT,
    description:
      'Fires when a message arrives in one of your recent Teams conversations. Filter with conversationIds and choose whether your own messages count with includeOwn. Channel thread replies are not in the channel feed; teams.reply covers those.',
    delivery: ['poll', 'push'],
    inputSchema: {
      type: 'object',
      properties: {
        conversationIds: {
          type: 'array',
          items: { type: 'string' },
          minItems: 1,
          description: 'Only deliver messages from these conversation IDs. Omit for all recent conversations.',
        },
        includeOwn: {
          type: 'boolean',
          default: false,
          description: 'Include messages you sent yourself (default false).',
        },
      },
    },
    payloadSchema: MESSAGE_PAYLOAD_SCHEMA,
  },
  {
    name: MENTION_EVENT,
    description:
      'Fires when a message @mentions you, including mentions inside channel threads. Served from the Teams activity feed.',
    delivery: ['poll', 'push'],
    inputSchema: { type: 'object', properties: {} },
    payloadSchema: MESSAGE_PAYLOAD_SCHEMA,
  },
  {
    name: REPLY_EVENT,
    description:
      'Fires when a reply arrives in one of your conversations, including replies inside channel threads. Served from the Teams activity feed.',
    delivery: ['poll', 'push'],
    inputSchema: { type: 'object', properties: {} },
    payloadSchema: MESSAGE_PAYLOAD_SCHEMA,
  },
];

/** Returns the definition for an event name, if this server offers it. */
export function getEventType(name: string): EventTypeDefinition | undefined {
  return EVENT_TYPES.find(event => event.name === name);
}

/** A validated subscription: event name plus its filters. */
export interface EventSubscription {
  name: EventTypeName;
  /** Conversation filter for {@link MESSAGE_NEW_EVENT}; undefined means all. */
  conversationIds?: Set<string>;
  /** Whether the caller's own messages count. */
  includeOwn: boolean;
}

/**
 * Validates `(name, arguments)` from an `events/poll` or `events/stream`
 * request. Throws `InvalidParams` for an unknown event or bad arguments, which
 * is the error the extension specifies for both cases.
 */
export function resolveSubscription(name: string, args: unknown): EventSubscription {
  if (!getEventType(name)) {
    throw new McpError(
      JsonRpcErrorCode.InvalidParams,
      `Unknown event type: ${name}. Call events/list for the supported event types.`
    );
  }

  if (name === MESSAGE_NEW_EVENT) {
    const parsed = MessageNewArgumentsSchema.safeParse(args ?? {});
    if (!parsed.success) {
      throw new McpError(
        JsonRpcErrorCode.InvalidParams,
        `Invalid arguments for ${name}: ${parsed.error.message}`
      );
    }
    return {
      name,
      includeOwn: parsed.data.includeOwn ?? false,
      conversationIds: parsed.data.conversationIds
        ? new Set(parsed.data.conversationIds)
        : undefined,
    };
  }

  const parsed = NoArgumentsSchema.safeParse(args ?? {});
  if (!parsed.success) {
    throw new McpError(
      JsonRpcErrorCode.InvalidParams,
      `${name} takes no subscription arguments.`
    );
  }
  return { name: name as EventTypeName, includeOwn: false };
}

/**
 * Returns true when a scanned message belongs to the subscription. Own-message
 * and scan filters are applied upstream (the scan takes `includeOwn`), so this
 * only narrows by conversation.
 */
export function matchesScanSubscription(
  subscription: EventSubscription,
  item: NewMessageItem
): boolean {
  return !subscription.conversationIds || subscription.conversationIds.has(item.conversationId);
}

/** An event occurrence as defined by SEP-3415. */
export interface EventOccurrence {
  eventId: string;
  name: EventTypeName;
  timestamp: string;
  data: {
    conversationId: string;
    messageId: string;
    senderName: string;
    sentAt: string;
    snippet: string;
  };
  cursor?: string;
}

/** Normalises a timestamp to ISO 8601, falling back to the input. */
function toIso(timestamp: string): string {
  const parsed = Date.parse(timestamp);
  return Number.isFinite(parsed) ? new Date(parsed).toISOString() : timestamp;
}

/** Builds the shared payload and event identity. */
function occurrence(
  name: EventTypeName,
  fields: {
    conversationId: string;
    messageId: string;
    senderName: string;
    sentAt: string;
    snippet: string;
  },
  cursor?: string
): EventOccurrence {
  const sentAt = toIso(fields.sentAt);
  return {
    // Teams message IDs are only unique inside a conversation.
    eventId: `${fields.conversationId}:${fields.messageId}`,
    name,
    timestamp: sentAt,
    data: {
      conversationId: fields.conversationId,
      messageId: fields.messageId,
      senderName: fields.senderName,
      sentAt,
      snippet: fields.snippet.slice(0, NEW_MESSAGES_SNIPPET_MAX_LENGTH),
    },
    ...(cursor ? { cursor } : {}),
  };
}

/** Maps a new-message scan item to an occurrence. */
export function newMessageToOccurrence(
  item: NewMessageItem,
  cursor?: string
): EventOccurrence {
  return occurrence(MESSAGE_NEW_EVENT, {
    conversationId: item.conversationId,
    messageId: item.messageId,
    senderName: item.senderName,
    sentAt: item.sentAt,
    snippet: item.snippet,
  }, cursor);
}

/**
 * Maps an activity-feed item to a mention or reply occurrence.
 * Returns null for items without a usable source conversation.
 */
export function activityToOccurrence(
  name: typeof MENTION_EVENT | typeof REPLY_EVENT,
  item: ActivityItem,
  cursor?: string
): EventOccurrence | null {
  if (!item.id || !item.conversationId) return null;
  return occurrence(name, {
    conversationId: item.conversationId,
    messageId: item.id,
    senderName: item.sender.displayName ?? '',
    sentAt: item.timestamp,
    snippet: item.content,
  }, cursor);
}
