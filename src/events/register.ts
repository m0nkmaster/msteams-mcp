/**
 * Registers the `io.modelcontextprotocol/events` request handlers.
 *
 * The MCP SDK this server targets predates protocol revision 2026-07-28, so
 * the extension's request methods and notifications are handled at the wire
 * level: the server advertises the extension in its capabilities and answers
 * `events/*` itself. Everything is plain JSON-RPC, so clients that understand
 * the extension can use it without SDK support on our side.
 */

import { z } from 'zod';
import type { Server } from '@modelcontextprotocol/sdk/server/index.js';
import type { RequestHandlerExtra } from '@modelcontextprotocol/sdk/shared/protocol.js';
import type { ServerNotification, ServerRequest, ServerResult } from '@modelcontextprotocol/sdk/types.js';
import type { EventsConfig } from './config.js';
import { EventsService } from './service.js';

/** Extension identifier from SEP-3415. */
export const EVENTS_EXTENSION_ID = 'io.modelcontextprotocol/events';

/** Server capability payload advertised when events are enabled. */
export const EVENTS_CAPABILITY = { listChanged: false } as const;

const EventsListRequestSchema = z.object({
  method: z.literal('events/list'),
  params: z.object({ cursor: z.string().optional() }).optional(),
});

const EventsPollRequestSchema = z.object({
  method: z.literal('events/poll'),
  params: z.object({
    name: z.string().min(1),
    arguments: z.unknown().optional(),
    cursor: z.string().nullable().optional(),
    maxAgeMs: z.number().int().positive().optional(),
    maxEvents: z.number().int().positive().optional(),
  }),
});

const EventsStreamRequestSchema = z.object({
  method: z.literal('events/stream'),
  params: z.object({
    name: z.string().min(1),
    arguments: z.unknown().optional(),
    cursor: z.string().nullable().optional(),
    maxAgeMs: z.number().int().positive().optional(),
    heartbeatMs: z.number().int().positive().optional(),
  }),
});

/** Narrows an extension result to the SDK's (older) result union. */
function asServerResult(value: unknown): ServerResult {
  return value as ServerResult;
}

/**
 * Registers a handler whose schema is a project Zod 3 schema.
 *
 * The SDK bundles its own Zod 4 and types `setRequestHandler` against it, while
 * this project's tool schemas use Zod 3. The SDK parses Zod 3 schemas at
 * runtime through its compatibility layer, so only the types need bridging.
 */
function setExtensionHandler<Schema extends z.AnyZodObject>(
  server: Server,
  schema: Schema,
  handler: (
    request: z.output<Schema>,
    extra: RequestHandlerExtra<ServerRequest, ServerNotification>
  ) => Promise<unknown>
): void {
  server.setRequestHandler(schema as never, handler as never);
}

/**
 * Wires the three extension request methods onto a server.
 * Returns the service so callers (and tests) can inspect backoff state.
 */
export function registerEventHandlers(server: Server, config: EventsConfig): EventsService {
  const service = new EventsService(config);

  setExtensionHandler(server, EventsListRequestSchema, async () => {
    return asServerResult(service.list());
  });

  setExtensionHandler(server, EventsPollRequestSchema, async (request) => {
    return asServerResult(await service.poll(request.params));
  });

  setExtensionHandler(server, EventsStreamRequestSchema, async (request, extra) => {
    return asServerResult(await service.stream(
      request.params,
      extra.requestId,
      async (method, params) => {
        await extra.sendNotification({ method, params } as never);
      },
      extra.signal
    ));
  });

  return service;
}

