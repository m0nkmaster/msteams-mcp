/**
 * Wire-level tests for the events extension.
 *
 * The SDK this project uses predates the extension, so its Client strips the
 * `capabilities.extensions` field from the initialize result and cannot issue
 * `events/*` requests with types. These tests therefore speak raw JSON-RPC
 * over the in-memory transport, which is exactly the contract a client that
 * implements SEP-3415 sees.
 */

import { describe, it, expect, afterEach, vi } from 'vitest';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import type { JSONRPCMessage } from '@modelcontextprotocol/sdk/types.js';
import { createServer } from '../server.js';
import { EVENTS_ENABLED_ENV } from './config.js';

interface RawResponse {
  result?: Record<string, unknown>;
  error?: { code: number; message: string };
}

interface RawClient {
  request(method: string, params?: unknown): Promise<Record<string, unknown>>;
  notify(method: string, params?: unknown): Promise<void>;
  close(): Promise<void>;
}

/** Starts the server and returns a minimal raw JSON-RPC client over stdio-like transport. */
async function rawConnect(): Promise<RawClient> {
  const server = await createServer();
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);

  const pending = new Map<number, (response: RawResponse) => void>();
  let nextId = 1;

  clientTransport.onmessage = (message: JSONRPCMessage) => {
    if (!('id' in message) || message.id === undefined) return;
    const resolve = pending.get(Number(message.id));
    if (!resolve) return;
    pending.delete(Number(message.id));
    resolve(message as RawResponse);
  };
  await clientTransport.start();

  return {
    request: (method, params) => new Promise((resolve, reject) => {
      const id = nextId++;
      pending.set(id, response => {
        if (response.error) {
          reject(Object.assign(new Error(response.error.message), { code: response.error.code }));
        } else {
          resolve(response.result ?? {});
        }
      });
      clientTransport.send({ jsonrpc: '2.0', id, method, params } as JSONRPCMessage)
        .catch(reject);
    }),
    notify: async (method, params) => {
      await clientTransport.send({ jsonrpc: '2.0', method, params } as JSONRPCMessage);
    },
    close: async () => {
      await clientTransport.close();
      await server.close();
    },
  };
}

/** Performs the initialize handshake and returns the server's capabilities. */
async function initialize(client: RawClient): Promise<Record<string, unknown>> {
  const result = await client.request('initialize', {
    protocolVersion: '2025-11-25',
    capabilities: {},
    clientInfo: { name: 'raw-events-test', version: '1' },
  });
  await client.notify('notifications/initialized');
  return result.capabilities as Record<string, unknown>;
}

afterEach(() => {
  vi.unstubAllEnvs();
});

describe('events extension', () => {
  it('advertises the extension and lists the Teams event types when enabled', async () => {
    vi.stubEnv(EVENTS_ENABLED_ENV, 'true');
    const client = await rawConnect();
    try {
      const capabilities = await initialize(client);
      expect(capabilities.extensions).toEqual({
        'io.modelcontextprotocol/events': { listChanged: false },
      });

      const response = await client.request('events/list');
      expect(response.resultType).toBe('complete');
      const events = response.events as Array<{
        name: string;
        delivery: string[];
        inputSchema: unknown;
        payloadSchema: unknown;
      }>;
      expect(events.map(event => event.name)).toEqual([
        'teams.message.new',
        'teams.mention',
        'teams.reply',
      ]);
      for (const event of events) {
        expect(event.delivery).toEqual(['poll', 'push']);
        expect(event.inputSchema).toBeTruthy();
        expect(event.payloadSchema).toBeTruthy();
      }
    } finally {
      await client.close();
    }
  });

  it('bootstraps a poll from now without touching Teams', async () => {
    vi.stubEnv(EVENTS_ENABLED_ENV, 'true');
    const client = await rawConnect();
    try {
      await initialize(client);
      const response = await client.request('events/poll', {
        name: 'teams.message.new',
        cursor: null,
      });

      expect(response.resultType).toBe('complete');
      expect(response.events).toEqual([]);
      expect(response.truncated).toBe(false);
      expect(response.hasMore).toBe(false);
      expect(response.cursor).toMatch(/^[A-Za-z0-9_-]+$/);
      expect(response.nextPollMs).toBeGreaterThanOrEqual(15_000);
    } finally {
      await client.close();
    }
  });

  it('rejects an unknown event type with InvalidParams', async () => {
    vi.stubEnv(EVENTS_ENABLED_ENV, 'true');
    const client = await rawConnect();
    try {
      await initialize(client);
      await expect(client.request('events/poll', { name: 'teams.nope', cursor: null }))
        .rejects.toMatchObject({ code: -32602 });
    } finally {
      await client.close();
    }
  });

  it('does not advertise or answer events when disabled', async () => {
    const client = await rawConnect();
    try {
      const capabilities = await initialize(client);
      expect(capabilities.extensions).toBeUndefined();

      await expect(client.request('events/list'))
        .rejects.toMatchObject({ code: -32601 });
    } finally {
      await client.close();
    }
  });
});
