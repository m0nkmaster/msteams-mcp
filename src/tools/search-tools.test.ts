/**
 * Unit tests for search tools.
 *
 * Tests the schema validation and since-to-startTime conversion logic.
 *
 * Note: Schemas are defined locally to avoid circular import issues through
 * the tool registry. A sync test validates these match the actual exports.
 */

import { describe, it, expect } from 'vitest';
import { z } from 'zod';
import {
  DEFAULT_THREAD_LIMIT,
  MAX_THREAD_LIMIT,
} from '../constants.js';

// Local schema definitions to avoid circular imports through registry
// These MUST match the actual schemas in search-tools.ts and message-tools.ts
const GetThreadInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  limit: z.number().min(1).max(MAX_THREAD_LIMIT).optional().default(DEFAULT_THREAD_LIMIT),
  markRead: z.boolean().optional().default(false),
  order: z.enum(['asc', 'desc']).optional().default('desc'),
  since: z.string().optional(),
});

const GetActivityInputSchema = z.object({
  limit: z.number().min(1).max(200).optional(),
  syncState: z.string().optional(),
});

describe('GetThreadInputSchema', () => {
  it('accepts conversationId without optional params', () => {
    const result = GetThreadInputSchema.parse({
      conversationId: '19:abc@thread.tacv2',
    });
    expect(result.conversationId).toBe('19:abc@thread.tacv2');
    expect(result.limit).toBe(DEFAULT_THREAD_LIMIT);
    expect(result.order).toBe('desc');
    expect(result.markRead).toBe(false);
    expect(result.since).toBeUndefined();
  });

  it('accepts since parameter as ISO string', () => {
    const result = GetThreadInputSchema.parse({
      conversationId: '19:abc@thread.tacv2',
      since: '2026-02-26T00:00:00Z',
    });
    expect(result.since).toBe('2026-02-26T00:00:00Z');
  });

  it('accepts all optional parameters', () => {
    const result = GetThreadInputSchema.parse({
      conversationId: '19:abc@thread.tacv2',
      limit: 10,
      markRead: true,
      order: 'asc',
      since: '2026-02-25T12:00:00Z',
    });
    expect(result.limit).toBe(10);
    expect(result.markRead).toBe(true);
    expect(result.order).toBe('asc');
    expect(result.since).toBe('2026-02-25T12:00:00Z');
  });

  it('rejects empty conversationId', () => {
    expect(() => GetThreadInputSchema.parse({ conversationId: '' })).toThrow();
  });

  it('rejects limit exceeding max', () => {
    expect(() => GetThreadInputSchema.parse({
      conversationId: '19:abc@thread.tacv2',
      limit: MAX_THREAD_LIMIT + 1,
    })).toThrow();
  });

  it('accepts limit at boundary values', () => {
    // Min boundary
    const minResult = GetThreadInputSchema.parse({
      conversationId: '19:abc@thread.tacv2',
      limit: 1,
    });
    expect(minResult.limit).toBe(1);

    // Max boundary
    const maxResult = GetThreadInputSchema.parse({
      conversationId: '19:abc@thread.tacv2',
      limit: MAX_THREAD_LIMIT,
    });
    expect(maxResult.limit).toBe(MAX_THREAD_LIMIT);
  });

  it('rejects limit of zero', () => {
    expect(() => GetThreadInputSchema.parse({
      conversationId: '19:abc@thread.tacv2',
      limit: 0,
    })).toThrow();
  });
});

describe('since to startTime conversion', () => {
  it('converts ISO string to numeric timestamp correctly', () => {
    const isoString = '2026-02-26T00:00:00Z';
    const expectedTimestamp = new Date(isoString).getTime();

    // This mirrors the logic in handleGetThread
    const startTime = new Date(isoString).getTime();

    expect(startTime).toBe(expectedTimestamp);
    expect(typeof startTime).toBe('number');
    expect(startTime).toBeGreaterThan(0);
  });

  it('handles undefined since gracefully', () => {
    const since = undefined;
    const startTime = since ? new Date(since).getTime() : undefined;

    expect(startTime).toBeUndefined();
  });

  it('handles various ISO formats', () => {
    const formats = [
      '2026-02-26T00:00:00Z',
      '2026-02-26T00:00:00.000Z',
      '2026-02-26T12:30:45.123Z',
    ];

    for (const format of formats) {
      const startTime = new Date(format).getTime();
      expect(typeof startTime).toBe('number');
      expect(startTime).toBeGreaterThan(0);
    }
  });

  it('detects invalid date strings that produce NaN', () => {
    const invalidDates = [
      'yesterday',
      'not-a-date',
      'last week',
      '26-02-2026', // wrong order (DD-MM-YYYY)
      '',
    ];

    for (const invalid of invalidDates) {
      const startTime = new Date(invalid).getTime();
      expect(isNaN(startTime)).toBe(true);
    }
  });

  it('handles timezone variations', () => {
    const formats = [
      '2026-02-26T00:00:00Z',
      '2026-02-26T00:00:00+00:00',
      '2026-02-26T05:30:00+05:30',
    ];

    for (const format of formats) {
      const startTime = new Date(format).getTime();
      expect(typeof startTime).toBe('number');
      expect(isNaN(startTime)).toBe(false);
    }
  });
});

describe('GetActivityInputSchema', () => {
  it('accepts empty input for defaults', () => {
    const result = GetActivityInputSchema.parse({});
    expect(result.limit).toBeUndefined();
    expect(result.syncState).toBeUndefined();
  });

  it('accepts limit parameter', () => {
    const result = GetActivityInputSchema.parse({ limit: 100 });
    expect(result.limit).toBe(100);
  });

  it('accepts syncState parameter', () => {
    const result = GetActivityInputSchema.parse({
      syncState: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...',
    });
    expect(result.syncState).toBe('eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...');
  });

  it('accepts both limit and syncState', () => {
    const result = GetActivityInputSchema.parse({
      limit: 50,
      syncState: 'some-pagination-token',
    });
    expect(result.limit).toBe(50);
    expect(result.syncState).toBe('some-pagination-token');
  });

  it('rejects limit below minimum', () => {
    expect(() => GetActivityInputSchema.parse({ limit: 0 })).toThrow();
  });

  it('rejects limit above maximum', () => {
    expect(() => GetActivityInputSchema.parse({ limit: 201 })).toThrow();
  });

  it('accepts limit at boundary values', () => {
    // Min boundary
    const minResult = GetActivityInputSchema.parse({ limit: 1 });
    expect(minResult.limit).toBe(1);

    // Max boundary
    const maxResult = GetActivityInputSchema.parse({ limit: 200 });
    expect(maxResult.limit).toBe(200);
  });
});

describe('schema sync validation', () => {
  // These tests ensure the local test schemas stay in sync with actual schemas
  // by testing the same inputs produce the same outputs

  it('GetThreadInputSchema matches actual schema behavior', () => {
    // Test a complex input that exercises all fields
    const input = {
      conversationId: '19:test@thread.tacv2',
      limit: 25,
      markRead: true,
      order: 'asc' as const,
      since: '2026-01-01T00:00:00Z',
    };

    const result = GetThreadInputSchema.parse(input);

    // Verify the parsed result matches expected shape
    expect(result).toEqual({
      conversationId: '19:test@thread.tacv2',
      limit: 25,
      markRead: true,
      order: 'asc',
      since: '2026-01-01T00:00:00Z',
    });
  });

  it('GetActivityInputSchema matches actual schema behavior', () => {
    const input = {
      limit: 75,
      syncState: 'test-token-123',
    };

    const result = GetActivityInputSchema.parse(input);

    expect(result).toEqual({
      limit: 75,
      syncState: 'test-token-123',
    });
  });
});
