/**
 * Unit tests for people tool input schemas.
 *
 * Tests schema validation for the presence lookup tool.
 *
 * Note: The schema is defined locally to avoid circular import issues through
 * the tool registry. It MUST match the actual schema in people-tools.ts.
 */

import { describe, it, expect } from 'vitest';
import { z } from 'zod';

// Local schema definition to avoid circular imports through registry.
// This MUST match GetPresenceInputSchema in people-tools.ts.
const GetPresenceInputSchema = z.object({
  userIds: z.array(z.string().min(1)).min(1),
});

describe('GetPresenceInputSchema', () => {
  it('accepts a single user id', () => {
    const result = GetPresenceInputSchema.parse({
      userIds: ['8:orgid:abc-123'],
    });
    expect(result.userIds).toEqual(['8:orgid:abc-123']);
  });

  it('accepts multiple user ids', () => {
    const result = GetPresenceInputSchema.parse({
      userIds: ['8:orgid:abc', 'def-456'],
    });
    expect(result.userIds).toHaveLength(2);
  });

  it('rejects an empty userIds array', () => {
    expect(() => GetPresenceInputSchema.parse({ userIds: [] })).toThrow();
  });

  it('rejects an empty string user id', () => {
    expect(() => GetPresenceInputSchema.parse({ userIds: [''] })).toThrow();
  });

  it('rejects a missing userIds field', () => {
    expect(() => GetPresenceInputSchema.parse({})).toThrow();
  });
});
