/**
 * Unit tests for meeting/calendar tool input schemas.
 *
 * Tests schema validation, defaults, and boundary behaviour for the
 * calendar management tools (create/get/update/cancel/respond/schedule).
 *
 * Note: Schemas are defined locally to avoid circular import issues through
 * the tool registry. These MUST match the actual schemas in meeting-tools.ts.
 */

import { describe, it, expect } from 'vitest';
import { z } from 'zod';

// Local schema definitions to avoid circular imports through registry.
// These MUST match the actual schemas in meeting-tools.ts.
const AttendeeSchema = z.object({
  email: z.string().email(),
  name: z.string().optional(),
});

const CreateMeetingInputSchema = z.object({
  subject: z.string().min(1),
  startTime: z.string().min(1),
  endTime: z.string().min(1),
  attendees: z.array(AttendeeSchema).optional(),
  body: z.string().optional(),
  isOnlineMeeting: z.boolean().optional().default(true),
  location: z.string().optional(),
});

const GetMeetingInputSchema = z.object({
  eventId: z.string().min(1),
});

const UpdateMeetingInputSchema = z.object({
  eventId: z.string().min(1),
  subject: z.string().min(1).optional(),
  startTime: z.string().min(1).optional(),
  endTime: z.string().min(1).optional(),
  attendees: z.array(AttendeeSchema).optional(),
  body: z.string().optional(),
  location: z.string().optional(),
});

const CancelMeetingInputSchema = z.object({
  eventId: z.string().min(1),
});

const RespondToMeetingInputSchema = z.object({
  eventId: z.string().min(1),
  response: z.enum(['accept', 'tentativelyAccept', 'decline']),
  comment: z.string().optional(),
  sendResponse: z.boolean().optional().default(true),
  proposedNewTime: z.object({
    start: z.string().min(1),
    end: z.string().min(1),
  }).optional(),
});

const GetScheduleInputSchema = z.object({
  schedules: z.array(z.string().email()).min(1),
  startTime: z.string().min(1),
  endTime: z.string().min(1),
  availabilityViewInterval: z.number().min(5).max(1440).optional().default(30),
});

// ─────────────────────────────────────────────────────────────────────────────
// CreateMeetingInputSchema
// ─────────────────────────────────────────────────────────────────────────────

describe('CreateMeetingInputSchema', () => {
  it('accepts the minimal required fields and defaults isOnlineMeeting', () => {
    const result = CreateMeetingInputSchema.parse({
      subject: 'Sync',
      startTime: '2026-06-05T10:00:00Z',
      endTime: '2026-06-05T10:30:00Z',
    });
    expect(result.subject).toBe('Sync');
    expect(result.isOnlineMeeting).toBe(true);
    expect(result.attendees).toBeUndefined();
  });

  it('accepts attendees with valid emails and optional names', () => {
    const result = CreateMeetingInputSchema.parse({
      subject: 'Sync',
      startTime: '2026-06-05T10:00:00Z',
      endTime: '2026-06-05T10:30:00Z',
      attendees: [
        { email: 'chris@company.com', name: 'Chris' },
        { email: 'noah@company.com' },
      ],
    });
    expect(result.attendees).toHaveLength(2);
    expect(result.attendees?.[1].name).toBeUndefined();
  });

  it('allows overriding isOnlineMeeting to false', () => {
    const result = CreateMeetingInputSchema.parse({
      subject: 'In person',
      startTime: '2026-06-05T10:00:00Z',
      endTime: '2026-06-05T10:30:00Z',
      isOnlineMeeting: false,
      location: 'Room 1',
    });
    expect(result.isOnlineMeeting).toBe(false);
    expect(result.location).toBe('Room 1');
  });

  it('rejects an empty subject', () => {
    expect(() => CreateMeetingInputSchema.parse({
      subject: '',
      startTime: '2026-06-05T10:00:00Z',
      endTime: '2026-06-05T10:30:00Z',
    })).toThrow();
  });

  it('rejects a missing startTime', () => {
    expect(() => CreateMeetingInputSchema.parse({
      subject: 'Sync',
      endTime: '2026-06-05T10:30:00Z',
    })).toThrow();
  });

  it('rejects an invalid attendee email', () => {
    expect(() => CreateMeetingInputSchema.parse({
      subject: 'Sync',
      startTime: '2026-06-05T10:00:00Z',
      endTime: '2026-06-05T10:30:00Z',
      attendees: [{ email: 'not-an-email' }],
    })).toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// GetMeetingInputSchema / CancelMeetingInputSchema
// ─────────────────────────────────────────────────────────────────────────────

describe('GetMeetingInputSchema', () => {
  it('accepts a non-empty eventId', () => {
    const result = GetMeetingInputSchema.parse({ eventId: 'abc123' });
    expect(result.eventId).toBe('abc123');
  });

  it('rejects an empty eventId', () => {
    expect(() => GetMeetingInputSchema.parse({ eventId: '' })).toThrow();
  });
});

describe('CancelMeetingInputSchema', () => {
  it('accepts a non-empty eventId', () => {
    expect(CancelMeetingInputSchema.parse({ eventId: 'abc123' }).eventId).toBe('abc123');
  });

  it('rejects a missing eventId', () => {
    expect(() => CancelMeetingInputSchema.parse({})).toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// UpdateMeetingInputSchema
// ─────────────────────────────────────────────────────────────────────────────

describe('UpdateMeetingInputSchema', () => {
  it('accepts eventId with no other fields', () => {
    const result = UpdateMeetingInputSchema.parse({ eventId: 'abc123' });
    expect(result.eventId).toBe('abc123');
    expect(result.subject).toBeUndefined();
  });

  it('accepts a partial update with only a new time', () => {
    const result = UpdateMeetingInputSchema.parse({
      eventId: 'abc123',
      startTime: '2026-06-05T11:00:00Z',
      endTime: '2026-06-05T11:30:00Z',
    });
    expect(result.startTime).toBe('2026-06-05T11:00:00Z');
  });

  it('rejects an empty subject when provided', () => {
    expect(() => UpdateMeetingInputSchema.parse({
      eventId: 'abc123',
      subject: '',
    })).toThrow();
  });

  it('rejects an invalid attendee email', () => {
    expect(() => UpdateMeetingInputSchema.parse({
      eventId: 'abc123',
      attendees: [{ email: 'nope' }],
    })).toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// RespondToMeetingInputSchema
// ─────────────────────────────────────────────────────────────────────────────

describe('RespondToMeetingInputSchema', () => {
  it('accepts a valid response and defaults sendResponse to true', () => {
    const result = RespondToMeetingInputSchema.parse({
      eventId: 'abc123',
      response: 'accept',
    });
    expect(result.response).toBe('accept');
    expect(result.sendResponse).toBe(true);
  });

  it('accepts a decline with a proposed new time and comment', () => {
    const result = RespondToMeetingInputSchema.parse({
      eventId: 'abc123',
      response: 'decline',
      comment: 'Conflicts with another call',
      sendResponse: false,
      proposedNewTime: {
        start: '2026-06-06T10:00:00Z',
        end: '2026-06-06T10:30:00Z',
      },
    });
    expect(result.sendResponse).toBe(false);
    expect(result.proposedNewTime?.start).toBe('2026-06-06T10:00:00Z');
  });

  it('rejects an unknown response value', () => {
    expect(() => RespondToMeetingInputSchema.parse({
      eventId: 'abc123',
      response: 'maybe',
    })).toThrow();
  });

  it('rejects a proposedNewTime missing the end field', () => {
    expect(() => RespondToMeetingInputSchema.parse({
      eventId: 'abc123',
      response: 'decline',
      proposedNewTime: { start: '2026-06-06T10:00:00Z' },
    })).toThrow();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// GetScheduleInputSchema
// ─────────────────────────────────────────────────────────────────────────────

describe('GetScheduleInputSchema', () => {
  it('accepts a single schedule and defaults the interval to 30', () => {
    const result = GetScheduleInputSchema.parse({
      schedules: ['chris@company.com'],
      startTime: '2026-06-05T09:00:00Z',
      endTime: '2026-06-05T18:00:00Z',
    });
    expect(result.schedules).toEqual(['chris@company.com']);
    expect(result.availabilityViewInterval).toBe(30);
  });

  it('rejects an empty schedules array', () => {
    expect(() => GetScheduleInputSchema.parse({
      schedules: [],
      startTime: '2026-06-05T09:00:00Z',
      endTime: '2026-06-05T18:00:00Z',
    })).toThrow();
  });

  it('rejects a non-email entry in schedules', () => {
    expect(() => GetScheduleInputSchema.parse({
      schedules: ['not-an-email'],
      startTime: '2026-06-05T09:00:00Z',
      endTime: '2026-06-05T18:00:00Z',
    })).toThrow();
  });

  it('accepts the interval at boundary values', () => {
    const min = GetScheduleInputSchema.parse({
      schedules: ['a@company.com'],
      startTime: '2026-06-05T09:00:00Z',
      endTime: '2026-06-05T18:00:00Z',
      availabilityViewInterval: 5,
    });
    expect(min.availabilityViewInterval).toBe(5);

    const max = GetScheduleInputSchema.parse({
      schedules: ['a@company.com'],
      startTime: '2026-06-05T09:00:00Z',
      endTime: '2026-06-05T18:00:00Z',
      availabilityViewInterval: 1440,
    });
    expect(max.availabilityViewInterval).toBe(1440);
  });

  it('rejects an interval below the minimum', () => {
    expect(() => GetScheduleInputSchema.parse({
      schedules: ['a@company.com'],
      startTime: '2026-06-05T09:00:00Z',
      endTime: '2026-06-05T18:00:00Z',
      availabilityViewInterval: 4,
    })).toThrow();
  });

  it('rejects an interval above the maximum', () => {
    expect(() => GetScheduleInputSchema.parse({
      schedules: ['a@company.com'],
      startTime: '2026-06-05T09:00:00Z',
      endTime: '2026-06-05T18:00:00Z',
      availabilityViewInterval: 1441,
    })).toThrow();
  });
});
