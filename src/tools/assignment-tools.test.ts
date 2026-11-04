import { describe, expect, it } from 'vitest';
import { ListAssignmentsInputSchema, GetAssignmentInputSchema, SubmissionActionInputSchema } from './index.js';

describe('Assignments tool input validation', () => {
  it('rejects fractional page sizes while accepting integer bounds', () => {
    expect(ListAssignmentsInputSchema.safeParse({ top: 2.5 }).success).toBe(false);
    expect(ListAssignmentsInputSchema.safeParse({ top: 1 }).success).toBe(true);
    expect(ListAssignmentsInputSchema.safeParse({ top: 100 }).success).toBe(true);
  });
  it.each(['..', '../assignment', 'id?x=y', 'id#hash', '%2f'])('rejects path manipulation: %s', id => {
    expect(GetAssignmentInputSchema.safeParse({ classId: id, assignmentId: 'a1' }).success).toBe(false);
    expect(SubmissionActionInputSchema.safeParse({ action: 'submit', classId: 'c1', assignmentId: 'a1', submissionId: id }).success).toBe(false);
  });
});
