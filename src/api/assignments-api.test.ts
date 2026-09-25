/**
 * Unit tests for the EDU Assignments API client.
 *
 * Mocks the HTTP layer and the auth guard so we exercise the real URL/filter
 * building and response parsing without a live session.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { httpRequest } from '../utils/http.js';
import { requireAssignmentsTokenAsync } from '../utils/auth-guards.js';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { listMyAssignments, getAssignment, actOnSubmission } from './assignments-api.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireAssignmentsTokenAsync: vi.fn(),
  // handleSubstrateError is imported by the module under test; keep it a passthrough.
  handleSubstrateError: (r: unknown) => r,
}));

const mockHttp = vi.mocked(httpRequest);
const mockToken = vi.mocked(requireAssignmentsTokenAsync);

function httpOk<T>(data: T) {
  return ok({ status: 200, headers: new Headers(), data });
}

beforeEach(() => {
  vi.clearAllMocks();
  mockToken.mockResolvedValue(ok('fake-token'));
});

describe('listMyAssignments', () => {
  it('parses assignments and the caller submission with a merged grade outcome', async () => {
    mockHttp.mockResolvedValue(httpOk({
      value: [
        {
          id: 'a1',
          classId: 'c1',
          displayName: 'Causes of climate change',
          status: 'assigned',
          dueDateTime: '2026-09-23T23:59:59Z',
          allowLateSubmissions: true,
          grading: { maxPoints: 5 },
          instructions: { content: '<p>Read chapter 3</p>' },
          webUrl: 'https://teams.microsoft.com/l/...',
          submissions: [
            {
              id: 's1',
              status: 'returned',
              submittedDateTime: '2026-09-22T10:00:00Z',
              returnedDateTime: '2026-09-24T09:00:00Z',
              outcomes: [
                { '@odata.type': '#...educationPointsOutcome', points: { points: 4 }, publishedPoints: { points: 4 } },
                { '@odata.type': '#...educationFeedbackOutcome', feedback: { text: { content: 'Good work' } } },
              ],
            },
          ],
        },
      ],
    }));

    const result = await listMyAssignments();

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.returned).toBe(1);
    const a = result.value.assignments[0];
    expect(a).toMatchObject({
      id: 'a1',
      classId: 'c1',
      displayName: 'Causes of climate change',
      maxPoints: 5,
      instructions: '<p>Read chapter 3</p>',
    });
    expect(a.submission).toMatchObject({
      id: 's1',
      status: 'returned',
      outcome: { points: 4, publishedPoints: 4, feedback: 'Good work' },
    });
  });

  it("builds the 'active' filter and default query params by default", async () => {
    mockHttp.mockResolvedValue(httpOk({ value: [] }));

    await listMyAssignments();

    const url = mockHttp.mock.calls[0][0] as string;
    const qs = new URL(url).searchParams;
    expect(qs.get('$filter')).toBe("status eq microsoft.education.assignments.api.educationAssignmentStatus'assigned' and isCompleted eq false");
    expect(qs.get('$top')).toBe('25');
    expect(qs.get('$orderby')).toBe('dueDateTime desc');
    expect(qs.get('$expand')).toBe('submissions($expand=outcomes)');
  });

  it("omits the filter for statusFilter 'all' and honours top", async () => {
    mockHttp.mockResolvedValue(httpOk({ value: [] }));

    await listMyAssignments({ statusFilter: 'all', top: 10 });

    const qs = new URL(mockHttp.mock.calls[0][0] as string).searchParams;
    expect(qs.has('$filter')).toBe(false);
    expect(qs.get('$top')).toBe('10');
  });

  it("builds the 'completed' filter", async () => {
    mockHttp.mockResolvedValue(httpOk({ value: [] }));

    await listMyAssignments({ statusFilter: 'completed' });

    const qs = new URL(mockHttp.mock.calls[0][0] as string).searchParams;
    expect(qs.get('$filter')).toBe('isCompleted eq true');
  });

  it('handles assignments with no submissions or outcomes', async () => {
    mockHttp.mockResolvedValue(httpOk({
      value: [{ id: 'a1', classId: 'c1', displayName: 'No submission yet' }],
    }));

    const result = await listMyAssignments();
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value.assignments[0].submission).toBeUndefined();
    expect(result.value.assignments[0].maxPoints).toBeUndefined();
  });

  it('propagates an auth error from the token guard without calling http', async () => {
    mockToken.mockResolvedValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no token')));

    const result = await listMyAssignments();
    expect(result.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });
});

describe('getAssignment', () => {
  it('requests the assignment path with submissions expanded', async () => {
    mockHttp.mockResolvedValue(httpOk({ id: 'a1', classId: 'c1', displayName: 'Detail' }));

    const result = await getAssignment('c1', 'a1');

    expect(result.ok).toBe(true);
    const url = mockHttp.mock.calls[0][0] as string;
    expect(url).toContain('/edu/classes/c1/assignments/a1');
    expect(new URL(url).searchParams.get('$expand')).toBe('submissions($expand=outcomes)');
  });
});

describe('actOnSubmission', () => {
  it('POSTs to the submit endpoint for submit', async () => {
    mockHttp.mockResolvedValue(httpOk({ status: 'submitted' }));

    const result = await actOnSubmission('submit', 'c1', 'a1', 's1');

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toMatchObject({ action: 'submit', submissionId: 's1', status: 'submitted' });
    const [url, opts] = mockHttp.mock.calls[0];
    expect(url).toContain('/submissions/s1/submit');
    expect((opts as { method: string }).method).toBe('POST');
  });

  it('POSTs to the unsubmit endpoint for unsubmit', async () => {
    mockHttp.mockResolvedValue(httpOk({ status: 'working' }));

    await actOnSubmission('unsubmit', 'c1', 'a1', 's1');

    const [url, opts] = mockHttp.mock.calls[0];
    expect(url).toContain('/submissions/s1/unsubmit');
    expect((opts as { method: string }).method).toBe('POST');
  });

  it('PATCHes the view endpoint for view', async () => {
    mockHttp.mockResolvedValue(httpOk({ status: 'working' }));

    await actOnSubmission('view', 'c1', 'a1', 's1');

    const [url, opts] = mockHttp.mock.calls[0];
    expect(url).toContain('/submissions/s1/view');
    expect((opts as { method: string }).method).toBe('PATCH');
  });
});
