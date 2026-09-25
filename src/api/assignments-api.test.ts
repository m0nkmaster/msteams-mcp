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
import { invalidateAssignmentsToken } from '../auth/token-extractor.js';
import { listMyAssignments, getAssignment, actOnSubmission } from './assignments-api.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireAssignmentsTokenAsync: vi.fn(),
}));

vi.mock('../auth/token-extractor.js', () => ({ invalidateAssignmentsToken: vi.fn() }));

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
    expect(qs.get('$orderby')).toBe('dueDateTime asc,id asc');
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
    mockHttp.mockResolvedValue(httpOk({ id: 's1', status: 'submitted' }));

    const result = await actOnSubmission('submit', 'c1', 'a1', 's1');

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.value).toMatchObject({ action: 'submit', submissionId: 's1', status: 'submitted' });
    const [url, opts] = mockHttp.mock.calls[0];
    expect(url).toContain('/submissions/s1/submit');
    expect((opts as { method: string }).method).toBe('POST');
  });

  it('POSTs to the unsubmit endpoint for unsubmit', async () => {
    mockHttp.mockResolvedValue(httpOk({ id: 's1', status: 'working' }));

    await actOnSubmission('unsubmit', 'c1', 'a1', 's1');

    const [url, opts] = mockHttp.mock.calls[0];
    expect(url).toContain('/submissions/s1/unsubmit');
    expect((opts as { method: string }).method).toBe('POST');
  });

  it('PATCHes the view endpoint for view', async () => {
    mockHttp.mockResolvedValue(httpOk({ id: 's1', status: 'working' }));

    await actOnSubmission('view', 'c1', 'a1', 's1');

    const [url, opts] = mockHttp.mock.calls[0];
    expect(url).toContain('/submissions/s1/view');
    expect((opts as { method: string }).method).toBe('PATCH');
  });
});


describe('Assignments regression cases', () => {
  it('exposes and follows server pagination without rebuilding its cursor', async () => {
    const nextLink = 'https://assignments.edu.cloud.microsoft/api/v1.0/edu/me/work?$skiptoken=opaque&$top=25';
    mockHttp.mockResolvedValueOnce(httpOk({ value: [{ id: 'a1' }], '@odata.nextLink': nextLink }));
    const first = await listMyAssignments();
    expect(first).toMatchObject({ ok: true, value: { nextLink: new URL(nextLink).toString() } });
    mockHttp.mockResolvedValueOnce(httpOk({ value: [{ id: 'a2' }] }));
    expect(await listMyAssignments({ nextLink })).toMatchObject({ ok: true, value: { assignments: [{ id: 'a2' }] } });
    expect(mockHttp.mock.calls[1][0]).toBe(new URL(nextLink).toString());
  });

  it('offers an offset continuation for a full page without nextLink', async () => {
    mockHttp.mockResolvedValue(httpOk({ value: [{ id: 'a1' }] }));
    const result = await listMyAssignments({ top: 1 });
    expect(result.ok).toBe(true);
    if (result.ok) expect(new URL(result.value.nextLink!).searchParams.get('$skip')).toBe('1');
  });

  it.each([
    'https://evil.example/api/v1.0/edu/me/work',
    'https://assignments.edu.cloud.microsoft/api/v1.0/edu/classes/c1',
    'https://assignments.edu.cloud.microsoft/api/v1.0/edu/me/work?$top=2.5',
  ])('rejects unsafe continuations: %s', async nextLink => {
    expect(await listMyAssignments({ nextLink })).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(mockHttp).not.toHaveBeenCalled();
    expect(mockToken).not.toHaveBeenCalled();
  });

  it.each(['../other', 's1?x=y', '.', '%2e%2e', 's1#fragment'])('rejects malformed IDs: %s', async id => {
    expect(await actOnSubmission('submit', 'c1', 'a1', id)).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(await getAssignment(id, 'a1')).toMatchObject({ ok: false });
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('rejects fractional page sizes before auth or HTTP', async () => {
    expect(await listMyAssignments({ top: 2.5 })).toMatchObject({ ok: false, error: { code: ErrorCode.INVALID_INPUT } });
    expect(mockToken).not.toHaveBeenCalled();
  });

  it('invalidates the rejected Assignments token on a 401', async () => {
    mockHttp.mockResolvedValue(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));
    expect(await listMyAssignments()).toMatchObject({ ok: false, error: { code: ErrorCode.AUTH_EXPIRED } });
    expect(invalidateAssignmentsToken).toHaveBeenCalledWith('fake-token');
  });

  it('does not trigger login for permission failures', async () => {
    mockHttp.mockResolvedValue(err(createError(ErrorCode.AUTH_REQUIRED, 'HTTP 403')));
    expect(await listMyAssignments()).toMatchObject({ ok: false, error: { code: ErrorCode.ACCESS_DENIED, retryable: false } });
    expect(invalidateAssignmentsToken).not.toHaveBeenCalled();
  });

  it('verifies turn-in with a separate GET and never replays the POST', async () => {
    mockHttp.mockResolvedValueOnce(httpOk('')).mockResolvedValueOnce(httpOk({ id: 's1', status: 'submitted' }));
    expect(await actOnSubmission('submit', 'c1', 'a1', 's1')).toMatchObject({ ok: true, value: { status: 'submitted' } });
    expect(mockHttp.mock.calls[0][1]).toMatchObject({ method: 'POST', maxRetries: 1 });
    expect(mockHttp.mock.calls[1][0]).toMatch(/\/submissions\/s1$/);
    expect(mockHttp.mock.calls[1][1]).toMatchObject({ method: 'GET' });
  });

  it.each(['<html>Login</html>', { id: 's1', status: 'working' }, { id: 'other', status: 'submitted' }])('does not report unconfirmed turn-in as success', async data => {
    mockHttp.mockResolvedValueOnce(httpOk({ id: 's1', status: 'submitted' })).mockResolvedValueOnce(httpOk(data));
    expect(await actOnSubmission('submit', 'c1', 'a1', 's1')).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR, retryable: false } });
  });

  it('does not trigger automatic mutation replay if verification needs authentication', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ id: 's1', status: 'submitted' }))
      .mockResolvedValueOnce(err(createError(ErrorCode.AUTH_EXPIRED, 'HTTP 401')));
    expect(await actOnSubmission('submit', 'c1', 'a1', 's1')).toMatchObject({ ok: false, error: { code: ErrorCode.API_ERROR, retryable: false } });
    expect(invalidateAssignmentsToken).toHaveBeenCalledWith('fake-token');
  });

  it('does not report an HTML view response as success', async () => {
    mockHttp.mockResolvedValue(httpOk('<html>Login</html>'));
    expect(await actOnSubmission('view', 'c1', 'a1', 's1')).toMatchObject({ ok: false });
  });
  it('marks a timed-out mutation as uncertain rather than inviting a retry', async () => {
    mockHttp.mockResolvedValue(err(createError(ErrorCode.TIMEOUT, 'Timed out', { retryable: true })));
    expect(await actOnSubmission('submit', 'c1', 'a1', 's1')).toMatchObject({ ok: false, error: { code: ErrorCode.TIMEOUT, retryable: false } });
    expect(mockHttp).toHaveBeenCalledTimes(1);
  });

  it('does not invent another cursor page after the server ends pagination', async () => {
    mockHttp.mockResolvedValue(httpOk({ value: [{ id: 'a1' }] }));
    const result = await listMyAssignments({ nextLink: 'https://assignments.edu.cloud.microsoft/api/v1.0/edu/me/work?$skiptoken=last&$top=1' });
    expect(result).toMatchObject({ ok: true, value: { nextLink: undefined } });
  });

});
