/**
 * EDU Assignments API client.
 *
 * Talks to the Teams Assignments service at `assignments.edu.cloud.microsoft`
 * (backed by OneNote EDU). This is the API behind the Teams "Assignments" tab —
 * not Microsoft Graph, though the data model mirrors Graph's `education` types.
 *
 * Auth: a Bearer token for the Assignments app (audience `8f348934-...`), obtained
 * via `requireAssignmentsTokenAsync()`. That token is minted by the HTTP token
 * refresh (the `EduAssignments` entry in `token-refresh-http.ts`).
 *
 * Read endpoints (my work, assignment detail, submissions) and the "mark viewed"
 * PATCH were verified against a captured web session. The submit / unsubmit
 * actions follow Microsoft Graph education parity on the same path shape.
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok } from '../types/result.js';
import { requireAssignmentsTokenAsync, handleSubstrateError } from '../utils/auth-guards.js';
import { ASSIGNMENTS_API, getAssignmentsHeaders } from '../utils/api-config.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Which slice of the user's assignments to list. */
export type AssignmentStatusFilter = 'active' | 'completed' | 'all';

/** A grade/feedback outcome on a submission (best-effort parse of Graph education shapes). */
export interface SubmissionOutcome {
  /** Awarded points, if a points grade has been given. */
  points?: number;
  /** Points visible to the student (after release). */
  publishedPoints?: number;
  /** Feedback text, if given. */
  feedback?: string;
  /** Feedback text visible to the student (after release). */
  publishedFeedback?: string;
}

/** A compact view of the signed-in user's submission for an assignment. */
export interface SubmissionSummary {
  id?: string;
  /** working | submitted | returned | reassigned | excused (Graph education states). */
  status?: string;
  submittedDateTime?: string;
  returnedDateTime?: string;
  /** Merged grade/feedback from the submission's outcomes. */
  outcome?: SubmissionOutcome;
}

/** A single assignment as surfaced to the caller. */
export interface Assignment {
  id: string;
  classId: string;
  displayName: string;
  status?: string;
  dueDateTime?: string;
  assignedDateTime?: string;
  closeDateTime?: string;
  allowLateSubmissions?: boolean;
  /** Max points, when the assignment is points-graded. */
  maxPoints?: number;
  /** Plain-ish instructions content (HTML as stored by the service). */
  instructions?: string;
  /** Deep link that opens the assignment in Teams. */
  webUrl?: string;
  /** The caller's own submission summary, when expanded. */
  submission?: SubmissionSummary;
}

/** Result of listing the user's assignments. */
export interface ListAssignmentsResult {
  statusFilter: AssignmentStatusFilter;
  returned: number;
  assignments: Assignment[];
}

// ─────────────────────────────────────────────────────────────────────────────
// Raw response shapes (partial — only fields we read)
// ─────────────────────────────────────────────────────────────────────────────

interface RawOutcome {
  '@odata.type'?: string;
  points?: { points?: number } | null;
  publishedPoints?: { points?: number } | null;
  feedback?: { text?: { content?: string } | string } | null;
  publishedFeedback?: { text?: { content?: string } | string } | null;
}

interface RawSubmission {
  id?: string;
  status?: string;
  submittedDateTime?: string;
  returnedDateTime?: string;
  outcomes?: RawOutcome[];
}

interface RawAssignment {
  id: string;
  classId: string;
  displayName: string;
  status?: string;
  dueDateTime?: string;
  assignedDateTime?: string;
  closeDateTime?: string;
  allowLateSubmissions?: boolean;
  grading?: { maxPoints?: number } | null;
  instructions?: { content?: string } | null;
  webUrl?: string;
  submissions?: RawSubmission[];
}

interface RawListResponse {
  value?: RawAssignment[];
}

// ─────────────────────────────────────────────────────────────────────────────
// Parsing
// ─────────────────────────────────────────────────────────────────────────────

function feedbackText(fb: RawOutcome['feedback']): string | undefined {
  if (!fb?.text) return undefined;
  return typeof fb.text === 'string' ? fb.text : fb.text.content;
}

/** Merges a submission's outcome array into a single grade/feedback object. */
function parseOutcomes(outcomes?: RawOutcome[]): SubmissionOutcome | undefined {
  if (!outcomes?.length) return undefined;
  const merged: SubmissionOutcome = {};
  for (const o of outcomes) {
    if (o.points?.points != null) merged.points = o.points.points;
    if (o.publishedPoints?.points != null) merged.publishedPoints = o.publishedPoints.points;
    const fb = feedbackText(o.feedback);
    if (fb) merged.feedback = fb;
    const pfb = feedbackText(o.publishedFeedback);
    if (pfb) merged.publishedFeedback = pfb;
  }
  return Object.keys(merged).length > 0 ? merged : undefined;
}

function parseSubmission(sub?: RawSubmission): SubmissionSummary | undefined {
  if (!sub) return undefined;
  return {
    id: sub.id,
    status: sub.status,
    submittedDateTime: sub.submittedDateTime,
    returnedDateTime: sub.returnedDateTime,
    outcome: parseOutcomes(sub.outcomes),
  };
}

function parseAssignment(raw: RawAssignment): Assignment {
  return {
    id: raw.id,
    classId: raw.classId,
    displayName: raw.displayName,
    status: raw.status,
    dueDateTime: raw.dueDateTime ?? undefined,
    assignedDateTime: raw.assignedDateTime ?? undefined,
    closeDateTime: raw.closeDateTime ?? undefined,
    allowLateSubmissions: raw.allowLateSubmissions,
    maxPoints: raw.grading?.maxPoints ?? undefined,
    instructions: raw.instructions?.content ?? undefined,
    webUrl: raw.webUrl ?? undefined,
    submission: parseSubmission(raw.submissions?.[0]),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// $filter builders
// ─────────────────────────────────────────────────────────────────────────────

const STATUS_PREFIX = "microsoft.education.assignments.api.educationAssignmentStatus";

function buildStatusFilter(filter: AssignmentStatusFilter): string | null {
  switch (filter) {
    case 'active':
      return `status eq ${STATUS_PREFIX}'assigned' and isCompleted eq false`;
    case 'completed':
      return `isCompleted eq true`;
    case 'all':
      return null;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// API functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Lists the signed-in user's assignments across all their classes.
 *
 * Uses the `/edu/me/work` endpoint with the caller's submission (and its grade
 * outcomes) expanded inline, so a single call yields due dates, status, points,
 * and the user's own submission state.
 *
 * @param options.statusFilter - active (assigned, not completed) | completed | all. Default active.
 * @param options.top - max assignments to return (default 25).
 */
export async function listMyAssignments(
  options: { statusFilter?: AssignmentStatusFilter; top?: number } = {}
): Promise<Result<ListAssignmentsResult>> {
  const tokenResult = await requireAssignmentsTokenAsync();
  if (!tokenResult.ok) return tokenResult;

  const statusFilter = options.statusFilter ?? 'active';
  const top = options.top ?? 25;

  const params = new URLSearchParams();
  const filter = buildStatusFilter(statusFilter);
  if (filter) params.set('$filter', filter);
  params.set('$top', String(top));
  params.set('$orderby', 'dueDateTime desc');
  params.set('$expand', 'submissions($expand=outcomes)');

  const url = `${ASSIGNMENTS_API.myWork()}?${params.toString()}`;

  const response = await httpRequest<RawListResponse>(url, {
    method: 'GET',
    headers: getAssignmentsHeaders(tokenResult.value),
  });

  if (!response.ok) return handleSubstrateError(response);

  const items = response.value.data.value ?? [];
  return ok({
    statusFilter,
    returned: items.length,
    assignments: items.map(parseAssignment),
  });
}

/**
 * Gets a single assignment's full detail.
 *
 * @param classId - the class (group) ID the assignment belongs to.
 * @param assignmentId - the assignment ID.
 */
export async function getAssignment(
  classId: string,
  assignmentId: string
): Promise<Result<Assignment>> {
  const tokenResult = await requireAssignmentsTokenAsync();
  if (!tokenResult.ok) return tokenResult;

  const params = new URLSearchParams();
  params.set('$expand', 'submissions($expand=outcomes)');

  const url = `${ASSIGNMENTS_API.assignment(classId, assignmentId)}?${params.toString()}`;

  const response = await httpRequest<RawAssignment>(url, {
    method: 'GET',
    headers: getAssignmentsHeaders(tokenResult.value),
  });

  if (!response.ok) return handleSubstrateError(response);
  return ok(parseAssignment(response.value.data));
}

/** The submission action to perform on the caller's own submission. */
export type SubmissionAction = 'submit' | 'unsubmit' | 'view';

/**
 * Performs an action on the caller's own submission: turn in (submit), undo
 * turn-in (unsubmit), or mark viewed. These change state on the user's real
 * Teams account, so callers should confirm intent before invoking.
 *
 * Only `view` (PATCH) is verified against a captured session; submit/unsubmit
 * (POST) follow Microsoft Graph education parity.
 */
export async function actOnSubmission(
  action: SubmissionAction,
  classId: string,
  assignmentId: string,
  submissionId: string
): Promise<Result<{ action: SubmissionAction; submissionId: string; status?: string }>> {
  const tokenResult = await requireAssignmentsTokenAsync();
  if (!tokenResult.ok) return tokenResult;

  let url: string;
  let method: 'POST' | 'PATCH';
  switch (action) {
    case 'submit':
      url = ASSIGNMENTS_API.submissionSubmit(classId, assignmentId, submissionId);
      method = 'POST';
      break;
    case 'unsubmit':
      url = ASSIGNMENTS_API.submissionUnsubmit(classId, assignmentId, submissionId);
      method = 'POST';
      break;
    case 'view':
      url = ASSIGNMENTS_API.submissionView(classId, assignmentId, submissionId);
      method = 'PATCH';
      break;
  }

  const response = await httpRequest<RawSubmission>(url, {
    method,
    headers: getAssignmentsHeaders(tokenResult.value),
    // Submission actions take no body; PATCH view returns the updated submission.
  });

  if (!response.ok) return handleSubstrateError(response);
  return ok({
    action,
    submissionId,
    status: response.value.data?.status,
  });
}
