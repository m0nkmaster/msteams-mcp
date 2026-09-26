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
 * Read endpoints (my work, assignment detail, submissions) and the submit /
 * unsubmit actions are verified live against an education tenant; the "mark
 * viewed" PATCH against a captured web session.
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { requireAssignmentsTokenAsync } from '../utils/auth-guards.js';
import { ASSIGNMENTS_API, getAssignmentsHeaders } from '../utils/api-config.js';
import { ErrorCode, createError } from '../types/errors.js';
import { invalidateAccessToken } from '../auth/token-extractor.js';
import { DEFAULT_PAGE_SIZE, MAX_PAGE_SIZE } from '../constants.js';

/** IDs are opaque single path segments; reject dot segments and URL delimiters. */
export const ASSIGNMENT_ID_PATTERN = /^[a-zA-Z0-9_-]+$/;

function validIds(...ids: string[]): boolean {
  return ids.every(id => ASSIGNMENT_ID_PATTERN.test(id));
}

function invalidInput(message: string) {
  return err(createError(ErrorCode.INVALID_INPUT, message));
}

function handleAssignmentsError<T>(response: Result<T>, token: string): Result<T> {
  // Never surface AUTH_EXPIRED: it would make the server re-authenticate the
  // whole Teams session (and replay the tool) for this optional feature.
  if (!response.ok && response.error.code === ErrorCode.AUTH_EXPIRED) {
    invalidateAccessToken(token);
    return err(createError(ErrorCode.API_ERROR,
      'Assignments rejected its access token; it has been discarded, so a retry will request a new one.', {
        retryable: true,
        suggestions: ['Retry once; if it fails again, Assignments may be unavailable for this account'],
      }));
  }
  if (!response.ok && response.error.code === ErrorCode.AUTH_REQUIRED) {
    return err(createError(ErrorCode.ACCESS_DENIED, response.error.message, {
      retryable: false,
      suggestions: ['Check your permission to access this assignment'],
    }));
  }
  return response;
}

/** Never send the bearer token to a host or endpoint supplied by a continuation link. */
function validateNextLink(link: string, base?: string): string | null {
  try {
    const url = new URL(link, base);
    const allowed = new URL(ASSIGNMENTS_API.myWork());
    if (url.origin !== allowed.origin || url.pathname !== allowed.pathname || url.username || url.password || url.hash) return null;
    const top = Number(url.searchParams.get('$top') ?? DEFAULT_PAGE_SIZE);
    const skip = Number(url.searchParams.get('$skip') ?? 0);
    if (!Number.isInteger(top) || top < 1 || top > MAX_PAGE_SIZE || !Number.isSafeInteger(skip) || skip < 0) return null;
    return url.toString();
  } catch {
    return null;
  }
}

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

/** A file, form or link attached to an assignment or submission. */
export interface AssignmentAttachment {
  id?: string;
  name?: string;
  /** word | powerpoint | excel | file | form | link | … (from the resource's OData type). */
  type: string;
  /** Graph drive-item URL; pass to teams_download_file. Absent for forms and links. */
  fileUrl?: string;
  /** Web URL for non-file resources (e.g. a Microsoft Forms quiz or a link). */
  url?: string;
  /** True when each student gets their own copy (found in their submission's attachments). */
  copiedForEachStudent?: boolean;
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
  /** The student's working files, including their copies of distributed attachments. */
  attachments?: AssignmentAttachment[];
  /** Files as they were when last turned in. */
  submittedAttachments?: AssignmentAttachment[];
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
  /** Teacher-provided attachments (teams_get_assignment only). */
  attachments?: AssignmentAttachment[];
}

/** Result of listing the user's assignments. */
export interface ListAssignmentsResult {
  statusFilter: AssignmentStatusFilter;
  returned: number;
  assignments: Assignment[];
  /** Follow this link for the next page. A full page may require an empty final fetch. */
  nextLink?: string;
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

interface RawResource {
  id?: string;
  distributeForStudentWork?: boolean;
  resource?: {
    '@odata.type'?: string;
    displayName?: string;
    fileUrl?: string;
    link?: string;
    viewUrl?: string;
  } | null;
}

interface RawSubmission {
  id?: string;
  status?: string;
  submittedDateTime?: string;
  returnedDateTime?: string;
  outcomes?: RawOutcome[];
  resources?: RawResource[];
  submittedResources?: RawResource[];
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
  resources?: RawResource[];
}

interface RawListResponse {
  '@odata.nextLink'?: string;
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

function parseAttachments(resources?: RawResource[]): AssignmentAttachment[] | undefined {
  if (!resources) return undefined;
  return resources.map(({ id, distributeForStudentWork, resource }) => ({
    id,
    name: resource?.displayName,
    // '#microsoft.education.assignments.api.educationPowerPointResource' -> 'powerpoint'
    type: (resource?.['@odata.type'] ?? '').replace(/^.*\.education/, '').replace(/Resource$/, '').toLowerCase() || 'unknown',
    fileUrl: resource?.fileUrl,
    url: resource?.link ?? resource?.viewUrl,
    copiedForEachStudent: distributeForStudentWork || undefined,
  }));
}

function parseSubmission(sub?: RawSubmission): SubmissionSummary | undefined {
  if (!sub) return undefined;
  return {
    id: sub.id,
    status: sub.status,
    submittedDateTime: sub.submittedDateTime,
    returnedDateTime: sub.returnedDateTime,
    outcome: parseOutcomes(sub.outcomes),
    attachments: parseAttachments(sub.resources),
    submittedAttachments: parseAttachments(sub.submittedResources),
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
    attachments: parseAttachments(raw.resources),
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

/** Recover the slice a continuation URL actually queries, whatever the caller passed. */
function statusFilterFromUrl(url: string): AssignmentStatusFilter | undefined {
  const filter = new URL(url).searchParams.get('$filter');
  return (['active', 'completed', 'all'] as const).find(s => buildStatusFilter(s) === filter);
}

/**
 * Offset links we synthesised, keyed to the IDs of the page that produced them.
 * If following one returns that same page, the service ignored $skip and
 * offering another offset link would loop forever.
 */
const syntheticPages = new Map<string, string>();
const MAX_SYNTHETIC_PAGES = 50;

function rememberSyntheticPage(link: string, pageIds: string): void {
  if (syntheticPages.size >= MAX_SYNTHETIC_PAGES) {
    syntheticPages.delete(syntheticPages.keys().next().value!);
  }
  syntheticPages.set(link, pageIds);
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
  options: { statusFilter?: AssignmentStatusFilter; top?: number; nextLink?: string } = {}
): Promise<Result<ListAssignmentsResult>> {
  let statusFilter = options.statusFilter ?? 'active';
  const top = options.top ?? DEFAULT_PAGE_SIZE;
  if (!Number.isInteger(top) || top < 1 || top > MAX_PAGE_SIZE) {
    return invalidInput(`top must be an integer from 1 to ${MAX_PAGE_SIZE}`);
  }
  const params = new URLSearchParams();
  const filter = buildStatusFilter(statusFilter);
  if (filter) params.set('$filter', filter);
  params.set('$top', String(top));
  params.set('$orderby', statusFilter === 'active' ? 'dueDateTime asc,createdDateTime asc' : 'dueDateTime desc,createdDateTime asc');
  params.set('$expand', 'submissions($expand=outcomes)');
  const url = options.nextLink ? validateNextLink(options.nextLink) : `${ASSIGNMENTS_API.myWork()}?${params.toString()}`;
  if (!url) return invalidInput('nextLink must be a continuation URL for the Assignments work endpoint');
  if (options.nextLink) statusFilter = statusFilterFromUrl(url) ?? statusFilter;
  const tokenResult = await requireAssignmentsTokenAsync();
  if (!tokenResult.ok) return tokenResult;

  const response = await httpRequest<RawListResponse>(url, {
    method: 'GET',
    headers: getAssignmentsHeaders(tokenResult.value),
  });

  if (!response.ok) return handleAssignmentsError(response, tokenResult.value);

  const data = response.value.data;
  if (!data || !Array.isArray(data.value)) {
    return err(createError(ErrorCode.API_ERROR, 'Invalid Assignments list response', { retryable: false }));
  }
  const items = data.value;
  const pageIds = items.map(item => item.id).join(',');
  const previousPageIds = syntheticPages.get(url);
  syntheticPages.delete(url);
  if (previousPageIds !== undefined && previousPageIds === pageIds) {
    return err(createError(ErrorCode.API_ERROR,
      'The Assignments service ignored the page offset and returned the previous page again; no further pages can be fetched.',
      { retryable: false, suggestions: ['Use a narrower statusFilter or a larger top to see more assignments'] }));
  }
  let nextLink = data['@odata.nextLink'];
  if (nextLink) {
    nextLink = validateNextLink(nextLink, url) ?? undefined;
    if (!nextLink) return err(createError(ErrorCode.API_ERROR, 'Invalid Assignments continuation URL', { retryable: false }));
  } else {
    // Some responses omit nextLink when $top was used. Expose an offset page
    // rather than silently cutting off all older work at the tool's size cap.
    const next = new URL(url);
    const pageSize = Number(next.searchParams.get('$top') ?? top);
    if (!next.searchParams.has('$skiptoken') && items.length > 0 && items.length >= pageSize) {
      next.searchParams.set('$skip', String(Number(next.searchParams.get('$skip') ?? 0) + items.length));
      nextLink = next.toString();
      rememberSyntheticPage(nextLink, pageIds);
    }
  }
  return ok({ statusFilter, returned: items.length, assignments: items.map(parseAssignment), nextLink });
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
  if (!validIds(classId, assignmentId)) return invalidInput('Invalid classId or assignmentId');
  const tokenResult = await requireAssignmentsTokenAsync();
  if (!tokenResult.ok) return tokenResult;

  // Expanding submissions on the assignment itself returns empty outcomes, so
  // grades need the submissions collection, as the Teams client requests it.
  const headers = getAssignmentsHeaders(tokenResult.value);
  const submissionsParams = new URLSearchParams({ '$expand': 'outcomes,resources,submittedResources' });
  const assignmentParams = new URLSearchParams({ '$expand': 'resources' });
  const [response, submissions] = await Promise.all([
    httpRequest<RawAssignment>(`${ASSIGNMENTS_API.assignment(classId, assignmentId)}?${assignmentParams.toString()}`, { method: 'GET', headers }),
    httpRequest<{ value?: RawSubmission[] }>(
      `${ASSIGNMENTS_API.submissions(classId, assignmentId)}?${submissionsParams.toString()}`,
      { method: 'GET', headers },
    ),
  ]);

  if (!response.ok) return handleAssignmentsError(response, tokenResult.value);
  if (!submissions.ok) return handleAssignmentsError(submissions, tokenResult.value);
  if (!response.value.data || typeof response.value.data.id !== 'string' || !Array.isArray(submissions.value.data?.value)) {
    return err(createError(ErrorCode.API_ERROR, 'Invalid assignment detail response', { retryable: false }));
  }
  return ok(parseAssignment({ ...response.value.data, submissions: submissions.value.data.value }));
}

/** The submission action to perform on the caller's own submission. */
export type SubmissionAction = 'submit' | 'unsubmit' | 'view';

/**
 * Performs an action on the caller's own submission: turn in (submit), undo
 * turn-in (unsubmit), or mark viewed. These change state on the user's real
 * Teams account, so callers should confirm intent before invoking.
 *
 * submit/unsubmit (POST, no body) are verified live; `view` (PATCH) against a
 * captured session.
 */
export async function actOnSubmission(
  action: SubmissionAction,
  classId: string,
  assignmentId: string,
  submissionId: string
): Promise<Result<{ action: SubmissionAction; submissionId: string; status?: string }>> {
  if (!validIds(classId, assignmentId, submissionId)) return invalidInput('Invalid assignment or submission ID');
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
    // Do not replay a mutation after a timeout or server error.
    maxRetries: 1,
    // Graph parity specifies no request body; do not guess a private API payload.
  });

  if (!response.ok) {
    if (response.error.retryable && response.error.code !== ErrorCode.RATE_LIMITED) {
      return err({ ...response.error, retryable: false,
        suggestions: ['The action may have completed. Read the submission state before repeating it'] });
    }
    return handleAssignmentsError(response, tokenResult.value);
  }
  let status = response.value.data?.status;
  if (action !== 'view') {
    const verification = await httpRequest<RawSubmission>(
      ASSIGNMENTS_API.submission(classId, assignmentId, submissionId),
      { method: 'GET', headers: getAssignmentsHeaders(tokenResult.value) },
    );
    if (!verification.ok) handleAssignmentsError(verification, tokenResult.value);
    status = verification.ok ? verification.value.data?.status : undefined;
    const expected = action === 'submit' ? 'submitted' : 'working';
    if (!verification.ok || verification.value.data?.id !== submissionId || status !== expected) {
      return err(createError(ErrorCode.API_ERROR,
        `The ${action} request was accepted, but submission state could not be confirmed as ${expected}. Check the submission before repeating the action.`,
        { retryable: false, suggestions: ['Read the assignment to check the current submission state; do not automatically repeat the action'] }));
    }
  } else if (response.value.status !== 204 && (typeof status !== 'string' || response.value.data?.id !== submissionId)) {
    return err(createError(ErrorCode.API_ERROR, 'The view request returned an unexpected response; its result is unconfirmed.', { retryable: false }));
  }
  return ok({ action, submissionId, status });
}
