/**
 * EDU Assignments tool handlers.
 *
 * Surface the Teams "Assignments" feature: list the signed-in user's
 * assignments, read a single assignment's detail and attachments, download an
 * attachment, and act on the user's own submission (turn in, undo turn-in,
 * mark viewed).
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import {
  listMyAssignments,
  ASSIGNMENT_ID_PATTERN,
  getAssignment,
  actOnSubmission,
} from '../api/assignments-api.js';
import { downloadDriveItem } from '../api/graph-files-api.js';

import { DEFAULT_PAGE_SIZE, MAX_PAGE_SIZE } from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const ListAssignmentsInputSchema = z.object({
  statusFilter: z.enum(['active', 'completed', 'all']).optional().default('active'),
  nextLink: z.string().url().optional(),
  top: z.number().int().min(1).max(MAX_PAGE_SIZE).optional().default(DEFAULT_PAGE_SIZE),
});

export const GetAssignmentInputSchema = z.object({
  classId: z.string().regex(ASSIGNMENT_ID_PATTERN),
  assignmentId: z.string().regex(ASSIGNMENT_ID_PATTERN),
});

export const DownloadAssignmentFileInputSchema = z.object({
  fileUrl: z.string().url(),
  outputPath: z.string().min(1),
});

export const SubmissionActionInputSchema = z.object({
  action: z.enum(['submit', 'unsubmit', 'view']),
  classId: z.string().regex(ASSIGNMENT_ID_PATTERN),
  assignmentId: z.string().regex(ASSIGNMENT_ID_PATTERN),
  submissionId: z.string().regex(ASSIGNMENT_ID_PATTERN),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const listAssignmentsToolDefinition: Tool = {
  name: 'teams_list_assignments',
  description:
    "List the signed-in user's Microsoft Teams assignments across all their classes (the Teams \"Assignments\" tab, for EDU tenants). Active work is ordered earliest due first. Results are paged: pass nextLink back until it is absent. The service applies the page size before the status filter, so any page (not only the last) may be short or empty while more follow; keep following nextLink, or use a larger top (e.g. 100) to need fewer calls. nextLink preserves the original query, status slice and page size, and the returned statusFilter always reflects the slice actually queried. If the service cannot page further, the call returns an API_ERROR rather than repeating results. Returns each assignment's title, class ID, due date, status, max points, instructions, a Teams deep link (webUrl), and the user's own submission summary including submission state and any grade/feedback. Use statusFilter to choose which slice: 'active' (assigned and not yet completed — the default), 'completed' (turned in or returned), or 'all'. To act on or read the full detail of one result, pass its classId + id to teams_get_assignment or its submission's id to teams_submission_action. Only available on education tenants that use Assignments; on other accounts it returns an ACCESS_DENIED or AUTH_INTERACTION_REQUIRED error. That does not affect any other Teams tool, and teams_login is not needed.",
  inputSchema: {
    type: 'object',
    properties: {
      statusFilter: {
        type: 'string',
        enum: ['active', 'completed', 'all'],
        description: "Which assignments to return: 'active' (default), 'completed', or 'all'.",
      },
      nextLink: { type: 'string', description: 'Continuation URL returned by the previous call. This URL determines the query, status slice and page size; statusFilter is ignored when it is given.' },
      top: {
        type: 'integer',
        description: 'Maximum number of assignments to return (default 25, max 100).',
      },
    },
    required: [],
  },
};

const getAssignmentToolDefinition: Tool = {
  name: 'teams_get_assignment',
  description:
    "Get the full detail of a single Teams assignment, including instructions, due/close dates, max points, a Teams deep link, teacher attachments (name, type, and a fileUrl for files or a url for forms and links), and the signed-in user's own submission (state, grade, feedback, and their own working and turned-in files, including personal copies of attachments marked copiedForEachStudent). Download any attachment with a fileUrl using teams_download_assignment_file. Requires the classId and assignmentId, both available from teams_list_assignments (fields classId and id). Only available on education tenants that use Assignments; on other accounts it returns an ACCESS_DENIED or AUTH_INTERACTION_REQUIRED error. That does not affect any other Teams tool, and teams_login is not needed.",
  inputSchema: {
    type: 'object',
    properties: {
      classId: {
        type: 'string',
        description: 'The class (group) ID the assignment belongs to (the classId field from teams_list_assignments).',
      },
      assignmentId: {
        type: 'string',
        description: 'The assignment ID (the id field from teams_list_assignments).',
      },
    },
    required: ['classId', 'assignmentId'],
  },
};

const downloadAssignmentFileToolDefinition: Tool = {
  name: 'teams_download_assignment_file',
  description:
    "Download a Teams assignment attachment to a local file. Pass the fileUrl of an entry in teams_get_assignment's attachments (teacher resources), or in its submission.attachments / submission.submittedAttachments (the student's own copies and turned-in work). Only Microsoft Graph drive-item URLs are accepted; forms and links have no fileUrl and cannot be downloaded (use their url instead). Saves the raw file to an absolute path on the MCP server machine and returns its name, size, content type and SHA-256. The parent directory must exist and existing files are never overwritten. Streams to disk with no size limit; a transfer that stalls for 30 seconds is cancelled and the partial file removed. Uses Microsoft Graph via the existing Teams session (no extra sign-in); if Graph is unavailable, only this tool fails and no other Teams tool is affected.",
  inputSchema: {
    type: 'object',
    properties: {
      fileUrl: { type: 'string', description: "An attachment's fileUrl from teams_get_assignment (https://graph.microsoft.com/v1.0/drives/{id}/items/{id})." },
      outputPath: { type: 'string', description: 'Absolute destination file path on the MCP server machine. Must not already exist; the parent directory must exist.' },
    },
    required: ['fileUrl', 'outputPath'],
  },
};

const submissionActionToolDefinition: Tool = {
  name: 'teams_submission_action',
  description:
    "Act on the signed-in user's OWN submission for a Teams assignment. action='submit' turns it in, action='unsubmit' undoes a turn-in, action='view' marks it as viewed. This changes state on the user's real Teams account, so confirm the user's intent before calling with 'submit' or 'unsubmit'. Requires classId, assignmentId and submissionId — get classId and assignmentId from teams_list_assignments (fields classId and id) and submissionId from the submission.id in that result or from teams_get_assignment. Only available on education tenants that use Assignments; on other accounts it returns an ACCESS_DENIED or AUTH_INTERACTION_REQUIRED error. That does not affect any other Teams tool, and teams_login is not needed.",
  inputSchema: {
    type: 'object',
    properties: {
      action: {
        type: 'string',
        enum: ['submit', 'unsubmit', 'view'],
        description: "The action: 'submit' (turn in), 'unsubmit' (undo turn-in), or 'view' (mark viewed).",
      },
      classId: {
        type: 'string',
        description: 'The class (group) ID (the classId field from teams_list_assignments).',
      },
      assignmentId: {
        type: 'string',
        description: 'The assignment ID (the id field from teams_list_assignments).',
      },
      submissionId: {
        type: 'string',
        description: "The submission ID (the submission.id field from teams_list_assignments or teams_get_assignment).",
      },
    },
    required: ['action', 'classId', 'assignmentId', 'submissionId'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleListAssignments(
  input: z.infer<typeof ListAssignmentsInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await listMyAssignments({
    statusFilter: input.statusFilter,
    top: input.top,
    nextLink: input.nextLink,
  });
  return handleApiResult(result, (value) => ({
    statusFilter: value.statusFilter,
    returned: value.returned,
    assignments: value.assignments,
    nextLink: value.nextLink,
  }));
}

async function handleGetAssignment(
  input: z.infer<typeof GetAssignmentInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getAssignment(input.classId, input.assignmentId);
  return handleApiResult(result, (value) => ({ assignment: value }));
}

async function handleDownloadAssignmentFile(
  input: z.infer<typeof DownloadAssignmentFileInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await downloadDriveItem(input.fileUrl, input.outputPath);
  return handleApiResult(result, (value) => ({ ...value }));
}

async function handleSubmissionAction(
  input: z.infer<typeof SubmissionActionInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await actOnSubmission(
    input.action,
    input.classId,
    input.assignmentId,
    input.submissionId,
  );
  return handleApiResult(result, (value) => ({
    action: value.action,
    submissionId: value.submissionId,
    status: value.status,
    succeeded: true,
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const listAssignmentsTool: RegisteredTool<typeof ListAssignmentsInputSchema> = {
  definition: listAssignmentsToolDefinition,
  schema: ListAssignmentsInputSchema,
  handler: handleListAssignments,
};

export const getAssignmentTool: RegisteredTool<typeof GetAssignmentInputSchema> = {
  definition: getAssignmentToolDefinition,
  schema: GetAssignmentInputSchema,
  handler: handleGetAssignment,
};

export const downloadAssignmentFileTool: RegisteredTool<typeof DownloadAssignmentFileInputSchema> = {
  definition: downloadAssignmentFileToolDefinition,
  schema: DownloadAssignmentFileInputSchema,
  handler: handleDownloadAssignmentFile,
};

export const submissionActionTool: RegisteredTool<typeof SubmissionActionInputSchema> = {
  definition: submissionActionToolDefinition,
  schema: SubmissionActionInputSchema,
  handler: handleSubmissionAction,
};

/** All assignment-related tools. */
export const assignmentTools = [
  listAssignmentsTool,
  getAssignmentTool,
  downloadAssignmentFileTool,
  submissionActionTool,
];
