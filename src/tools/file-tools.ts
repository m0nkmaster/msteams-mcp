/**
 * File-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { downloadFile } from '../api/download-api.js';
import { getSharedFiles } from '../api/files-api.js';
import {
  DEFAULT_FILES_PAGE_SIZE,
  MAX_FILES_PAGE_SIZE,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const GetSharedFilesInputSchema = z.object({
  conversationId: z.string().min(1),
  pageSize: z.number().min(1).max(MAX_FILES_PAGE_SIZE).optional().default(DEFAULT_FILES_PAGE_SIZE),
  skipToken: z.string().optional(),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const getSharedFilesToolDefinition: Tool = {
  name: 'teams_get_shared_files',
  description: 'Get files and links shared in a Teams conversation. Returns file names, URLs, extensions, sizes, and who shared them. Works for channels, group chats, 1:1 chats, and meeting chats. Use the conversationId from other tools (teams_get_favorites, teams_search, teams_find_channel, teams_get_chat). Supports pagination via skipToken for conversations with many files. Pass a File item webUrl to teams_download_file to download its contents.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to get shared files for (e.g., "19:abc@thread.tacv2" for a channel, or a chat conversation ID).',
      },
      pageSize: {
        type: 'number',
        description: `Number of items per page (default: ${DEFAULT_FILES_PAGE_SIZE}, max: ${MAX_FILES_PAGE_SIZE})`,
      },
      skipToken: {
        type: 'string',
        description: 'Continuation token from a previous response to get the next page of results.',
      },
    },
    required: ['conversationId'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleGetSharedFiles(
  input: z.infer<typeof GetSharedFilesInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getSharedFiles(input.conversationId, {
    pageSize: input.pageSize,
    skipToken: input.skipToken,
  });

  return handleApiResult(result, (value) => ({
    conversationId: value.conversationId,
    returned: value.returned,
    files: value.files,
    ...(value.skipToken ? { skipToken: value.skipToken, hasMore: true } : { hasMore: false }),
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const getSharedFilesTool: RegisteredTool<typeof GetSharedFilesInputSchema> = {
  definition: getSharedFilesToolDefinition,
  schema: GetSharedFilesInputSchema,
  handler: handleGetSharedFiles,
};

export const DownloadFileInputSchema = z.object({
  url: z.string().url(),
  outputPath: z.string().min(1),
});

export const downloadFileTool: RegisteredTool<typeof DownloadFileInputSchema> = {
  definition: {
    name: 'teams_download_file',
    description: 'Download a file using its webUrl from teams_get_shared_files and the current Teams session. Supports direct SharePoint/OneDrive for Business file URLs, including chat uploads and channel files; arbitrary Link items and short sharing links are not supported. Saves raw bytes to an absolute path on the MCP server machine and returns file name, size, content type and SHA-256. Parent directory must exist; existing files are never overwritten. Streams directly to disk with no fixed file-size limit. Stalled transfers time out after 30 seconds without progress; failed transfers remove the partial file. For long downloads, increase your MCP client tool timeout.',
    inputSchema: {
      type: 'object',
      properties: {
        url: { type: 'string', description: 'The File item webUrl returned by teams_get_shared_files.' },
        outputPath: { type: 'string', description: 'Absolute destination file path on the MCP server machine; must not already exist.' },
      },
      required: ['url', 'outputPath'],
    },
  },
  schema: DownloadFileInputSchema,
  handler: async (input) => handleApiResult(await downloadFile(input.url, input.outputPath), value => ({ ...value })),
};

/** All file-related tools. */
export const fileTools = [getSharedFilesTool, downloadFileTool];
