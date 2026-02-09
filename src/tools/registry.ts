/**
 * Tool registry - centralised tool management.
 * 
 * All tools are registered here and can be looked up by name.
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';

import { searchTools } from './search-tools.js';
import { messageTools } from './message-tools.js';
import { peopleTools } from './people-tools.js';
import { authTools } from './auth-tools.js';
import { meetingTools } from './meeting-tools.js';

// ─────────────────────────────────────────────────────────────────────────────
// Registry
// ─────────────────────────────────────────────────────────────────────────────

/** All registered tools (cast to base type for registry). */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const allTools: RegisteredTool<any>[] = [
  ...searchTools,
  ...messageTools,
  ...peopleTools,
  ...authTools,
  ...meetingTools,
];

/** Lookup map for tools by name. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const toolsByName = new Map<string, RegisteredTool<any>>(
  allTools.map(tool => [tool.definition.name, tool])
);

// ─────────────────────────────────────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets all tool definitions for MCP ListTools.
 */
export function getToolDefinitions(): Tool[] {
  return allTools.map(tool => tool.definition);
}

/**
 * Gets a tool by name.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function getTool(name: string): RegisteredTool<any> | undefined {
  return toolsByName.get(name);
}

/**
 * Invokes a tool by name with the given arguments and context.
 */
export async function invokeTool(
  name: string,
  args: unknown,
  ctx: ToolContext
): Promise<ToolResult> {
  const tool = toolsByName.get(name);
  
  if (!tool) {
    throw new Error(`Unknown tool: ${name}`);
  }

  // Validate input
  const parseResult = tool.schema.safeParse(args);
  if (!parseResult.success) {
    throw new Error(`Invalid input: ${parseResult.error.message}`);
  }

  // Invoke handler
  return tool.handler(parseResult.data, ctx);
}

/**
 * Checks if a tool exists.
 */
export function hasTool(name: string): boolean {
  return toolsByName.has(name);
}
