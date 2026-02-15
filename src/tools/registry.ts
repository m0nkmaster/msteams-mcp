/**
 * Tool registry - centralised tool management.
 * 
 * All tools are registered here and can be looked up by name.
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { z } from 'zod';
import type { ToolContext, ToolResult } from './index.js';

/** Type-erased tool entry for the registry — avoids generic variance issues. */
interface RegistryEntry {
  definition: Tool;
  schema: z.ZodTypeAny;
  handler: (input: z.output<z.ZodTypeAny>, ctx: ToolContext) => Promise<ToolResult>;
}

import { searchTools } from './search-tools.js';
import { messageTools } from './message-tools.js';
import { peopleTools } from './people-tools.js';
import { authTools } from './auth-tools.js';
import { meetingTools } from './meeting-tools.js';
import { fileTools } from './file-tools.js';

// ─────────────────────────────────────────────────────────────────────────────
// Registry
// ─────────────────────────────────────────────────────────────────────────────

/** All registered tools (cast to base type for registry). */
const allTools: RegistryEntry[] = [
  ...searchTools,
  ...messageTools,
  ...peopleTools,
  ...authTools,
  ...meetingTools,
  ...fileTools,
];

/** Lookup map for tools by name. */
const toolsByName = new Map<string, RegistryEntry>(
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
export function getTool(name: string): RegistryEntry | undefined {
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
