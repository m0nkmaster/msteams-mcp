/**
 * Tool handler registry.
 * 
 * Provides a modular way to define MCP tool handlers without a monolithic
 * switch statement. Each tool is defined with its schema, handler, and metadata.
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { z } from 'zod';
import type { McpError } from '../types/errors.js';

// Import and re-export TeamsServer from the shared types module
import type { TeamsServer } from '../types/server.js';
export type { TeamsServer };

/** The context passed to tool handlers. */
export interface ToolContext {
  /** Reference to the server for browser operations. */
  server: TeamsServer;
}

/** Result returned by tool handlers. */
export type ToolResult = 
  | { success: true; data: Record<string, unknown> }
  | { success: false; error: McpError };

// Import Result type for the helper function
import type { Result } from '../types/result.js';

/**
 * Helper to convert an API Result to a ToolResult.
 * 
 * Reduces boilerplate in tool handlers by standardising the pattern:
 * ```typescript
 * if (!result.ok) {
 *   return { success: false, error: result.error };
 * }
 * return { success: true, data: transform(result.value) };
 * ```
 * 
 * @param result - The API result to convert
 * @param transform - Function to transform the success value to response data
 * @returns A ToolResult suitable for returning from a handler
 */
export function handleApiResult<T>(
  result: Result<T>,
  transform: (value: T) => Record<string, unknown>
): ToolResult {
  if (!result.ok) {
    return { success: false, error: result.error };
  }
  return { success: true, data: transform(result.value) };
}

/** A registered tool with its handler. */
export interface RegisteredTool<TInput extends z.ZodType = z.ZodType> {
  /** Tool definition for MCP. */
  definition: Tool;
  /** Zod schema for input validation. */
  schema: TInput;
  /** Handler function. */
  handler: (input: z.infer<TInput>, ctx: ToolContext) => Promise<ToolResult>;
}

// Re-export tool registrations
export * from './search-tools.js';
export * from './message-tools.js';
export * from './people-tools.js';
export * from './auth-tools.js';
export * from './meeting-tools.js';
export * from './file-tools.js';
export * from './graph-tools.js';
export * from './registry.js';
