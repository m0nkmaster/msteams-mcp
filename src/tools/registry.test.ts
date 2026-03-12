/**
 * Unit tests for tools/registry.
 */

import { describe, it, expect, vi } from 'vitest';
import { z } from 'zod';
import type { ToolContext } from './index.js';

// Mock the MCP SDK
vi.mock('@modelcontextprotocol/sdk/types.js', () => ({
  Tool: {},
}));

vi.mock('./search-tools.js', () => ({
  searchTools: [{
    definition: { name: 'teams_search_messages' },
    schema: z.object({ query: z.string() }),
    handler: vi.fn().mockResolvedValue({ success: true }),
  }],
}));

vi.mock('./message-tools.js', () => ({
  messageTools: [{
    definition: { name: 'teams_send_message' },
    schema: z.object({ conversationId: z.string(), content: z.string() }),
    handler: vi.fn().mockResolvedValue({ success: true }),
  }],
}));

vi.mock('./people-tools.js', () => ({
  peopleTools: [{
    definition: { name: 'teams_get_user' },
    schema: z.object({ userId: z.string() }),
    handler: vi.fn().mockResolvedValue({ success: true }),
  }],
}));

vi.mock('./auth-tools.js', () => ({
  authTools: [{
    definition: { name: 'teams_login' },
    schema: z.object({}),
    handler: vi.fn().mockResolvedValue({ success: true }),
  }],
}));

vi.mock('./meeting-tools.js', () => ({
  meetingTools: [{
    definition: { name: 'teams_list_meetings' },
    schema: z.object({}),
    handler: vi.fn().mockResolvedValue({ success: true }),
  }],
}));

vi.mock('./file-tools.js', () => ({
  fileTools: [{
    definition: { name: 'teams_upload_file' },
    schema: z.object({}),
    handler: vi.fn().mockResolvedValue({ success: true }),
  }],
}));

import { getToolDefinitions, getTool, invokeTool, hasTool } from './registry.js';

describe('getToolDefinitions', () => {
  it('returns an array of tool definitions', () => {
    const tools = getToolDefinitions();
    
    expect(Array.isArray(tools)).toBe(true);
    expect(tools.length).toBeGreaterThan(0);
  });

  it('returns tools with expected properties', () => {
    const tools = getToolDefinitions();
    
    for (const tool of tools) {
      expect(tool).toHaveProperty('name');
      expect(typeof tool.name).toBe('string');
    }
  });
});

describe('getTool', () => {
  it('returns a tool by name', () => {
    const tool = getTool('teams_search_messages');
    
    expect(tool).toBeDefined();
    expect(tool?.definition.name).toBe('teams_search_messages');
  });

  it('returns undefined for unknown tool', () => {
    const tool = getTool('nonexistent_tool');
    
    expect(tool).toBeUndefined();
  });
});

describe('hasTool', () => {
  it('returns true for existing tool', () => {
    expect(hasTool('teams_search_messages')).toBe(true);
  });

  it('returns false for unknown tool', () => {
    expect(hasTool('nonexistent_tool')).toBe(false);
  });
});

describe('invokeTool', () => {
  const mockContext: ToolContext = {
    server: {
      ensureBrowser: async () => ({} as never),
      resetBrowserState: () => {},
      getBrowserManager: () => null,
      setBrowserManager: () => {},
      markInitialised: () => {},
      isInitialisedState: () => false,
    },
  };

  it('invokes a known tool', async () => {
    const result = await invokeTool('teams_search_messages', { query: 'test' }, mockContext);
    
    expect(result.success).toBe(true);
  });

  it('returns error for unknown tool', async () => {
    const result = await invokeTool('nonexistent_tool', {}, mockContext);
    
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('INVALID_INPUT');
      expect(result.error.message).toContain('Unknown tool');
    }
  });

  it('validates input against schema', async () => {
    const result = await invokeTool('teams_search_messages', { query: 'test' }, mockContext);
    
    expect(result).toBeDefined();
    expect(result.success).toBe(true);
  });

  it('returns validation error for invalid input', async () => {
    // teams_search_messages requires a string query
    const result = await invokeTool('teams_search_messages', { query: 123 } as Record<string, unknown>, mockContext);
    
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('INVALID_INPUT');
      expect(result.error.message).toContain('Invalid input');
    }
  });
});
