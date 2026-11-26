/**
 * Unit tests for TeamsServer tool-call handling, especially the
 * auto-re-authentication retry path.
 */

import { beforeEach, describe, expect, it, vi } from 'vitest';
import { ErrorCode, createError } from './types/errors.js';
import { recommendsForceNewWithoutApproval } from './__fixtures__/force-new-guard.js';

vi.mock('./tools/registry.js', () => ({
  invokeTool: vi.fn(),
  getToolDefinitions: vi.fn(() => []),
}));

vi.mock('./auth/token-refresh.js', () => ({
  refreshTokensViaBrowser: vi.fn(),
}));

vi.mock('./api/csa-api.js', () => ({
  getFavorites: vi.fn(),
}));

import { TeamsServer } from './server.js';
import { invokeTool } from './tools/registry.js';
import { refreshTokensViaBrowser } from './auth/token-refresh.js';

type ToolCallResponse = {
  content: Array<{ type: string; text: string }>;
  isError?: boolean;
};

function callTool(name: string, args: unknown = {}): Promise<ToolCallResponse> {
  const server = new TeamsServer() as unknown as {
    handleToolCall: (name: string, args: unknown) => Promise<ToolCallResponse>;
  };
  return server.handleToolCall(name, args);
}

describe('TeamsServer auto-login failure', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('directs to a plain teams_login when automatic re-authentication fails', async () => {
    vi.mocked(invokeTool).mockResolvedValue({
      success: false,
      error: createError(ErrorCode.AUTH_EXPIRED, 'expired'),
    });
    vi.mocked(refreshTokensViaBrowser).mockResolvedValue({
      ok: false,
      error: createError(ErrorCode.AUTH_REQUIRED, 'refresh failed'),
    });

    const response = await callTool('teams_search');
    const body = JSON.parse(response.content[0].text) as {
      success: boolean;
      error: string;
      errorCode: string;
      suggestions: string[];
    };

    expect(response.isError).toBe(true);
    expect(body.errorCode).toBe('AUTH_REQUIRED');
    expect(recommendsForceNewWithoutApproval(body.error)).toBe(false);
    for (const suggestion of body.suggestions) {
      expect(recommendsForceNewWithoutApproval(suggestion)).toBe(false);
    }
    expect(body.error).not.toContain('Do NOT call teams_login without forceNew');
    expect(body.suggestions).toContain(
      'Call teams_login (without forceNew) and tell the user a Teams sign-in window may open'
    );
  });

  it('returns success when re-authentication and the retry both succeed', async () => {
    vi.mocked(invokeTool)
      .mockResolvedValueOnce({
        success: false,
        error: createError(ErrorCode.AUTH_EXPIRED, 'expired'),
      })
      .mockResolvedValueOnce({ success: true, data: { results: [] } });
    vi.mocked(refreshTokensViaBrowser).mockResolvedValue({ ok: true, value: undefined });

    const response = await callTool('teams_search');
    const body = JSON.parse(response.content[0].text) as { success: boolean };

    expect(response.isError).toBeUndefined();
    expect(body.success).toBe(true);
    expect(invokeTool).toHaveBeenCalledTimes(2);
  });
});
