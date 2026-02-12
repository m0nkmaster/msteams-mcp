/**
 * MCP Server implementation for Microsoft Teams.
 * Exposes tools and resources for searching, messaging, calendar, files, and more.
 */

import { createRequire } from 'module';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';

const require = createRequire(import.meta.url);
const pkg = require('../package.json') as { version: string };
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  ListResourcesRequestSchema,
  ReadResourceRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

import { createBrowserContext, closeBrowser, type BrowserManager } from './browser/context.js';
import { ensureAuthenticated } from './browser/auth.js';

// Auth modules
import {
  hasSessionState,
  isSessionLikelyExpired,
} from './auth/session-store.js';
import {
  getSubstrateTokenStatus,
  extractMessageAuth,
  extractCsaToken,
  getUserProfile,
  clearTokenCache,
} from './auth/token-extractor.js';
import { refreshTokensViaBrowser } from './auth/token-refresh.js';

// API modules
import { getFavorites } from './api/csa-api.js';

// Tool registry
import { getToolDefinitions, invokeTool } from './tools/registry.js';
import type { ToolContext } from './tools/index.js';

// Types
import { ErrorCode, createError, type McpError } from './types/errors.js';
import type { TeamsServer as ITeamsServer } from './types/server.js';

// ─────────────────────────────────────────────────────────────────────────────
// MCP Server Class
// ─────────────────────────────────────────────────────────────────────────────

/**
 * MCP Server for Teams integration.
 * 
 * Encapsulates all server state to allow multiple instances.
 * Implements ITeamsServer interface for use by tool handlers.
 */
export class TeamsServer implements ITeamsServer {
  private browserManager: BrowserManager | null = null;
  private isInitialised = false;

  // ───────────────────────────────────────────────────────────────────────────
  // Response Formatting
  // ───────────────────────────────────────────────────────────────────────────

  /**
   * Returns a standard MCP error response.
   */
  private formatError(error: McpError) {
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify({
            success: false,
            error: error.message,
            errorCode: error.code,
            retryable: error.retryable,
            retryAfterMs: error.retryAfterMs,
            suggestions: error.suggestions,
          }, null, 2),
        },
      ],
      isError: true,
    };
  }

  /**
   * Returns a standard MCP success response.
   */
  private formatSuccess(data: Record<string, unknown>) {
    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify({ success: true, ...data }, null, 2),
        },
      ],
    };
  }

  // ───────────────────────────────────────────────────────────────────────────
  // Browser State Management (exposed for tool handlers)
  // ───────────────────────────────────────────────────────────────────────────

  /**
   * Gets the current browser manager.
   */
  getBrowserManager(): BrowserManager | null {
    return this.browserManager;
  }

  /**
   * Sets the browser manager.
   */
  setBrowserManager(manager: BrowserManager): void {
    this.browserManager = manager;
  }

  /**
   * Resets browser state.
   */
  resetBrowserState(): void {
    this.browserManager = null;
    this.isInitialised = false;
  }

  /**
   * Marks the server as initialised.
   */
  markInitialised(): void {
    this.isInitialised = true;
  }

  /**
   * Checks if the server is initialised.
   */
  isInitialisedState(): boolean {
    return this.isInitialised;
  }

  /**
   * Ensures the browser is running and authenticated.
   */
  async ensureBrowser(headless: boolean = true): Promise<BrowserManager> {
    if (this.browserManager && this.isInitialised) {
      return this.browserManager;
    }

    // Close existing browser if any
    if (this.browserManager) {
      try {
        await closeBrowser(this.browserManager, true);
      } catch {
        // Ignore cleanup errors
      }
    }

    this.browserManager = await createBrowserContext({ headless });

    await ensureAuthenticated(
      this.browserManager.page,
      this.browserManager.context,
      (msg) => console.error(`[auth] ${msg}`)
    );

    this.isInitialised = true;
    return this.browserManager;
  }

  /**
   * Cleans up browser resources.
   */
  async cleanup(): Promise<void> {
    if (this.browserManager) {
      await closeBrowser(this.browserManager, true);
      this.browserManager = null;
      this.isInitialised = false;
    }
  }

  // ───────────────────────────────────────────────────────────────────────────
  // Auto-Login on Auth Failure
  // ───────────────────────────────────────────────────────────────────────────

  /** Auth tool names that should not trigger auto-login retry. */
  private static readonly AUTH_TOOL_NAMES = new Set(['teams_login', 'teams_status']);

  /**
   * Checks if an error is an authentication error that could be resolved by logging in.
   */
  private isAuthError(error: McpError): boolean {
    return error.code === ErrorCode.AUTH_REQUIRED || error.code === ErrorCode.AUTH_EXPIRED;
  }

  /**
   * Checks if a tool is an auth-related tool (login/status) that shouldn't trigger auto-login.
   */
  private isAuthTool(name: string): boolean {
    return TeamsServer.AUTH_TOOL_NAMES.has(name);
  }

  /**
   * Attempts automatic re-authentication via headless browser.
   * Returns true if login succeeded and tokens are now available.
   */
  private async attemptAutoLogin(): Promise<boolean> {
    try {
      // First try the lightweight token refresh (headless browser, persistent profile)
      const refreshResult = await refreshTokensViaBrowser();
      if (refreshResult.ok) {
        this.markInitialised();
        return true;
      }

      // Token refresh failed — try a full headless login
      // (covers cases where session cookies are still valid but token cache is stale)
      console.error('[auto-login] Token refresh failed, trying full headless login...');
      clearTokenCache();

      const headlessManager = await createBrowserContext({ headless: true });
      try {
        await ensureAuthenticated(
          headlessManager.page,
          headlessManager.context,
          (msg) => console.error(`[auto-login:headless] ${msg}`),
          false, // No overlay in headless
          true   // Headless mode — throw if user interaction required
        );

        await closeBrowser(headlessManager, true);
        this.resetBrowserState();
        this.markInitialised();
        return true;
      } catch {
        // Headless login also failed — user interaction required
        try {
          await closeBrowser(headlessManager, false);
        } catch {
          // Ignore cleanup errors
        }
        this.resetBrowserState();
        return false;
      }
    } catch (error) {
      console.error(`[auto-login] Auto-login failed: ${error instanceof Error ? error.message : String(error)}`);
      return false;
    }
  }

  // ───────────────────────────────────────────────────────────────────────────
  // Server Creation
  // ───────────────────────────────────────────────────────────────────────────

  /**
   * Creates and configures the MCP server.
   */
  async createServer(): Promise<Server> {
    const server = new Server(
      {
        name: 'teams-mcp',
        version: pkg.version,
      },
      {
        capabilities: {
          tools: {},
          resources: {},
        },
      }
    );

    // Handle resource listing
    server.setRequestHandler(ListResourcesRequestSchema, async () => {
      return {
        resources: [
          {
            uri: 'teams://me/profile',
            name: 'Current User Profile',
            description: 'The authenticated user\'s Teams profile including email and display name',
            mimeType: 'application/json',
          },
          {
            uri: 'teams://me/favorites',
            name: 'Pinned Conversations',
            description: 'The user\'s favourite/pinned Teams conversations',
            mimeType: 'application/json',
          },
          {
            uri: 'teams://status',
            name: 'Authentication Status',
            description: 'Current authentication status for all Teams APIs',
            mimeType: 'application/json',
          },
        ],
      };
    });

    // Handle resource reading
    server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
      const { uri } = request.params;

      switch (uri) {
        case 'teams://me/profile': {
          const profile = getUserProfile();
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(profile ?? { error: 'No valid session' }, null, 2),
              },
            ],
          };
        }

        case 'teams://me/favorites': {
          const result = await getFavorites();
          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(
                  result.ok ? result.value.favorites : { error: result.error.message },
                  null,
                  2
                ),
              },
            ],
          };
        }

        case 'teams://status': {
          const tokenStatus = getSubstrateTokenStatus();
          const messageAuth = extractMessageAuth();
          const csaToken = extractCsaToken();

          const status = {
            directApi: {
              available: tokenStatus.hasToken,
              expiresAt: tokenStatus.expiresAt,
              minutesRemaining: tokenStatus.minutesRemaining,
            },
            messaging: {
              available: messageAuth !== null,
            },
            favorites: {
              available: messageAuth !== null && csaToken !== null,
            },
            session: {
              exists: hasSessionState(),
              likelyExpired: isSessionLikelyExpired(),
            },
          };

          return {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(status, null, 2),
              },
            ],
          };
        }

        default:
          throw new Error(`Unknown resource: ${uri}`);
      }
    });

    // Handle tool listing
    server.setRequestHandler(ListToolsRequestSchema, async () => {
      return { tools: getToolDefinitions() };
    });

    // Handle tool calls (with auto-login retry for auth errors)
    server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        const ctx: ToolContext = { server: this };
        const result = await invokeTool(name, args, ctx);

        if (result.success) {
          return this.formatSuccess(result.data);
        }

        // Auto-login retry for auth errors (skip for login/status tools themselves)
        if (this.isAuthError(result.error) && !this.isAuthTool(name)) {
          console.error(`[auto-login] Tool '${name}' returned ${result.error.code}, attempting automatic re-authentication...`);
          const loginSuccess = await this.attemptAutoLogin();

          if (loginSuccess) {
            console.error('[auto-login] Re-authentication succeeded, retrying tool...');
            const retryResult = await invokeTool(name, args, ctx);
            if (retryResult.success) {
              return this.formatSuccess(retryResult.data);
            }
            // Retry also failed — fall through to return the new error
            return this.formatError(retryResult.error);
          }

          // Auto-login failed — return a strongly-worded error
          console.error('[auto-login] Automatic re-authentication failed.');
          return this.formatError(createError(
            ErrorCode.AUTH_REQUIRED,
            'ACTION REQUIRED: Teams session has expired and automatic re-authentication failed. You MUST call the teams_login tool now to re-authenticate before using any other Teams tools. Do NOT skip this step or tell the user Teams is unavailable.',
            { suggestions: [
              'IMMEDIATELY call teams_login to re-authenticate',
              'After login succeeds, retry the original request',
              'Do NOT tell the user that Teams is unavailable — authentication just needs refreshing',
            ] }
          ));
        }

        return this.formatError(result.error);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);

        return this.formatError(createError(
          ErrorCode.UNKNOWN,
          message,
          { retryable: false }
        ));
      }
    });

    // Cleanup on server close
    server.onclose = async () => {
      await this.cleanup();
    };

    return server;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Creates and runs the MCP server.
 * Exported for backward compatibility.
 */
export async function createServer(): Promise<Server> {
  const teamsServer = new TeamsServer();
  return teamsServer.createServer();
}

/**
 * Runs the server with stdio transport.
 */
export async function runServer(): Promise<void> {
  const teamsServer = new TeamsServer();
  const server = await teamsServer.createServer();
  const transport = new StdioServerTransport();

  await server.connect(transport);

  // Handle shutdown signals
  const shutdown = async () => {
    await teamsServer.cleanup();
    process.exit(0);
  };

  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);
}
