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
} from './auth/token-extractor.js';

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

    // Handle tool calls
    server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        const ctx: ToolContext = { server: this };
        const result = await invokeTool(name, args, ctx);

        if (result.success) {
          return this.formatSuccess(result.data);
        } else {
          return this.formatError(result.error);
        }
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
