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
import { refreshTokensViaBrowser } from './auth/token-refresh.js';

// API modules
import { getFavorites } from './api/csa-api.js';

// Tool registry
import { getToolDefinitions, invokeTool } from './tools/registry.js';

// Types
import { ErrorCode, createError, type McpError } from './types/errors.js';
import * as log from './utils/logger.js';

// ─────────────────────────────────────────────────────────────────────────────
// MCP Server Class
// ─────────────────────────────────────────────────────────────────────────────

/**
 * MCP Server for Teams integration.
 */
export class TeamsServer {
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
  // Auto-Login on Auth Failure
  // ───────────────────────────────────────────────────────────────────────────

  /** Auth tool names that should not trigger auto-login retry. */
  private static readonly AUTH_TOOL_NAMES = new Set(['teams_login', 'teams_status']);

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
        const result = await invokeTool(name, args);

        if (result.success) {
          return this.formatSuccess(result.data);
        }

        // Auto-login retry for auth errors (skip for login/status tools themselves)
        const isAuthError = result.error.code === ErrorCode.AUTH_REQUIRED || result.error.code === ErrorCode.AUTH_EXPIRED;
        if (isAuthError && !TeamsServer.AUTH_TOOL_NAMES.has(name)) {
          log.warn('auto-login', `Tool '${name}' returned ${result.error.code}, attempting automatic re-authentication...`);
          // HTTP refresh, then headless browser SSO; concurrent calls share one attempt
          const refresh = await refreshTokensViaBrowser();

          if (refresh.ok) {
            log.info('auto-login', 'Re-authentication succeeded, retrying tool...');
            const retryResult = await invokeTool(name, args);
            if (retryResult.success) {
              return this.formatSuccess(retryResult.data);
            }
            // Retry also failed after successful re-auth — log for diagnostics
            log.error('auto-login', `Retry failed after successful re-authentication: ${retryResult.error.message}`);
            return this.formatError(retryResult.error);
          }

          // Auto-login failed — return a strongly-worded error directing to forceNew
          log.error('auto-login', 'Automatic re-authentication failed.');
          return this.formatError(createError(
            ErrorCode.AUTH_REQUIRED,
            'ACTION REQUIRED: Teams session has expired and automatic re-authentication failed. You MUST call teams_login with forceNew: true to clear stale state and re-authenticate. Do NOT call teams_login without forceNew — it will return early without fixing the problem.',
            { suggestions: [
              'IMMEDIATELY call teams_login with { "forceNew": true }',
              'After login succeeds, retry the original request',
              'Do NOT skip this step or tell the user Teams is unavailable — authentication just needs refreshing',
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

  process.on('SIGINT', () => process.exit(0));
  process.on('SIGTERM', () => process.exit(0));
}
