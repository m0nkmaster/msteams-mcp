# Agent Guidelines for Teams MCP

Non-obvious project knowledge for AI agents. This is not a tool reference: tool parameters and usage live in the tool definitions (`src/tools/*.ts`), which are the source of truth and are sent to AI assistants via MCP.

## Repository

- **Repository**: https://github.com/m0nkmaster/msteams-mcp
- **Install**: `npx -y msteams-mcp@latest`, or clone, `npm install && npm run build`, then point your MCP client at `dist/index.js`

## Project Overview

An MCP (Model Context Protocol) server that lets AI assistants interact with Microsoft Teams. Instead of the Microsoft Graph API, it calls Teams' own APIs (Substrate, chatsvc, CSA) using tokens extracted from a browser session. The browser is only used for initial login; all operations are direct API calls.

## Directory Structure

```
src/
├── index.ts              # Entry point, runs the MCP server
├── cli.ts                # Standalone CLI (msteams bin) - full tool parity via in-memory MCP transport
├── server.ts             # MCP server (TeamsServer class) - delegates to tool registry
├── constants.ts          # Shared constants (page sizes, timeouts, thresholds)
├── tools/                # Tool handlers (modular design)
│   ├── index.ts          # Tool context and type definitions
│   ├── registry.ts       # Tool registry - maps names to handlers
│   ├── search-tools.ts   # Search and channel tools
│   ├── message-tools.ts  # Messaging, favourites, save/unsave tools
│   ├── people-tools.ts   # People search and profile tools
│   ├── meeting-tools.ts  # Calendar and meeting tools
│   ├── file-tools.ts     # Shared files tools
│   └── auth-tools.ts     # Login and status tools
├── auth/                 # Authentication and credential management
│   ├── index.ts          # Module exports
│   ├── crypto.ts         # AES-256-GCM encryption for credentials at rest
│   ├── session-store.ts  # Secure session state storage with encryption
│   ├── token-extractor.ts # Extract tokens from Playwright session state
│   ├── token-refresh.ts  # Token refresh orchestrator (HTTP-first, browser fallback)
│   └── token-refresh-http.ts # Browserless token refresh via direct OAuth2 calls
├── api/                  # API client modules (one per API surface)
│   ├── index.ts          # Module exports
│   ├── substrate-api.ts  # Search and people APIs (Substrate v2)
│   ├── chatsvc-api.ts    # Barrel file re-exporting all chatsvc sub-modules
│   ├── chatsvc-common.ts # Shared utilities (date formatting)
│   ├── chatsvc-messaging.ts # Send, edit, delete, threads, 1:1/group chat
│   ├── chatsvc-activity.ts  # Activity feed (mentions, reactions, replies)
│   ├── chatsvc-reactions.ts # Add/remove emoji reactions
│   ├── chatsvc-virtual.ts   # Saved messages, followed threads, save/unsave
│   ├── chatsvc-readstatus.ts # Consumption horizons, mark as read, unread
│   ├── csa-api.ts        # Favorites API (CSA)
│   ├── calendar-api.ts   # Calendar/meetings API
│   ├── transcript-api.ts # Meeting transcripts (Substrate WorkingSetFiles)
│   ├── files-api.ts      # Shared files (Substrate AllFiles)
│   └── profile-api.ts    # Resolve MRIs to profiles (middleTier fetchShortProfile)
├── browser/              # Playwright browser automation (login only)
│   ├── context.ts        # Persistent browser profile management
│   └── auth.ts           # Authentication detection and manual login handling
├── utils/
│   ├── parsers.ts        # Pure parsing functions (barrel; testable submodules)
│   ├── parsers-reactions.ts # Emoji reaction parsing from raw messages
│   ├── parsers.test.ts   # Unit tests for parsers
│   ├── http.ts           # HTTP client with retry, timeout, error handling
│   ├── api-config.ts     # API endpoints and header configuration
│   └── auth-guards.ts    # Reusable auth check utilities (Result types)
├── types/
│   ├── teams.ts          # Teams data interfaces
│   ├── errors.ts         # Error taxonomy with machine-readable codes
│   ├── result.ts         # Result<T, E> type for explicit error handling
│   └── api-responses.ts  # Typed interfaces for raw API response shapes
├── __fixtures__/
│   └── api-responses.ts  # Mock API responses for testing
```

## Implementation Patterns

- **Result types**: API functions return `Result<T, McpError>` for explicit success/failure handling.
- **Error taxonomy**: Errors use machine-readable codes (`ErrorCode` enum), `retryable` flags, and `suggestions` arrays so LLMs can understand failures and recover.
- **HTTP utilities**: A centralised client (`utils/http.ts`) provides retry with exponential backoff, timeouts, and rate-limit tracking. Use `httpRequest()` for all new API calls.
- **Server class**: `TeamsServer` encapsulates all state (browser manager, init flag), allowing multiple instances and simpler testing.
- **Tool registry**: Tools are grouped by category (`search-tools.ts`, `message-tools.ts`, etc.) and wired through `tools/registry.ts`.
- **Auth guards**: `utils/auth-guards.ts` provides reusable, `Result`-returning auth checks plus cached `getTenantId()`, `getRegion()`, and `getTeamsBaseUrl()` helpers.
- **Shared constants**: Magic numbers (page sizes, timeouts, thresholds) live in `constants.ts`.
- **MCP resources**: Passive resources (`teams://me/profile`, `teams://me/favorites`, `teams://status`) provide context discovery without tool calls.
- **Markdown to Teams HTML**: `markdownToTeamsHtml()` in `utils/parsers.ts` converts markdown (bold, italic, code, code blocks, strikethrough, lists, newlines) to the `RichText/Html` Teams expects. Used by `sendMessage()` and `editMessage()`. When messages contain @mentions or links, `parseContentWithMentionsAndLinks()` applies the same conversion to text between inline elements.
- **Auto-login on auth failure**: The `CallToolRequestSchema` handler in `server.ts` retries tool calls that fail with `AUTH_REQUIRED`/`AUTH_EXPIRED`, first attempting headless re-auth (token refresh, then full headless login). Auth tools (`teams_login`, `teams_status`) are excluded to avoid loops. Concurrent failures are deduplicated via a Promise-based mutex so only one auto-login runs at a time.

### Dynamic Configuration from Session

All tenant-specific config is extracted from the user's session localStorage, so the server works across Teams environments (commercial, GCC, GCC-High, DoD):

- **Region & partition**: From `DISCOVER-REGION-GTM` (e.g. region `amer`, partition `02`), via cached `getRegion()`.
- **Teams base URL**: From the `chatServiceAfd` URL in `DISCOVER-REGION-GTM` (e.g. `https://teams.microsoft.com`, or `https://teams.microsoft.us` for government clouds), via cached `getTeamsBaseUrl()`.
- **User details**: From `DISCOVER-USER-DETAILS`, including user MRI, licence info, and user/tenant partitions.
- **Service URLs**: Full chatsvc, CSA, and mt/part URLs are in the config and passed to endpoint builders.

**Note**: The Substrate search URL (`substrate.office.com`) is hardcoded, as no config source has been found for it. This may need to become configurable if GCC users report issues.

## Authentication

### Login flow

All operations use direct API calls. A persistent browser profile (`~/.teams-mcp-server/browser-profile/`) stores Microsoft session cookies and MSAL tokens, enabling silent re-authentication.

1. **First login**: Visible browser opens, user authenticates, session state is saved, browser closes.
2. **Token expiry**: Headless browser (persistent profile) refreshes tokens via silent SSO, no interaction.
3. **Session fully expired**: Falls back to a visible browser for manual re-login (with extensions and saved form data available).
4. **All API operations**: Use cached tokens, no browser.

`teams_login` always tries headless SSO before showing a visible browser. Long-lived Microsoft session cookies (days/weeks) mean users rarely re-authenticate manually, even though MSAL tokens expire after ~1 hour.

The server uses the system browser via Playwright's `launchPersistentContext()` (Edge on Windows, Chrome on macOS/Linux; ~180MB saved vs bundled Chromium). Only one process can use the profile at a time (Chromium lock); the token-refresh module uses a module-level flag to prevent concurrent access. If no system browser is found, the error suggests installing Chrome or running `npx playwright install chromium`.

### Token refresh

Refresh uses an HTTP-first strategy in `token-refresh-http.ts`:

1. **HTTP (~1s)**: Extract the MSAL refresh token from session state and POST to Azure AD's OAuth2 token endpoint for each required scope (Substrate, Skype Spaces, chatsvcagg). Exchange the Skype Spaces token for the `skypetoken_asm` cookie via `authsvc.teams.microsoft.com/v1.0/authz`. Write updated tokens back to session state in MSAL cache format so `token-extractor.ts` finds them. The `Origin: https://teams.microsoft.com` header is required (the Teams client ID is a SPA; without it Azure AD returns AADSTS9002327).
2. **Browser fallback (~8s)**: If HTTP fails (e.g. refresh token expired, Conditional Access), a headless browser uses the persistent profile's session cookies for silent SSO.

Both are seamless (no window, no interaction). If both fail, the user must re-authenticate via `teams_login`. First login always needs a browser (no refresh token yet). Works identically for standard MS login and corporate SSO (ADFS/Okta federation). Test with `npm run cli -- login`.

### Auth per API

Different Teams APIs use different auth mechanisms:

| API | Auth Method | Helper Function |
|-----|-------------|-----------------|
| **Search / Email / People** (Substrate) | JWT Bearer from MSAL (People also needs `cvid`/`logicalId`) | `getValidSubstrateToken()` |
| **Messaging / Threads** (chatsvc) | `skypetoken_asm` cookie | `extractMessageAuth()` |
| **Favorites** (csa/conversationFolders) | CSA token from MSAL + `skypetoken_asm` | `extractCsaToken()` + `extractMessageAuth()` |
| **Calendar** (mt/part/calendarView) | Skype Spaces token + `skypetoken_asm` | `extractSkypeSpacesToken()` |
| **Transcripts** (Substrate WorkingSetFiles) | Substrate JWT + `Prefer` header | `getValidSubstrateToken()` |
| **Files** (Substrate AllFiles) | Substrate JWT + message auth for user MRI | `getValidSubstrateToken()` + `extractMessageAuth()` |
| **Profiles** (mt/part fetchShortProfile) | Skype Spaces token + `skypetoken_asm` | `requireSkypeSpacesAuthWithConfig()` |

All helpers live in `auth/token-extractor`. Notes:
- The CSA API (favorites) needs GET to read; POST is only for modifications.
- The Substrate suggestions API requires `cvid` and `logicalId` correlation IDs in the body.
- Regional APIs (chatsvc, csa, mt/part) resolve their region via `getRegion()`; partitioned endpoints (mt/part Calendar) also use the partition suffix from `DISCOVER-REGION-GTM`.

### Session persistence and encryption

Two layers work together:

1. **Persistent browser profile** (`browser-profile/`): retains Microsoft session cookies, extensions, and autofill across launches; enables silent headless re-auth.
2. **Encrypted session state** (`session-state.json`): Playwright `storageState()` output from which tokens are extracted for browserless API calls.

Session state and token cache files are encrypted at rest with AES-256-GCM using a key derived from machine-specific values (hostname + username), stored with 0o600 permissions. Existing plaintext files are migrated to encrypted form on first read.

## MCP Tools

Tool definitions in `src/tools/*.ts` are the source of truth. Their `description` fields are sent to AI assistants via MCP, so keep them comprehensive: what the tool does, key parameters, common gotchas, and related tools. Do not maintain a duplicate tool list here.

**Design philosophy**: a minimal toolset - fewer, more powerful, composable tools rather than convenience wrappers. The AI builds queries using search operators.

For manual testing of all tools, see `docs/MANUAL-TEST-SCRIPT.md`.

## Development

```bash
npm run dev           # Run MCP server in development mode
npm run build         # Compile TypeScript
npm run lint          # Run ESLint (lint:fix to auto-fix)
npm start             # Run compiled MCP server
npm test              # Run unit tests
npm run test:watch    # Unit tests in watch mode
npm run test:coverage # Coverage report
npm run typecheck     # Type checking only
```

CI (GitHub Actions) runs lint, typecheck, tests, and build on every push and PR, plus a debounced doc-accuracy review on main commits. See `.github/workflows/`.

### CLI as test harness

`src/cli.ts` is a first-class, installable CLI (`msteams` bin) with full tool parity to the MCP server. It runs the server in-process over the SDK's `InMemoryTransport`, so every call exercises the real MCP protocol layer (tool definitions, input validation, response formatting). This makes it both the integration-test harness and the way to test against live Teams APIs. Use `npm run cli` from a clone (via `tsx`, no build needed) or the installed binary.

The CLI can call any tool generically; unrecognised commands are treated as tool names (auto-prefixed with `teams_`). Pass parameters as `--key value`.

```bash
npm run cli                                     # List tools and shortcuts
npm run cli -- find_channel --query "support"   # Generic tool call
npm run cli -- search "your query"              # Shortcut for teams_search
npm run cli -- login                            # Headless SSO first
npm run cli -- login --force true               # Clear session and re-login
npm run cli -- search "your query" --json       # Raw MCP JSON response
npm run cli -- search "your query" --from 25 --size 25  # Pagination
```

### Unit tests

Vitest, focused on pure functions and outcomes over implementation. Fixtures (`src/__fixtures__/api-responses.ts`) mirror real API shapes. Tested functions live in `src/utils/parsers.ts` and cover HTML stripping/entity decoding, deep-link generation, timestamp extraction, search/people/email result parsing, JWT profile extraction, token-status calculation, base64 GUID decoding, and user-ID extraction across formats.

### Extending

**New tool**: pick the appropriate `src/tools/*.ts` file (or add a category), define a Zod input schema and the MCP tool definition, implement the handler returning `ToolResult`, export it into the module's `*Tools` array, and register a new category in `tools/registry.ts`. Back it with `Result<T, McpError>` API functions.

**New API endpoint**: add the URL to `utils/api-config.ts`, add a function in the relevant `src/api/*.ts` using `httpRequest()`, and return `Result<T, McpError>`.

### Roadmap

When a roadmap item is completed, remove it entirely from `ROADMAP.md` - do not cross it out or mark it done.

## Troubleshooting

**Session/token expired**: call `teams_login` with `forceNew: true`, or delete the config directory and run `npm run cli -- login`.

**Browser won't launch (login)**: ensure Chrome (macOS/Linux) or Edge (Windows) is installed; as a fallback run `npx playwright install chromium`; check for browser processes holding the profile lock.

**Login timeout with MFA**: MCP clients have request timeouts (2-5 min). If SSO/MFA takes longer, the request times out but login continues in the background - completing MFA in the browser saves the session and later calls work. This is rare because headless SSO is tried first and a visible browser only opens when credentials are actually required.

**Search misses thread replies**: Substrate search is full-text and only returns messages matching the terms. A reply that does not contain the search keywords will not appear even if it is a direct reply. Workaround: after finding a message, use `teams_get_thread` with its `conversationId` for full context.

## Reference

### File locations

Session files live in a per-user config directory (consistent regardless of invocation method):

- **macOS/Linux**: `~/.teams-mcp-server/`
- **Windows**: `%APPDATA%\teams-mcp-server\`

Contents: `session-state.json` (encrypted session), `token-cache.json` (encrypted tokens), `browser-profile/` (persistent Chrome/Edge profile). Legacy files in the project root are migrated on first read. Dev-only: `./debug-output/` (gitignored screenshots/HTML). Reference docs: `docs/API-REFERENCE.md`, `docs/SESSION-DATA-REFERENCE.md`.

### Message deep links

Teams needs different deep-link formats per conversation type:

| Conversation Type | Format |
|-------------------|--------|
| **Channel (top-level)** | `/l/message/{channelId}/{msgTimestamp}` |
| **Channel (thread reply)** | `/l/message/{channelId}/{msgTimestamp}?parentMessageId={parentId}` |
| **1:1 / Group chat** | `/l/message/{chatId}/{msgTimestamp}?context={"contextType":"chat"}` |
| **Meeting chat** | `/l/message/{meetingId}/{msgTimestamp}?context={"contextType":"chat"}` |

Conversation ID patterns: channels `19:xxx@thread.tacv2`; meetings `19:meeting_xxx@thread.v2`; 1:1 `19:guid_guid@unq.gbl.spaces` (GUIDs sorted lexicographically); group chats `19:xxx@thread.v2`. To detect a thread reply, compare the `messageid` in `ClientConversationId` with the message's own `DateTimeReceived` timestamp; if they differ it needs `parentMessageId`.

### API internals

- **Conversation types**: chatsvc returns `threadType` (`topic`, `space`, `meeting`, `chat`) and `productThreadType` (`TeamsStandardChannel`, `TeamsTeam`, `TeamsPrivateChannel`, `Meeting`, `Chat`, `OneOnOne`). See `docs/API-REFERENCE.md`.
- **Virtual conversations**: IDs like `48:saved`, `48:threads`, `48:mentions`, `48:notifications`, `48:notes` aggregate across conversations; messages include `clumpId` for the source.
- **User ID formats**: `extractObjectId()` in `parsers.ts` handles raw GUIDs, MRIs (`8:orgid:...`), tenant-suffixed IDs, base64 GUIDs (little-endian), and Skype IDs.
- **Deleted messages**: returned with empty `content` and a `deletetime` property; filtered out in `getThreadMessages()` to avoid phantom "empty" messages.
- **Message ordering**: chatsvc returns newest-first by default; ascending when `startTime` is given. `getThreadMessages()` defaults to `order: 'desc'` but accepts `'asc'` for chronological reading.

### Known limitations

- **Presence/status**: real-time via WebSocket, not available over HTTP.

## Dependencies

- `@modelcontextprotocol/sdk`: MCP protocol implementation
- `playwright`: browser automation (login only)
- `zod`: runtime input validation
- `vitest`: unit testing (dev)
