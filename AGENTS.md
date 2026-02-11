# Agent Guidelines for Teams MCP

This document captures project knowledge to help AI agents work effectively with this codebase.

## Repository

- **Repository**: https://github.com/m0nkmaster/msteams-mcp
- **Install**: `npx -y msteams-mcp@latest` or clone the repo, `npm install && npm run build`, then point your MCP client to `dist/index.js`

## Project Overview

This is an MCP (Model Context Protocol) server that enables AI assistants to interact with Microsoft Teams. Rather than using the complex Microsoft Graph API, it uses Teams APIs (Substrate, chatsvc, CSA) with authentication tokens extracted from a browser session. The browser is only used for initial login - all operations use direct API calls.

## Architecture

### Directory Structure

```
src/
├── index.ts              # Entry point, runs the MCP server
├── server.ts             # MCP server (TeamsServer class) - delegates to tool registry
├── constants.ts          # Shared constants (page sizes, timeouts, thresholds)
├── tools/                # Tool handlers (modular design)
│   ├── index.ts          # Tool context and type definitions
│   ├── registry.ts       # Tool registry - maps names to handlers
│   ├── search-tools.ts   # Search and channel tools
│   ├── message-tools.ts  # Messaging, favourites, save/unsave tools
│   ├── people-tools.ts   # People search and profile tools
│   ├── meeting-tools.ts  # Calendar and meeting tools
│   └── auth-tools.ts     # Login and status tools
├── auth/                 # Authentication and credential management
│   ├── index.ts          # Module exports
│   ├── crypto.ts         # AES-256-GCM encryption for credentials at rest
│   ├── session-store.ts  # Secure session state storage with encryption
│   ├── token-extractor.ts # Extract tokens from Playwright session state
│   └── token-refresh.ts  # Proactive token refresh via OAuth2 endpoint
├── api/                  # API client modules (one per API surface)
│   ├── index.ts          # Module exports
│   ├── substrate-api.ts  # Search and people APIs (Substrate v2)
│   ├── chatsvc-api.ts    # Messaging, threads, save/unsave (chatsvc)
│   ├── csa-api.ts        # Favorites API (CSA)
│   ├── calendar-api.ts   # Calendar/meetings API
│   └── transcript-api.ts # Meeting transcripts (Graph API)
├── browser/              # Playwright browser automation (login only)
│   ├── context.ts        # Browser/context management with encrypted session
│   └── auth.ts           # Authentication detection and manual login handling
├── utils/
│   ├── parsers.ts        # Pure parsing functions (testable)
│   ├── parsers.test.ts   # Unit tests for parsers
│   ├── http.ts           # HTTP client with retry, timeout, error handling
│   ├── api-config.ts     # API endpoints and header configuration
│   └── auth-guards.ts    # Reusable auth check utilities (Result types)
├── types/
│   ├── teams.ts          # Teams data interfaces
│   ├── errors.ts         # Error taxonomy with machine-readable codes
│   └── result.ts         # Result<T, E> type for explicit error handling
├── __fixtures__/
│   └── api-responses.ts  # Mock API responses for testing
└── test/                 # Integration test tools (CLI, MCP harness)
```

### Implementation Patterns

1. **Credential Encryption**: Session state and token cache are encrypted at rest using AES-256-GCM with a machine-specific key derived from hostname and username. Files have restrictive permissions (0o600).

2. **Server Class Pattern**: `TeamsServer` class encapsulates all state (browser manager, initialisation flag), allowing multiple server instances and simpler testing.

3. **Error Taxonomy**: Errors use machine-readable codes (`ErrorCode` enum), `retryable` flags, and `suggestions` arrays to help LLMs understand failures and recover appropriately.

4. **Result Types**: API functions return `Result<T, McpError>` for type-safe error handling with explicit success/failure discrimination.

5. **HTTP Utilities**: Centralised HTTP client (`utils/http.ts`) provides automatic retry with exponential backoff, request timeouts, and rate limit tracking.

6. **Dynamic Configuration from Session**: All tenant-specific configuration is extracted from the user's session localStorage, ensuring compatibility across different Teams environments (commercial, GCC, GCC-High, DoD):

   - **Region & Partition**: Extracted from `DISCOVER-REGION-GTM` (e.g., region `amer`, partition `02`). The `getRegion()` and `getTeamsBaseUrl()` helpers in `auth-guards.ts` provide cached access.
   - **Teams Base URL**: Extracted from the `chatServiceAfd` URL in `DISCOVER-REGION-GTM` (e.g., `https://teams.microsoft.com` or `https://teams.microsoft.us` for government clouds). API endpoints use this dynamically.
   - **User Details**: Extracted from `DISCOVER-USER-DETAILS` including user MRI, license info (Copilot, transcription, etc.), and user/tenant partitions.
   - **Service URLs**: Full URLs for chatsvc, CSA, and mt/part APIs are available in the config and passed to API endpoint builders.

   **Note**: The Substrate search URL (`substrate.office.com`) is currently hardcoded as we haven't found a config source for it. If GCC users report issues, this may need to be configurable.

7. **MCP Resources**: Passive resources (`teams://me/profile`, `teams://me/favorites`, `teams://status`) provide context discovery without tool calls.

7. **Tool Registry Pattern**: Tools are organised into logical groups (`search-tools.ts`, `message-tools.ts`, etc.) with a central registry (`tools/registry.ts`). This enables:
   - Better separation of concerns
   - Easier testing of individual tools
   - Simpler addition of new tools

8. **Auth Guards**: Reusable authentication check utilities in `utils/auth-guards.ts` return `Result` types for consistent error handling across API modules.

9. **Shared Constants**: Magic numbers are centralised in `constants.ts` for maintainability (page sizes, timeouts, thresholds).

10. **Markdown to Teams HTML**: Outgoing messages support markdown formatting via `markdownToTeamsHtml()` in `utils/parsers.ts`. This converts markdown (`**bold**`, `*italic*`, `` `code` ``, ` ```code blocks``` `, `~~strikethrough~~`, lists, newlines) to the HTML that Teams expects for `RichText/Html` messages. The converter is used by `sendMessage()` and `editMessage()` in `chatsvc-api.ts`. When messages contain @mentions or links, `parseContentWithMentionsAndLinks()` applies the same conversion to text segments between inline elements.

## How It Works

### Authentication Flow

All operations use direct API calls to Teams APIs. The browser is only used for authentication:

1. **Login**: Opens visible browser → user authenticates → session state saved → browser closed
2. **All subsequent operations**: Use cached tokens for direct API calls (no browser)
3. **Token expiry**: When tokens expire (~1 hour), proactive refresh is attempted; if that fails, user must re-authenticate via `teams_login`

This approach provides faster, more reliable operations compared to DOM scraping, with structured JSON responses and proper pagination support.

The server uses the system's installed browser rather than downloading Playwright's bundled Chromium (~180MB savings):

- **Windows**: Uses Microsoft Edge (always pre-installed on Windows 10+)
- **macOS/Linux**: Uses Google Chrome

This is configured via Playwright's `channel` option in `src/browser/context.ts`. If the system browser isn't available, a helpful error message suggests installing Chrome or running `npx playwright install chromium` as a fallback.

### Token Management

- Tokens are extracted from browser localStorage after login
- The Substrate search token (`SubstrateSearch-Internal.ReadWrite` scope) is required for search
- Tokens typically expire after ~1 hour
- **Proactive token refresh**: When tokens have less than 10 minutes remaining, the server automatically refreshes them using a headless browser. MSAL handles the token refresh when Teams loads, then we save the updated session state.
- This is seamless to the user - the browser is invisible
- If refresh fails, user must re-authenticate via `teams_login`

**How token refresh works:** When tokens are nearly expired, a headless browser loads Teams and triggers a search to force MSAL's `acquireTokenSilent`, then saves the refreshed session state.

**Testing token refresh:** Use `npm run cli -- login` which will attempt headless SSO first, only showing a browser if user interaction is actually required.

### API Authentication

Different Teams APIs use different authentication mechanisms:

| API | Auth Method | Module | Helper Function |
|-----|-------------|--------|-----------------|
| **Search** (Substrate v2/query) | JWT Bearer token from MSAL | `auth/token-extractor` | `getValidSubstrateToken()` |
| **People/Suggestions** (Substrate v1/suggestions) | Same JWT + `cvid`/`logicalId` fields | `auth/token-extractor` | `getValidSubstrateToken()` |
| **Messaging** (chatsvc) | `skypetoken_asm` cookie | `auth/token-extractor` | `extractMessageAuth()` |
| **Favorites** (csa/conversationFolders) | CSA token from MSAL + `skypetoken_asm` | `auth/token-extractor` | `extractCsaToken()` + `extractMessageAuth()` |
| **Threads** (chatsvc) | `skypetoken_asm` cookie | `auth/token-extractor` | `extractMessageAuth()` |
| **Calendar** (mt/part/calendarView) | Skype Spaces token (`api.spaces.skype.com` scope) + `skypetoken_asm` | `auth/token-extractor` | `extractSkypeSpacesToken()` |
| **Transcripts** (Substrate WorkingSetFiles) | Same JWT as Search (Substrate scope) + `Prefer` header | `auth/token-extractor` | `getValidSubstrateToken()` |

**Important**: The CSA API (for favorites) requires a GET request to retrieve data, POST only for modifications. The Substrate suggestions API requires `cvid` and `logicalId` correlation IDs in the request body.

**Region Discovery**: All regional APIs (chatsvc, csa, mt/part) use the region from the user's session via `getRegion()` in `auth-guards.ts`. This extracts the region from `DISCOVER-REGION-GTM` in localStorage (e.g., `amer`, `emea`, `apac`). For partitioned endpoints like mt/part (Calendar), the partition suffix (e.g., `02`) is also extracted from the same config.

### Session Persistence

Playwright's `storageState()` is used to save browser session state after login. This includes:
- Session cookies (for messaging APIs)
- MSAL tokens in localStorage (for search and people APIs)
- Tokens are extracted and cached for direct API use

Session state and token cache files are protected by:
1. **Encryption at rest**: AES-256-GCM encryption using a key derived from machine-specific values (hostname + username)
2. **File permissions**: Restrictive 0o600 permissions (owner read/write only)
3. **Automatic migration**: Existing plaintext files are automatically encrypted on first read

## MCP Tools

### Overview

| Tool | Purpose |
|------|---------|
| `teams_search` | Search messages with query operators, supports pagination |
| `teams_send_message` | Send a message to a Teams conversation (use `replyToMessageId` for thread replies) |
| `teams_get_me` | Get current user profile (email, name, ID) |
| `teams_get_frequent_contacts` | Get frequently contacted people (for name resolution) |
| `teams_search_people` | Search for people by name or email |
| `teams_login` | Trigger manual login (visible browser) |
| `teams_status` | Check auth status (search, messaging, favorites tokens) |
| `teams_get_favorites` | Get pinned/favourite conversations |
| `teams_add_favorite` | Pin a conversation to favourites |
| `teams_remove_favorite` | Unpin a conversation from favourites |
| `teams_save_message` | Bookmark a message |
| `teams_unsave_message` | Remove bookmark from a message |
| `teams_get_saved_messages` | Get list of saved/bookmarked messages with source references |
| `teams_get_followed_threads` | Get list of followed threads with source references |
| `teams_get_thread` | Get messages from a conversation/thread |
| `teams_find_channel` | Find channels by name (your teams + org-wide), shows membership |
| `teams_get_chat` | Get conversation ID for 1:1 chat with a person |
| `teams_create_group_chat` | Create a new group chat with multiple people |
| `teams_edit_message` | Edit one of your own messages |
| `teams_delete_message` | Delete one of your own messages (soft delete) |
| `teams_get_unread` | Get unread status for favourites (aggregate) or specific conversation |
| `teams_mark_read` | Mark a conversation as read up to a specific message |
| `teams_get_activity` | Get activity feed (mentions, reactions, replies, notifications) |
| `teams_search_emoji` | Search for emojis by name (standard + custom org emojis) |
| `teams_add_reaction` | Add an emoji reaction to a message |
| `teams_remove_reaction` | Remove an emoji reaction from a message |
| `teams_get_meetings` | Get meetings from calendar (upcoming/past by date range) |
| `teams_get_transcript` | Get meeting transcript (requires threadId from teams_get_meetings) |

### Design Philosophy

The toolset follows a **minimal tool philosophy**: fewer, more powerful tools that AI can compose together. Rather than convenience wrappers for common patterns, the AI builds queries using search operators.

### Tool Documentation

**Source of truth**: Tool parameters, descriptions, and usage guidance are defined in the tool definitions themselves (`src/tools/*.ts`). These descriptions are sent to AI assistants via MCP and should be comprehensive.

When adding or modifying tools, ensure the `description` field in the tool definition includes:
- What the tool does
- Key parameters and their meaning
- Common pitfalls or gotchas
- Related tools to use together

For manual testing of all tools, see `docs/MANUAL-TEST-SCRIPT.md`.

## Development

### Commands

```bash
npm run research      # Explore Teams APIs (visible browser, logs network calls)
npm run dev           # Run MCP server in development mode
npm run build         # Compile TypeScript
npm run lint          # Run ESLint (also lint:fix to auto-fix)
npm start             # Run compiled MCP server
```

### Testing

#### CLI / MCP Test Harness

The CLI (`npm run cli`) tests the server through the actual MCP protocol using in-memory transports. This verifies the full MCP layer works correctly, not just the underlying functions.

The harness can call **any tool** generically. Unrecognised commands are treated as tool names (with `teams_` prefix added if missing). Use `--key value` for parameters.

```bash
# List available MCP tools and shortcuts
npm run cli

# Generic tool call (any tool works - auto-prefixes teams_ if missing)
npm run cli -- teams_find_channel --query "support"
npm run cli -- find_channel --query "support"

# Common shortcuts
npm run cli -- search "your query"              # teams_search
npm run cli -- status                           # teams_status
npm run cli -- login                            # teams_login (tries headless SSO first)
npm run cli -- login --force true               # Clear session and re-login
npm run cli -- send "Hello!" --to "conv-id"     # teams_send_message
npm run cli -- thread --to "conv-id"            # teams_get_thread
npm run cli -- activity                         # teams_get_activity

# Output raw MCP response as JSON
npm run cli -- search "your query" --json

# Pagination: get page 2 (results 25-49)
npm run cli -- search "your query" --from 25 --size 25
```

#### Unit Tests

The project uses Vitest for unit testing pure functions. Tests focus on outcomes, not implementations.

```bash
npm test              # Run all tests once
npm run test:watch    # Run tests in watch mode
npm run test:coverage # Run tests with coverage report
npm run typecheck     # TypeScript type checking only
```

**Test Structure:**
- **`src/utils/parsers.ts`**: Pure parsing functions extracted for testability
- **`src/utils/parsers.test.ts`**: Unit tests for all parsing functions
- **`src/__fixtures__/api-responses.ts`**: Mock API response data based on real API structures

**What's Tested:**
- HTML stripping and entity decoding (`stripHtml`)
- Teams deep link generation (`buildMessageLink`)
- Message timestamp extraction (`extractMessageTimestamp`)
- Person suggestion parsing (`parsePersonSuggestion`)
- Search result parsing (`parseV2Result`, `parseSearchResults`)
- JWT profile extraction (`parseJwtProfile`)
- Token expiry calculations (`calculateTokenStatus`)
- People results parsing (`parsePeopleResults`)
- Base64 GUID decoding (`decodeBase64Guid`)
- User ID extraction from various formats (`extractObjectId`)

#### Integration Testing

For testing against the live Teams APIs:
- Use `npm run cli -- search "query"` to test via the full MCP protocol layer
- Use `npm run research` to explore new API patterns (logs all network traffic)

The MCP test harness (`test:mcp`) uses the SDK's `InMemoryTransport` to connect a test client to the server in-process, verifying that tool definitions, input validation, and response formatting all work correctly through the protocol layer.

#### CI/CD

GitHub Actions runs on every push and PR:
- Linting (`npm run lint`)
- Type checking (`npm run typecheck`)
- Unit tests (`npm test`)
- Build (`npm run build`)
- Documentation review (on main commits, debounced to 2 hours) - checks README/AGENTS.md accuracy against code

See `.github/workflows/ci.yml` and `.github/workflows/doc-reviewer.yml` for workflow configurations.

### Extending the MCP

#### Adding New Tools

1. Choose the appropriate tool file in `src/tools/` (or create a new one for a new category)
2. Define the input schema with Zod
3. Define the tool definition (MCP Tool interface)
4. Implement the handler function returning `ToolResult`
5. Export the registered tool and add it to the module's `*Tools` array
6. Add the new array to `src/tools/registry.ts` if creating a new category
7. Use `Result<T, McpError>` return types in underlying API modules
8. Add shared constants to `src/constants.ts` if needed

#### Adding New API Endpoints

1. Add endpoint URL to `src/utils/api-config.ts`
2. Create a function in the appropriate `src/api/*.ts` module
3. Use `httpRequest()` from `src/utils/http.ts` for automatic retry and timeout handling
4. Return `Result<T, McpError>` for type-safe error handling

#### Capturing New API Endpoints

Run `npm run research`, perform actions in Teams, and check the terminal output for captured requests.

## Troubleshooting

### Session/Token Expired
If API calls fail with authentication errors:
1. Call `teams_login` with `forceNew: true`
2. Or delete the config directory (`~/.teams-mcp-server/` on macOS/Linux, `%APPDATA%\teams-mcp-server\` on Windows) and run `npm run cli -- login`

### Browser Won't Launch (for login)
- Ensure you have Chrome (macOS/Linux) or Edge (Windows) installed
- On Windows, Edge should be pre-installed; try updating Windows if missing
- On macOS/Linux, install Chrome from https://www.google.com/chrome/
- Alternatively, download Playwright's bundled browser: `npx playwright install chromium`
- Check for existing browser processes that may be blocking

### Login Timeout with MFA
MCP clients like Cursor have request timeouts (typically 2-5 minutes). If your organisation's SSO/MFA flow takes longer than this, the MCP request may timeout.

**What happens:** The AI receives a timeout error, but the login process continues in the background. Complete the MFA in the browser - the session will be saved and subsequent tool calls will work.

**Why this is rare:** The server attempts headless SSO first (no user interaction needed). A visible browser only opens when credentials are actually required. Most logins complete silently via SSO without hitting any timeout.

### Search Doesn't Find All Thread Replies
The Substrate search API is a **full-text search** which only returns messages matching the search terms. If someone replied to your message but their reply doesn't contain your search keywords, it won't appear in results.

**Example:** Searching for "Easter blockout" won't find a reply that says "Given World of Frozen opens the week before, I'd put a fair amount of money on 'yes'", even though it's a direct reply.

**Workaround:** After finding a message of interest, use `teams_get_thread` with the `conversationId` to retrieve the full thread context including all replies.

### Message Deep Links

Teams requires different deep link formats depending on conversation type:

| Conversation Type | Format | Notes |
|-------------------|--------|-------|
| **Channel (top-level)** | `/l/message/{channelId}/{msgTimestamp}` | No extra params needed |
| **Channel (thread reply)** | `/l/message/{channelId}/{msgTimestamp}?parentMessageId={parentId}` | Parent ID from `ClientConversationId;messageid=xxx` |
| **1:1 / Group chat** | `/l/message/{chatId}/{msgTimestamp}?context={"contextType":"chat"}` | Context param required |
| **Meeting chat** | `/l/message/{meetingId}/{msgTimestamp}?context={"contextType":"chat"}` | Context param required |

**Conversation ID patterns:**
- Channels: `19:xxx@thread.tacv2`
- Meetings: `19:meeting_xxx@thread.v2`
- 1:1 chats: `19:guid_guid@unq.gbl.spaces`
- Group chats: `19:xxx@thread.v2` (non-meeting)

**Detecting thread replies:** Compare the `messageid` in `ClientConversationId` with the message's own timestamp from `DateTimeReceived`. If they differ, it's a thread reply and needs `parentMessageId`.

## Reference

### File Locations

Session files are stored in a user-specific config directory to ensure consistency regardless of how the server is invoked (npx, global install, local dev, etc.):

- **macOS/Linux**: `~/.teams-mcp-server/`
- **Windows**: `%APPDATA%\teams-mcp-server\` (e.g., `C:\Users\name\AppData\Roaming\teams-mcp-server\`)

Contents:
- `session-state.json` (encrypted browser session)
- `token-cache.json` (encrypted OAuth tokens)
- `.user-data/` (browser profile)

Legacy session files from the project root (`./session-state.json`) are automatically migrated to the new location on first read.

Development-only files (created in project root):
- **Debug output**: `./debug-output/` (gitignored, screenshots and HTML dumps)

Development files:
- **API reference**: `./docs/API-REFERENCE.md`
- **Session data reference**: `./docs/SESSION-DATA-REFERENCE.md`

### API Internals

**Conversation Types:** The chatsvc API returns `threadType` (`topic`, `space`, `meeting`, `chat`) and `productThreadType` (`TeamsStandardChannel`, `TeamsTeam`, `TeamsPrivateChannel`, `Meeting`, `Chat`, `OneOnOne`). See `docs/API-REFERENCE.md` for details.

**Virtual Conversations:** Special IDs like `48:saved`, `48:threads`, `48:mentions`, `48:notifications`, `48:notes` aggregate data across conversations. Messages include `clumpId` for the source conversation.

**User ID Formats:** The `extractObjectId()` function in `parsers.ts` handles all ID formats: raw GUIDs, MRIs (`8:orgid:...`), tenant-suffixed IDs, base64-encoded GUIDs (little-endian), and Skype IDs.

**1:1 Chat ID format:** `19:{userId1}_{userId2}@unq.gbl.spaces` (GUIDs sorted lexicographically).

**Deleted Messages:** The chatsvc messages API returns deleted messages with empty `content` and a `deletetime` property in `properties`. These are filtered out in `getThreadMessages()` to avoid confusing the AI with phantom "empty" messages that appear to be newer than actual content.

**Message Ordering:** The chatsvc messages API returns messages in **descending order (newest first)** by default. When `startTime` is provided, messages are returned in **ascending order** from that timestamp. The `getThreadMessages()` function defaults to newest-first (`order: 'desc'`) but accepts an `order` parameter to switch to oldest-first (`'asc'`) for chronological reading.

### API Endpoints

See `docs/API-REFERENCE.md` for full endpoint documentation with request/response examples.

Regional identifiers: `amer`, `emea`, `apac`

### Possible Tools

Based on API research, these tools could be implemented:

| Tool | API | Difficulty |
|------|-----|------------|
| `teams_get_person` | Delve person API | Easy |
| `teams_get_files` | AllFiles API | Medium |

**Known Limitations:**
- **Chat list** - Partially addressed by `teams_get_favorites` (pinned chats) and `teams_get_frequent_contacts` (common contacts), but no full chat list API
- **Presence/Status** - Real-time via WebSocket, not HTTP
- **Calendar** - Outlook APIs exist but need separate research

## Dependencies

- `@modelcontextprotocol/sdk`: MCP protocol implementation
- `playwright`: Browser automation
- `zod`: Runtime input validation
- `vitest`: Unit testing framework (dev)
