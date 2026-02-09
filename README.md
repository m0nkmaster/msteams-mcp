# Teams MCP Server

[![CI](https://github.com/m0nkmaster/msteams-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/m0nkmaster/msteams-mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/node-%3E%3D18-brightgreen.svg)](https://nodejs.org/)

An MCP (Model Context Protocol) server that enables AI assistants to interact with Microsoft Teams. Search messages, send replies, manage favourites, and more.

## How It Works

This server calls Microsoft's Teams APIs directly (Substrate, chatsvc, CSA)  - the same APIs the Teams web app uses. No Azure AD app registration or admin consent required.

**Authentication flow:**
1. AI runs `teams_login` to open a browser for you to log in
2. OAuth tokens are extracted and cached
3. All operations use cached tokens directly (no browser needed)
4. Automatic token refresh (~1 hour)

**Security:** Uses the same authentication as the Teams web client - your access is limited to what your account can already do.

## Installation

### Prerequisites

- Node.js 18+
- A Microsoft account with Teams access
- Google Chrome, Microsoft Edge, or Chromium browser installed

### Configure Your MCP Client

Add to your MCP client configuration (e.g., Claude Desktop, Windsurf, Cursor):

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["-y", "msteams-mcp@latest"]
    }
  }
}
```

That's it. `npx` will automatically download and run the latest version.

### From Source (alternative)

If you prefer to run from a local clone:

```bash
git clone https://github.com/m0nkmaster/msteams-mcp.git
cd msteams-mcp
npm install && npm run build
```

Then configure your MCP client:

```json
{
  "mcpServers": {
    "teams": {
      "command": "node",
      "args": ["/path/to/msteams-mcp/dist/index.js"]
    }
  }
}
```

The server uses your system's Chrome (macOS/Linux) or Edge (Windows) for authentication.

## Available Tools

### Search & Discovery

| Tool | Description |
|------|-------------|
| `teams_search` | Search messages with operators (`from:`, `sent:`, `in:`, `hasattachment:`, etc.) |
| `teams_get_thread` | Get messages from a conversation/thread |
| `teams_find_channel` | Find channels by name (your teams + org-wide discovery) |
| `teams_get_activity` | Get activity feed (mentions, reactions, replies, notifications) |

### Messaging

| Tool | Description |
|------|-------------|
| `teams_send_message` | Send a message (default: self-chat/notes). Use `replyToMessageId` for thread replies |
| `teams_edit_message` | Edit one of your own messages |
| `teams_delete_message` | Delete one of your own messages (soft delete) |

### People & Contacts

| Tool | Description |
|------|-------------|
| `teams_get_me` | Get current user profile (email, name, ID) |
| `teams_search_people` | Search for people by name or email |
| `teams_get_frequent_contacts` | Get frequently contacted people (useful for name resolution) |
| `teams_get_chat` | Get conversation ID for 1:1 chat with a person |
| `teams_create_group_chat` | Create a new group chat with multiple people (2+ others) |

### Organisation

| Tool | Description |
|------|-------------|
| `teams_get_favorites` | Get pinned/favourite conversations |
| `teams_add_favorite` | Pin a conversation |
| `teams_remove_favorite` | Unpin a conversation |
| `teams_save_message` | Bookmark a message |
| `teams_unsave_message` | Remove bookmark from a message |
| `teams_get_saved_messages` | Get list of saved/bookmarked messages with source references |
| `teams_get_followed_threads` | Get list of followed threads with source references |
| `teams_get_unread` | Get unread counts (aggregate or per-conversation) |
| `teams_mark_read` | Mark a conversation as read up to a message |

### Reactions

| Tool | Description |
|------|-------------|
| `teams_search_emoji` | Search for emojis by name (standard + custom org emojis) |
| `teams_add_reaction` | Add an emoji reaction to a message |
| `teams_remove_reaction` | Remove an emoji reaction from a message |

**Quick reactions:** `like`, `heart`, `laugh`, `surprised`, `sad`, `angry` can be used directly without searching.

### Calendar

| Tool | Description |
|------|-------------|
| `teams_get_meetings` | Get meetings from calendar (defaults to next 7 days) |

Returns: subject, times, organiser, join URL, `threadId` for meeting chat. Use with `teams_get_thread` to read meeting discussions.

### Session

| Tool | Description |
|------|-------------|
| `teams_login` | Trigger manual login (opens browser) |
| `teams_status` | Check authentication and session state |

### Search Operators

The search supports Teams' native operators:

```
from:sarah@company.com     # Messages from person
sent:2026-01-20            # Messages from specific date
sent:>=2026-01-15          # Messages since date
in:project-alpha           # Messages in channel
"Rob Smith"                # Find @mentions (name in quotes)
hasattachment:true         # Messages with files
NOT from:email@co.com      # Exclude results
```

Combine operators: `from:sarah@co.com sent:>=2026-01-18 hasattachment:true`

**Note:** `@me`, `from:me`, `to:me` do NOT work. Use `teams_get_me` first to get your email/displayName. `sent:today` works, but `sent:lastweek` and `sent:thisweek` do NOT - use explicit dates or omit (results are sorted by recency).

## MCP Resources

The server also exposes passive resources for context discovery:

| Resource URI | Description |
|--------------|-------------|
| `teams://me/profile` | Current user's profile |
| `teams://me/favorites` | Pinned conversations |
| `teams://status` | Authentication status |

## CLI Tools (Development)

For local development, CLI tools are available for testing and debugging:

```bash
# Check authentication status
npm run cli -- status

# Search messages
npm run cli -- search "meeting notes"
npm run cli -- search "project" --from 0 --size 50

# Send messages (default: your own notes/self-chat)
npm run cli -- send "Hello from Teams MCP!"
npm run cli -- send "Message" --to "conversation-id"

# Force login
npm run cli -- login --force

# Output as JSON
npm run cli -- search "query" --json
```

### MCP Test Harness

Test the server through the actual MCP protocol:

```bash
# List available tools
npm run test:mcp

# Call any tool
npm run test:mcp -- search "your query"
npm run test:mcp -- status
npm run test:mcp -- people "john smith"
npm run test:mcp -- favorites
npm run test:mcp -- activity              # Get activity feed
npm run test:mcp -- unread                # Check unread counts
npm run test:mcp -- teams_search_emoji --query "heart"  # Search emojis
```

## Limitations

- **Login required** - Run `teams_login` to authenticate (opens browser)
- **Token expiry** - Tokens expire after ~1 hour; headless refresh is attempted or run `teams_login` again when needed
- **Undocumented APIs** - Uses Microsoft's internal APIs which may change without notice
- **Search limitations** - Full-text search only; thread replies not matching search terms won't appear (use `teams_get_thread` for full context)
- **Own messages only** - Edit/delete only works on your own messages

## Session Files

Session files are stored in a user config directory (encrypted):

- **macOS/Linux**: `~/.teams-mcp-server/`
- **Windows**: `%APPDATA%\teams-mcp-server\`

Contents: `session-state.json`, `token-cache.json`, `.user-data/`

If your session expires, call `teams_login` or delete the config directory.

## Development

For local development:

```bash
git clone https://github.com/m0nkmaster/msteams-mcp.git
cd msteams-mcp
npm install
npm run build
```

Development commands:

```bash
npm run dev          # Run MCP server in dev mode
npm run build        # Compile TypeScript
npm run lint         # Run ESLint
npm run research     # Explore Teams APIs (logs network calls)
npm test             # Run unit tests
npm run typecheck    # TypeScript type checking
```

For development with hot reload, configure your MCP client:

```json
{
  "mcpServers": {
    "teams": {
      "command": "npx",
      "args": ["tsx", "/path/to/msteams-mcp/src/index.ts"]
    }
  }
}
```

See [AGENTS.md](AGENTS.md) for detailed architecture and contribution guidelines.

---

## Teams Chat Export Bookmarklet

This repo also includes a standalone bookmarklet for exporting Teams chat messages to Markdown. See [teams-bookmarklet/README.md](teams-bookmarklet/README.md).
