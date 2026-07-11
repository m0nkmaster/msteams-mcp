# Teams MCP Server & CLI

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

### CLI (alternative to the MCP server)

The same functionality is also available as a standalone command-line tool, `msteams`. This is useful when you want Teams from a shell or script instead of an MCP client (for example, if an MCP integration is unreliable in your environment).

Install it globally from npm:

```bash
npm install -g msteams-mcp
```

This installs two binaries: `msteams-mcp` (the MCP server) and `msteams` (the CLI). Or run it without installing:

```bash
npx -y msteams-mcp msteams status
```

See [CLI Usage](#cli-usage) for commands.

## Available Tools

### Search & Discovery

| Tool | Description |
|------|-------------|
| `teams_search` | Search Teams messages with operators (`from:`, `sent:`, `in:`, `hasattachment:`, etc.) |
| `teams_search_email` | Search emails in your mailbox (same auth as Teams — no extra login) |
| `teams_list_chats` | List recent conversations (1:1, group, meeting, channel) with a last-message preview |
| `teams_get_message` | Get a single message by ID with full content (any age); includes reactions |
| `teams_get_thread` | Get messages from a conversation/thread; includes reactions; `threadRootId` scopes to one channel thread; `fromUrl` accepts a Teams message deep link |
| `teams_find_channel` | Find channels by name (your teams + org-wide discovery) |
| `teams_get_activity` | Get activity feed (mentions, reactions, replies, notifications) |

### Messaging

| Tool | Description |
|------|-------------|
| `teams_send_message` | Send a message (default: self-chat/notes). `replyToMessageId` for thread replies, `subject` for a new channel thread, `scheduleAt` to schedule, `contentType` (`auto`/`text`/`html`/`markdown`) to control formatting |
| `teams_wait_for_reply` | Block until a new message arrives (server-side poll, capped ~110s); idempotent `after`/`nextAfter` cursor — pair with `teams_send_message` |
| `teams_edit_message` | Edit one of your own messages (`contentType` supported) |
| `teams_delete_message` | Delete one of your own messages (soft delete) |

### People & Contacts

| Tool | Description |
|------|-------------|
| `teams_get_me` | Get current user profile (email, name, ID) |
| `teams_search_people` | Search for people by name or email |
| `teams_get_frequent_contacts` | Get frequently contacted people (useful for name resolution) |
| `teams_get_person` | Resolve one or more MRIs to full profiles (name, email, job title, department) |
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

### Calendar & Meetings

| Tool | Description |
|------|-------------|
| `teams_get_meetings` | Get meetings from calendar (defaults to next 7 days) |
| `teams_get_transcript` | Get meeting transcript (requires `threadId` from `teams_get_meetings`) |

`teams_get_meetings` returns: subject, times, organiser, join URL, `threadId` for meeting chat. Use `threadId` with `teams_get_thread` to read meeting chat, or with `teams_get_transcript` to get the full transcript with speakers and timestamps.

### Files

| Tool | Description |
|------|-------------|
| `teams_get_shared_files` | Get files and links shared in a conversation (supports pagination) |

Returns both files (name, extension, URL, size) and links (URL, title), along with who shared each item. Works for channels, group chats, 1:1 chats, and meeting chats.

### Session

| Tool | Description |
|------|-------------|
| `teams_login` | Trigger manual login (opens browser) |
| `teams_status` | Check authentication and session state |

### Search Operators

Both `teams_search` (Teams messages) and `teams_search_email` (emails) support native operators:

```
from:sarah@company.com     # Messages/emails from person
sent:2026-01-20            # From specific date
sent:>=2026-01-15          # Since date
in:project-alpha           # Messages in channel (Teams only)
subject:"budget"           # By subject (email)
"Rob Smith"                # Find @mentions (name in quotes)
hasattachment:true         # With files
is:unread                  # Unread emails (email only)
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

## CLI Usage

`msteams` exposes every tool the MCP server does - full parity, same authentication, same session files. Run with no arguments to list all tools and shortcuts.

If you installed globally (`npm install -g msteams-mcp`), invoke it directly:

```bash
# List available tools and shortcuts
msteams

# Check authentication status
msteams status

# Log in (opens a browser; tries silent SSO first)
msteams login
msteams login --force        # clear session and re-login

# Search messages
msteams search "meeting notes"
msteams search "project" --from 0 --size 50

# Search emails
msteams teams_search_email --query "from:sarah@company.com"

# Send a message (default: your own notes/self-chat)
msteams send "Hello from Teams MCP!"
msteams send "Message" --to "conversation-id"

# People, contacts, favourites, activity, unread
msteams people "john smith"
msteams favorites
msteams activity
msteams unread

# Any tool by name (the teams_ prefix is optional)
msteams teams_search_emoji --query "heart"
msteams find_channel --query "support"

# Machine-readable output
msteams search "query" --json
```

**Command form:** `msteams <command> [primaryArg] [--key value ...]`. Any unrecognised command is treated as a tool name (`teams_` is added automatically). Common flags like `--to`, `--from`, `--size`, `--query`, `--force` map to the matching tool parameters; run `msteams` with no arguments to see the full list.

### From a repo clone

If you're working from source, the same CLI is wired to `npm run cli` (runs via `tsx`, no build needed):

```bash
npm run cli                          # list tools
npm run cli -- search "your query"
npm run cli -- status
npm run cli -- send "Hi" --to "conversation-id"
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

Contents: `session-state.json`, `token-cache.json`, `browser-profile/`

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
