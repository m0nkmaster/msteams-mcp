# Teams Chat Export v2.0

> **Note:** This is an alternative/legacy documentation file. See [README.md](./README.md) for the main documentation.

Export Microsoft Teams chat messages to Markdown format using Teams' internal APIs.

## Quick Start

1. Open Teams in your browser (teams.microsoft.com)
2. Navigate to the chat/channel you want to export
3. Open DevTools (F12) → Console tab
4. Copy the contents of `teams-export.js` and paste into the console
5. Press Enter
6. Configure options in the dialog
7. Click **Export**

## Key Features

- **API-first approach** - Uses Teams' chatsvc API for fast, reliable export
- **Message deep links** - Click to open any message directly in Teams
- **Automatic fallback** - Falls back to DOM scraping if API fails
- **Thread detection** - Offers to export just the current thread if one is open
- **Date filtering** - Configure how many days back to capture

## How the API Export Works

When you have a chat open, the script:

1. Extracts the conversation ID from the page/URL
2. Calls `GET /api/chatsvc/{region}/v1/users/ME/conversations/{id}/messages`
3. Uses cookies automatically (same-origin, credentials included)
4. Tries amer, emea, apac regions until one works
5. Paginates to get up to 2000 messages

This is the same API the Teams MCP server uses, adapted for browser context.

## Comparison: API vs DOM Scraping

| Aspect | API Export | DOM Scraping |
|--------|-----------|--------------|
| Speed | Fast (~2 seconds) | Slow (~30+ seconds) |
| Reliability | High | Medium (virtual scrolling can miss messages) |
| Message Links | ✅ Yes | ✅ Yes (computed from timestamp) |
| Reactions | ❌ Not in API | ✅ Yes |
| Auth | Cookies (automatic) | N/A |

## See Also

- [README.md](./README.md) - Full documentation
- [Teams MCP Server](../README.md) - The MCP server this knowledge came from
