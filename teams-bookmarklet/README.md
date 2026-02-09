# Teams Chat Export

Export Microsoft Teams chat messages to Markdown format.

## What's New in v2.0

This version uses knowledge from building the [Teams MCP server](../README.md) to provide:

- **API-first approach**: Uses Teams' internal chatsvc API for fast, reliable export (10x faster than DOM scraping)
- **Message deep links**: Each message includes a clickable link to open it directly in Teams
- **Automatic region detection**: Tries amer/emea/apac until one works
- **Graceful fallback**: Falls back to DOM scraping if API fails
- **Better date filtering**: Increased default from 2 days to 7 days

## Features

- Captures sender names and timestamps (using ISO dates for accuracy)
- **Generates deep links to each message** (click to open in Teams)
- Preserves links (formatted as markdown links)
- Captures emoji reactions
- Detects edited messages
- **Detects open threads** and offers to export just the thread
- Filters out "Replied in thread" preview messages
- Sorts messages chronologically
- Filters by configurable date range (days back)
- Groups messages by date with section headers

## Usage

### Console Script (Recommended)

Due to Teams' strict Content Security Policy, bookmarklets are blocked. Use the console script instead:

1. Open Teams in your browser (teams.microsoft.com)
2. Navigate to the chat/channel you want to export
3. Open DevTools (F12) → Console tab
4. Copy the contents of `teams-export.js` and paste into the console
5. Press Enter

### Export Dialog

You'll see a dialog showing:

1. **Chat name** - Detected automatically from the page
2. **Conversation ID** - If detected, enables API export (much faster)
3. **Export method** - Shows whether API or DOM scraping will be used
4. **Days to capture** - How many days back to include (default: 7)
5. **Include message links** - Toggle to include/exclude deep links
6. **Include thread replies** - For channels, fetches all replies to each post (enabled by default)
7. Click **Export**
8. Wait for the progress bar
9. Markdown is copied to your clipboard

### If a Thread is Open

If you have a thread panel open when you run the script, you'll see:

- **Export This Thread Only** - Immediately exports just the thread messages
- **Export Full Chat** - Closes the thread and shows the full chat export options

## Output Format

```markdown
# Project Alpha Team

**Exported:** 26/01/2026
**Messages:** 47

---

## Monday 20 January 2026

**Smith, Jane** (09:15) [[link](https://teams.microsoft.com/l/message/19:xxx@thread.tacv2/1737363300000)]:
> Morning all! Quick reminder that the design review is at 2pm today.
>
> 🔗 [Meeting link](https://example.com/meeting)
>
> 3 Like reactions

**Chen, David** (10:30) *(edited)* [[link](https://teams.microsoft.com/l/message/19:xxx@thread.tacv2/1737367800000)]:
> Just pushed the latest changes to the feature branch. Ready for review.
>
> 🔗 [Pull request](https://github.com/example/repo/pull/123)

---

## Tuesday 21 January 2026

**Williams, Sarah** (08:45) [[link](https://teams.microsoft.com/l/message/19:xxx@thread.tacv2/1737448500000)]:
> Has anyone seen the updated requirements doc?
```

## How It Works

### API Export (Preferred)

1. **Conversation ID extraction**: Tries to get the conversation ID from URL params, hash, DOM attributes, or localStorage
2. **Region auto-detection**: Uses session config if available, otherwise tries amer, emea, and apac regions
3. **Paginated fetch**: Fetches up to 2000 messages (200 per page × 10 pages)
4. **Thread expansion (channels)**: For each top-level post, fetches replies via the thread API endpoint
5. **Message parsing**: Strips HTML, extracts timestamps, builds deep links

### DOM Export (Fallback)

Used when API export fails or conversation ID cannot be detected:

1. **Scrolling**: Teams uses virtual scrolling (only visible messages are in DOM). Scrolls to load messages.
2. **Element extraction**: Parses DOM elements for sender, timestamp, content, reactions.
3. **Message ID calculation**: Computes message ID from timestamp for deep links.

## Message Links

Each exported message includes a deep link in the format:

```
https://teams.microsoft.com/l/message/{conversationId}/{messageTimestamp}
```

Clicking this link opens Teams and navigates directly to that specific message. This is useful for:

- Referencing specific messages in documentation
- Jumping back to important conversations
- Sharing links with colleagues

## Troubleshooting

### "API failed, falling back to DOM"

This means the API couldn't be reached. Common causes:
- Session expired (refresh the page and try again)
- Network issues
- Unusual conversation type

The script automatically falls back to DOM scraping, so export still works.

### "Could not detect conversation ID"

The script couldn't find the conversation ID. This means:
- API export won't be attempted
- DOM scraping will be used (slower but works)

This typically happens with older URLs or unusual navigation paths.

### Not all messages captured (DOM mode)

For very long chats with DOM scraping, the scroll might not capture everything. Try:
- Running it twice
- Using a smaller "days to capture" value

### Clipboard access denied

Check the browser console (F12 → Console) where the markdown is also logged.

## Technical Notes

- Works with Teams web app (teams.microsoft.com)
- Tested with the new Teams interface (Fluent UI)
- API endpoint: `/api/chatsvc/{region}/v1/users/ME/conversations/{id}/messages`
- Regions tried: amer, emea, apac
- Cookies are automatically included for authentication

## Files

- `teams-export.js` - The full export script (run in browser console)
- `teams-chat-export-bookmarklet.md` - Alternative documentation
- `README.md` - This file
