# Teams MCP Manual Test Script

This is a test script for verifying all MCP tools work correctly before release.

**Usage**: Ask an AI agent to "run this test script" and it will exercise all tools, reporting pass/fail for each. This verifies authentication, search, messaging, and management features.

---

## Available Tools

### Authentication & Status
| Tool | Purpose |
|------|---------|
| `teams_status` | Check authentication status and token expiry |
| `teams_get_me` | Get your profile (email, name, ID) |
| `teams_login` | Trigger manual login if session expired |

### Search & Discovery
| Tool | Purpose |
|------|---------|
| `teams_search` | Search Teams messages with query operators |
| `teams_search_email` | Search emails in your mailbox (same auth as Teams) |
| `teams_find_channel` | Find channels by name (your teams + org-wide) |
| `teams_search_people` | Search for people by name or email |
| `teams_get_frequent_contacts` | Get frequently contacted people (for name resolution) |

### Reading Messages
| Tool | Purpose |
|------|---------|
| `teams_get_thread` | Get messages from a conversation/thread |
| `teams_get_unread` | Get unread status (aggregate or specific conversation) |
| `teams_get_activity` | Get activity feed (mentions, reactions, replies) |

### Sending Messages
| Tool | Purpose |
|------|---------|
| `teams_send_message` | Send a message with optional @mentions (defaults to self-notes). Use `replyToMessageId` for thread replies |
| `teams_get_chat` | Get conversation ID for 1:1 chat with a person |
| `teams_create_group_chat` | Create a new group chat with multiple people (2+ others) |

**@Mentions**: Use `@[DisplayName](mri)` syntax inline in content. Get MRI from `teams_search_people` or `teams_get_frequent_contacts`. Example: `"Hey @[John](8:orgid:abc...), check this"`

### Message Management
| Tool | Purpose |
|------|---------|
| `teams_edit_message` | Edit your own messages |
| `teams_delete_message` | Delete your own messages (soft delete) |
| `teams_save_message` | Bookmark a message |
| `teams_unsave_message` | Remove bookmark |
| `teams_get_saved_messages` | Get list of saved/bookmarked messages |
| `teams_get_followed_threads` | Get list of followed threads |
| `teams_mark_read` | Mark conversation as read up to a message |

### Reactions
| Tool | Purpose |
|------|---------|
| `teams_add_reaction` | Add an emoji reaction to a message |
| `teams_remove_reaction` | Remove an emoji reaction from a message |
| `teams_search_emoji` | Search for emojis by name (standard + custom org emojis) |

### Files
| Tool | Purpose |
|------|---------|
| `teams_get_shared_files` | Get files/links shared in a conversation (channels, chats, meetings) |

### Calendar & Meetings
| Tool | Purpose |
|------|---------|
| `teams_get_meetings` | Get meetings from calendar (defaults to next 7 days, set startDate to past for recent) |
| `teams_get_transcript` | Get meeting transcript (requires threadId from teams_get_meetings) |

### Favourites
| Tool | Purpose |
|------|---------|
| `teams_get_favorites` | Get pinned/favourite conversations |
| `teams_add_favorite` | Pin a conversation |
| `teams_remove_favorite` | Unpin a conversation |

### Assignments (education tenants only)
| Tool | Purpose |
|------|---------|
| `teams_list_assignments` | List your assignments across classes (active, completed or all) |
| `teams_get_assignment` | Get one assignment's full detail and your submission |
| `teams_submission_action` | Turn in, undo turn-in, or mark viewed on your own submission |

---

## Search Operators

| Operator | Example | Description |
|----------|---------|-------------|
| `from:` | `from:user@company.com` | Messages from a person |
| `sent:` | `sent:2026-01-20`, `sent:>=2026-01-15` | Messages by date (explicit dates only) |
| `in:` | `in:channel-name` | Messages in a channel |
| `"Name"` | `"John Smith"` | Find @mentions |
| `NOT` | `NOT from:user@company.com` | Exclude results |
| `hasattachment:` | `hasattachment:true` | Messages with files |

**Important**: `@me`, `from:me`, `to:me` do NOT work. Use `teams_get_me` first to get your actual email/name. `sent:today` works, but `sent:lastweek` and `sent:thisweek` do NOT - use explicit dates (e.g., `sent:>=2026-01-18`) or omit since results are sorted by recency.

---

## Common Workflows

### Find messages mentioning me
```
1. teams_get_me → get displayName and email
2. teams_search "Display Name" NOT from:my.email@company.com
```

### Find and message someone
```
1. teams_search_people "person name" → get their user ID
2. teams_get_chat userId → get conversation ID
3. teams_send_message content="Hello" conversationId="..."
```

### Send a message with @mentions
```
1. teams_search_people "person name" → get their MRI (e.g., "8:orgid:abc...")
2. teams_send_message content="Hey @[Person Name](8:orgid:abc...), can you check this?" conversationId="..."
```
Use `@[DisplayName](mri)` syntax inline. The display name can be anything (e.g., first name only).

### Reply to a channel thread
```
1. teams_search or teams_get_thread → get conversationId and messageId
2. teams_send_message content="Reply" conversationId="..." replyToMessageId="..."
```

### Create a group chat
```
1. teams_search_people "person 1" → get their MRI
2. teams_search_people "person 2" → get their MRI
3. teams_create_group_chat userIds=["mri1","mri2"] topic="Chat Name" → get conversationId
4. teams_send_message content="Hello everyone!" conversationId="..."
```

### Test reply and delete (for manual testing)
```
1. teams_get_me → get your email
2. teams_search from:your.email@company.com → find your own channel message
3. teams_send_message content="Test reply" conversationId="..." replyToMessageId="..." → reply to your own message
4. teams_delete_message → delete the test reply to clean up
```

### Check unread messages
```
1. teams_get_unread → aggregate unread across favourites
2. teams_get_thread conversationId="..." → read the messages
3. teams_mark_read conversationId="..." messageId="..." → mark as read
```

### Find a channel to read
```
1. teams_find_channel "channel name" → get channelId
2. teams_get_thread conversationId="channelId" → read messages
```

### Understanding channel thread structure
When reading channel messages with `teams_get_thread`:
- **Default order is newest-first** - latest messages appear at the top of the results
- Use `order: "asc"` for chronological reading order (oldest-first)
- Messages **without** `threadRootId` are **top-level posts** (new conversations)
- Messages **with** `threadRootId` are **replies** to the post with that ID
- Use `isThreadReply: true` to identify replies
- The `when` field gives day of week (e.g., "Friday, January 30, 2026, 10:45 AM UTC")
- To summarize "what topics are being discussed", filter to messages where `isThreadReply` is not true (i.e., top-level posts only)

### Review saved messages
```
1. teams_get_saved_messages → get list of bookmarked messages
2. teams_get_thread conversationId="sourceConversationId" → read full context
```

### Check followed threads
```
1. teams_get_followed_threads → get list of threads you're following
2. teams_get_thread conversationId="sourceConversationId" → read thread content
```

### Get shared files from a conversation
```
1. teams_get_favorites or teams_find_channel → get a conversationId
2. teams_get_shared_files conversationId="..." → list files with names, URLs, sizes, who shared them
3. Use skipToken from response to paginate if there are more files
```

### Search emails
```
1. teams_search_email query="from:sarah@company.com" → search emails from a person
2. teams_search_email query="subject:budget sent:>=2026-01-15" → search by subject and date
3. teams_search_email query="hasattachment:true is:unread" → unread emails with attachments
```
Returns: subject, sender, recipients, preview, read status, importance, attachments, and pagination.
Uses the same Substrate token as Teams search — no additional login required.

### Get a meeting transcript
```
1. teams_get_meetings startDate="past date" → find the meeting, get threadId
2. teams_get_transcript threadId="19:meeting_xxx@thread.v2" meetingDate="startTime from step 1"
```
Returns: meeting title, speakers list, entry count, and full formatted transcript with timestamps.
Note: Only works for meetings where transcription was enabled.

### React to a message
```
1. teams_search or teams_get_thread → find the message
2. teams_add_reaction conversationId="..." messageId="..." emoji="like"
```
Quick reactions: `like` (👍), `heart` (❤️), `laugh` (😂), `surprised` (😮), `sad` (😢), `angry` (😠)

### Find custom emojis
```
1. teams_search_emoji query="thumbs" → find emoji keys
2. teams_add_reaction emoji="..." → use the key from search
```

### Check assignments (education tenant)
```
1. teams_list_assignments → active work, earliest due first
2. teams_list_assignments nextLink="..." → next page; statusFilter comes from the link
3. teams_get_assignment classId="..." assignmentId="..." → full detail and your submission
4. teams_submission_action action="view" classId="..." assignmentId="..." submissionId="..." → mark viewed
```
Only confirm `submit`/`unsubmit` against a test assignment you own: they change your real submission.

### Assignments on a non-education account
```
1. teams_list_assignments → expect ACCESS_DENIED or AUTH_INTERACTION_REQUIRED
2. teams_get_me, teams_search, teams_list_chats → must still work, with no browser opened
```

---

## Known Limitations

| Limitation | Details |
|------------|---------|
| Save message | Only works on root messages, not thread replies |
| Unread on channels | May fail ACL check; works reliably for chats/meetings |
| Token expiry | Tokens last ~1 hour; call `teams_login` to refresh |
| Assignments | Education tenants only; failures never affect other tools. Submit/unsubmit not yet verified live |

---

## Safety Guidelines

- **Never send messages to others** without explicit user confirmation
- **Default to self-notes** (`48:notes`) for testing or drafts
- **Verify recipients** before sending by confirming email/name
- **Be cautious with delete** - it's a soft delete but still removes content

---

## Example Prompts

### Catch up on Teams
> "Check my Teams for any unread messages or mentions. Summarise what needs my attention."

### Find information
> "Search Teams for recent discussions about [topic]. Who's been involved and what are the key points?"

### Draft a message
> "Help me draft a message to [person] about [topic]. Save it to my notes first so I can review."

### Channel monitoring
> "Check the [channel name] channel for recent activity and summarise any important updates."

### Meeting transcript
> "Get the transcript from my last standup meeting and summarise the key decisions."

### Search emails
> "Search my emails for anything from [person] about [topic] in the last week."

### People lookup
> "Find [person name]'s contact details and check if I have any recent conversations with them."
