# Teams MCP User Stories

This document defines user stories and personas to guide development of the Teams MCP. Each story maps to specific API capabilities needed.

---

## Personas

### 🧑‍💼 Alex - Busy Manager
- Receives 100+ messages daily across multiple channels
- Needs to quickly catch up on what matters
- Often works across time zones, misses real-time conversations
- Wants AI to help prioritise and respond

### 👩‍💻 Sam - Developer
- Part of 10+ project channels
- Gets tagged in technical discussions
- Needs to find past decisions and context quickly
- Wants to automate routine responses

### 🧑‍🎨 Jordan - Creative Lead
- Collaborates with multiple teams
- Shares files and feedback frequently
- Needs to track project updates across channels
- Wants summaries rather than reading everything

---

## User Stories

### 1. Search & Reply

#### 1.1 Find and reply to a message
> "Find the message from Sarah about the budget review and reply saying I'll review it tomorrow."

**Flow:**
1. Search for messages matching "budget review from:sarah"
2. Display results with context
3. User confirms which message
4. Send reply to that conversation

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented (returns conversationId, messageId) |
| `teams_send_message` | ✅ Implemented - with replyToMessageId for thread replies |
| `teams_get_thread` | ✅ Implemented - get surrounding messages |

**Status:** ✅ Implemented - search returns `conversationId` and `messageId`, use `teams_send_message` with `replyToMessageId` to reply to channel threads, or `teams_get_thread` for context first.

---

#### 1.2 Search with date filters
> "Find messages from last week mentioning 'deployment'"

**Flow:**
1. Search with `deployment` (results are sorted by recency, so recent messages appear first)
2. Or use explicit date: `sent:>=2026-01-18 deployment`
3. Return matching messages

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented (explicit date operators work) |

**Status:** ✅ Implemented. Explicit dates only. Note: `sent:lastweek` does NOT work - use `sent:>=YYYY-MM-DD` or omit.

---

### 2. Catch Up & Prioritise

#### 2.1 Review questions asked of me
> "Show me any questions people have asked me today that I haven't answered."

**Flow:**
1. Search for messages mentioning me with question marks
2. Filter to unanswered (no reply from me after)
3. Prioritise by sender importance/urgency

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_get_me` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented - check if I replied |

**Status:** ✅ Implemented - search for mentions with "?", then use `teams_get_thread` on each result to check if you've replied. AI can filter to show only unanswered.

---

#### 2.2 Catch up on unread messages
> "What unread messages do I have?"

**Flow:**
1. Get list of conversations with unread counts via `teams_get_unread`
2. Fetch unread messages from each using `teams_get_thread`
3. Optionally mark as read with `teams_mark_read` or `teams_get_thread` with `markRead: true`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_unread` | ✅ Implemented (aggregate or per-conversation) |
| `teams_get_thread` | ✅ Implemented (with optional markRead flag) |
| `teams_mark_read` | ✅ Implemented |

**Status:** ✅ Implemented - can check unread counts across favourites or for specific conversations, then read and mark as read.

---

#### 2.3 Catch up on a channel
> "Summarise what happened in #project-alpha today"

**Flow:**
1. Find channel by name using `teams_find_channel`
2. Get recent messages from channel using `teams_get_thread`
3. Generate summary

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_find_channel` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented (works with channel IDs) |

**Status:** ✅ Implemented - use `teams_find_channel` to discover channels by name, then `teams_get_thread` with the returned `channelId` to get messages.

---

#### 2.4 Check for replies to my message
> "Have there been any replies to my PR review request message?"

**Flow:**
1. Search for the original message to get its `conversationId`
2. Call `teams_get_thread` to get all messages in that conversation
3. Display replies after the original message timestamp

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented (returns conversationId) |
| `teams_get_thread` | ✅ Implemented |

**Status:** ✅ Implemented - search returns `conversationId`, then `teams_get_thread` retrieves all messages in that thread.

**Note:** Reactions (👍) are still not surfaced by this API. Only actual message replies are returned.

---

#### 2.5 Morning catch-up routine
> "What do I need to know this morning?"

**Flow:**
1. Check activity feed for mentions, reactions, replies via `teams_get_activity`
2. Check unread counts across favourites via `teams_get_unread`
3. Optionally summarise key threads using `teams_get_thread`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_activity` | ✅ Implemented |
| `teams_get_unread` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented |

**Status:** ✅ Implemented - AI can combine activity feed + unread status to provide a comprehensive morning summary.

---

#### 2.6 Cross-channel summary
> "What's happening across all my project channels?"

**Flow:**
1. Get favourites or find multiple channels via `teams_find_channel`
2. Get recent messages from each using `teams_get_thread`
3. AI summarises activity across all channels

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_find_channel` | ✅ Implemented |
| `teams_get_favorites` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented |

**Status:** ✅ Implemented - requires multiple API calls but AI can orchestrate.

---

### 3. Favourites & Navigation

#### 3.1 List favourite channels
> "Show me my pinned/favourite channels"

**Flow:**
1. Get user's favourite channels list
2. Display with recent activity indicator

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_favorites` | ✅ Implemented |
| `teams_add_favorite` | ✅ Implemented |
| `teams_remove_favorite` | ✅ Implemented |

**Status:** ✅ Implemented - can list, add, and remove favourites via the conversationFolders API.

---

#### 3.2 List recent chats
> "Who have I been chatting with recently?"

**Flow:**
1. Get recent 1:1 and group chats
2. Show with last message preview

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_recent_chats` | ❌ Needed |

**Gap:** Chat list loaded at startup. No dedicated API endpoint.

**Partial Workarounds:**
- `teams_get_frequent_contacts` shows who you frequently interact with
- `teams_get_favorites` shows your pinned conversations
- Together these cover most "who do I chat with" use cases, but won't show true recency ordering

---

### 4. People & Profiles

#### 4.1 Find and message someone
> "Send a message to John Smith asking about the project status"

**Flow:**
1. Search for person by name using `teams_search_people`
2. Get their conversation ID using `teams_get_chat`
3. Send message using `teams_send_message`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_people` | ✅ Implemented |
| `teams_get_chat` | ✅ Implemented |
| `teams_send_message` | ✅ Implemented |

**Status:** ✅ Implemented - can find anyone and start a new 1:1 chat with them.

**Technical Note:** The conversation ID for 1:1 chats follows a predictable format: `19:{id1}_{id2}@unq.gbl.spaces` where the two user object IDs are sorted lexicographically. The `teams_get_chat` tool computes this ID from the user's object ID (from people search). The conversation is created implicitly when the first message is sent.

---

#### 4.2 Create a group chat
> "Start a group chat with John and Sarah about the project"

**Flow:**
1. Search for people using `teams_search_people`
2. Create group chat using `teams_create_group_chat` with their MRIs
3. Send initial message using `teams_send_message`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_people` | ✅ Implemented |
| `teams_create_group_chat` | ✅ Implemented |
| `teams_send_message` | ✅ Implemented |

**Status:** ✅ Implemented - uses `POST /api/chatsvc/{region}/v1/threads` endpoint.

---

#### 4.3 Check someone's availability
> "Is Sarah available for a call right now?"

**Flow:**
1. Find person
2. Get their presence/availability status

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_people` | ✅ Implemented |
| `teams_get_presence` | ❌ Needed (WebSocket-based) |

**Gap:** People search works, but presence/availability is real-time via WebSocket, not HTTP API.

---

#### 4.4 Get my profile
> "What's my Teams email address?"

**Flow:**
1. Get current user profile

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_me` | ✅ Implemented |

**Status:** ✅ Implemented - returns `id`, `mri`, `email`, `displayName`, `tenantId`.

---

### 5. Notifications & Activity

#### 5.1 Review activity feed
> "Show me my recent notifications"

**Flow:**
1. Get activity/notification feed via `teams_get_activity`
2. Display with context (mentions, reactions, replies)

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_activity` | ✅ Implemented |

**Status:** ✅ Implemented - returns mentions, reactions, replies with direct links to open in Teams.

---

#### 5.2 Find mentions of me
> "Show messages where I was @mentioned this week"

**Flow:**
1. Get user's display name via `teams_get_me`
2. Search for `"Display Name" NOT from:email` (results sorted by recency, or add `sent:>=YYYY-MM-DD`)

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_get_me` | ✅ Implemented |

**Status:** ✅ Implemented using search operators with user's display name.

---

### 6. Reactions & Engagement

#### 6.1 React to a message
> "Give a thumbs up to Sarah's message about the release"

**Flow:**
1. Search for the message to get `conversationId` and `messageId`
2. Add reaction using `teams_add_reaction` with emoji key

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_add_reaction` | ✅ Implemented |

**Status:** ✅ Implemented - supports standard reactions (`like`, `heart`, `laugh`, `surprised`, `sad`, `angry`) plus any emoji by key.

---

#### 6.2 Remove a reaction
> "Remove my like from that message"

**Flow:**
1. Find the message via search
2. Remove reaction using `teams_remove_reaction`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_remove_reaction` | ✅ Implemented |

**Status:** ✅ Implemented.

---

#### 6.3 Find custom emojis
> "What custom emojis does our org have?"

**Flow:**
1. Search for emojis by keyword using `teams_search_emoji`
2. Returns both standard and organisation-specific custom emojis

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_emoji` | ✅ Implemented |

**Status:** ✅ Implemented - searches standard Teams emojis and org-specific custom emojis.

---

### 7. Message Management

#### 7.1 Edit a message
> "Fix the typo in my last message - change 'teh' to 'the'"

**Flow:**
1. Get recent messages from the conversation via `teams_get_thread`
2. Find your message that needs editing
3. Edit using `teams_edit_message` with corrected content

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_thread` | ✅ Implemented |
| `teams_edit_message` | ✅ Implemented |

**Status:** ✅ Implemented - can only edit your own messages.

---

#### 7.2 Delete a message
> "Delete that message I just sent"

**Flow:**
1. Get recent messages from the conversation via `teams_get_thread`
2. Find your message
3. Delete using `teams_delete_message`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_thread` | ✅ Implemented |
| `teams_delete_message` | ✅ Implemented |

**Status:** ✅ Implemented - can only delete your own messages (soft delete).

---

#### 7.3 Send a note to yourself
> "Make a note to myself about the Friday deadline"

**Flow:**
1. Send message using `teams_send_message` with default `conversationId` (48:notes)

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_send_message` | ✅ Implemented |

**Status:** ✅ Implemented - defaults to self-chat when no `conversationId` provided.

---

#### 7.4 Save a message for later
> "Bookmark that important message about the API changes"

**Flow:**
1. Search for the message to get `conversationId` and `messageId`
2. Save using `teams_save_message`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_save_message` | ✅ Implemented |
| `teams_unsave_message` | ✅ Implemented |

**Status:** ✅ Implemented.

**Note:** Saved messages can be retrieved via `teams_get_saved_messages`. This returns references with source conversation IDs - use `teams_get_thread` to fetch full content.

---

#### 7.5 View saved messages
> "Show me all my bookmarked messages"

**Flow:**
1. Get saved messages via `teams_get_saved_messages`
2. Optionally fetch full content from source using `teams_get_thread`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_saved_messages` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented (for full content) |

**Status:** ✅ Implemented - returns references with `sourceConversationId` and `messageLink`.

---

#### 7.6 View followed threads
> "What threads am I following?"

**Flow:**
1. Get followed threads via `teams_get_followed_threads`
2. Optionally fetch full thread content using `teams_get_thread`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_followed_threads` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented (for full content) |

**Status:** ✅ Implemented - returns references with `sourceConversationId` and `messageLink`.

---

### 8. Files & Attachments

#### 8.1 Find shared files
> "Find the Excel file Sarah shared last week"

**Flow:**
1. Search for file by name/sender
2. Return download link or preview

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_files` | ❌ Needed |
| `teams_get_shared_files` | ❌ Needed (AllFiles API available) |

**Status:** API discovered, implementation pending.

**Workaround:** Use story 8.2 below - search for messages with attachments using `hasattachment:true from:sarah` to find messages where Sarah shared files.

---

#### 8.2 Find messages with attachments
> "Find messages where someone shared a file about the project"

**Flow:**
1. Search using `hasattachment:true` operator combined with keywords
2. Results show messages that have files attached

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |

**Status:** ✅ Implemented - use `hasattachment:true` operator to filter.

**Example:** `hasattachment:true from:sarah project proposal`

---

#### 8.3 Extract links from messages
> "Find messages with links about the documentation"

**Flow:**
1. Search for messages mentioning documentation
2. Results include a `links` array with extracted URLs and display text
3. AI can present or act on the links

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented |

**Status:** ✅ Implemented - all message responses include a `links` field (when URLs are present) containing `{ url, text }` objects. Works across search results, threads, saved messages, followed threads, and activity feed.

---

### 9. Calendar & Meetings

#### 9.1 Check next meeting
> "When's my next meeting?"

**Flow:**
1. Get upcoming meetings via `teams_get_meetings` (default: next 7 days)
2. Return the first meeting with time, subject, and join link

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_meetings` | ✅ Implemented |

**Status:** ✅ Implemented.

---

#### 9.2 Count meetings for a day
> "How many meetings do I have tomorrow?"

**Flow:**
1. Get meetings for tomorrow via `teams_get_meetings` with specific date range
2. Return count and optionally list subjects

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_meetings` | ✅ Implemented |

**Status:** ✅ Implemented.

---

#### 9.3 Get meeting summary (recent meeting)
> "Get me a summary of my last meeting"

**Flow:**
1. Get recent meetings via `teams_get_meetings` with past date range
2. Get the meeting thread ID from the response
3. If transcription was enabled: fetch transcript via `teams_get_transcript` using `threadId`
4. Otherwise: fetch chat messages via `teams_get_thread` using the thread ID
5. AI summarises the meeting from transcript or chat

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_meetings` | ✅ Implemented |
| `teams_get_transcript` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented (fallback for meetings without transcription) |

**Status:** ✅ Implemented - set `startDate` to past to get recent meetings. Use `teams_get_transcript` for full transcript with speakers/timestamps, or `teams_get_thread` for meeting chat.

---

#### 9.4 Find meetings with a person
> "When's the next meeting with Dave?"

**Flow:**
1. Find Dave's email via `teams_search_people`
2. Get upcoming meetings via `teams_get_meetings`
3. Filter to meetings where Dave is organiser

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_people` | ✅ Implemented |
| `teams_get_meetings` | ✅ Implemented |

**Status:** ✅ Implemented (organiser filter only). Note: Currently filters by organiser email. Attendee list filtering requires additional API research.

---

#### 9.5 Get meeting chat context
> "What was discussed in yesterday's standup meeting chat?"

**Flow:**
1. Get meeting via `teams_get_meetings` with yesterday's date range
2. Get the conversation thread using `teams_get_thread` with the meeting's `threadId`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_meetings` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented |

**Status:** ✅ Implemented - `teams_get_meetings` returns `threadId` for each meeting.

---

#### 9.6 Today's schedule overview
> "What's my schedule for today?"

**Flow:**
1. Get today's meetings via `teams_get_meetings` with today's date range
2. AI presents as a timeline with gaps shown

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_meetings` | ✅ Implemented |

**Status:** ✅ Implemented.

---

#### 9.7 Prepare for a meeting
> "Help me prepare for my 2pm meeting"

**Flow:**
1. Get today's meetings via `teams_get_meetings`, find the 2pm one
2. Get recent messages in that meeting's chat via `teams_get_thread`
3. Search for related messages from the organiser
4. AI compiles context and prep notes

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_meetings` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented |
| `teams_search` | ✅ Implemented |

**Status:** ✅ Implemented.

---

#### 9.8 Get meeting transcript
> "Get the transcript from yesterday's standup"

**Flow:**
1. Get recent meetings via `teams_get_meetings` with past date range
2. Identify the target meeting, get its `threadId` and `startTime`
3. Fetch transcript via `teams_get_transcript` with `threadId` and optionally `meetingDate`
4. AI presents or summarises the transcript

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_meetings` | ✅ Implemented |
| `teams_get_transcript` | ✅ Implemented |

**Status:** ✅ Implemented - returns formatted transcript with speaker names, timestamps, and spoken text. Only available for meetings where transcription was enabled.

---

#### 9.9 Check free time
> "When am I free this afternoon?"

**Flow:**
1. Get today's meetings via `teams_get_meetings`
2. AI identifies gaps between meetings

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_meetings` | ✅ Implemented |

**Status:** ✅ Implemented.

---

### 10. Advanced AI Workflows

These patterns combine multiple tools for sophisticated interactions.

#### 10.1 Forward a message to someone
> "Forward John's budget update to Sarah"

**Flow:**
1. Search for John's message about budget
2. Find Sarah using `teams_search_people`
3. Get chat ID using `teams_get_chat`
4. Send message with quoted content using `teams_send_message`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search` | ✅ Implemented |
| `teams_search_people` | ✅ Implemented |
| `teams_get_chat` | ✅ Implemented |
| `teams_send_message` | ✅ Implemented |

**Status:** ✅ Implemented - AI composes the forward by quoting original message content.

---

#### 10.2 Draft and send a message
> "Help me write a message to the team about the delayed release"

**Flow:**
1. AI drafts message content based on user's intent
2. User reviews and approves
3. Find channel/chat using `teams_find_channel` or `teams_get_chat`
4. Send using `teams_send_message`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_find_channel` | ✅ Implemented |
| `teams_get_chat` | ✅ Implemented |
| `teams_send_message` | ✅ Implemented |

**Status:** ✅ Implemented - AI can draft, user confirms, then send.

---

#### 10.3 Respond to all pending questions
> "Reply to all questions I haven't answered today"

**Flow:**
1. Get user info via `teams_get_me`
2. Search for mentions with questions: `"Display Name" ? NOT from:email`
3. For each result, check thread via `teams_get_thread` for your reply
4. Draft and send replies to unanswered questions

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_get_me` | ✅ Implemented |
| `teams_search` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented |
| `teams_send_message` | ✅ Implemented |

**Status:** ✅ Implemented - AI orchestrates the full workflow.

---

#### 10.4 Check in on a person's activity
> "What's Sarah been up to lately?"

**Flow:**
1. Find Sarah's email via `teams_search_people`
2. Search for messages from her: `from:sarah@company.com`
3. Optionally check frequent contacts to see if you interact often

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_search_people` | ✅ Implemented |
| `teams_search` | ✅ Implemented |
| `teams_get_frequent_contacts` | ✅ Implemented |

**Status:** ✅ Implemented.

---

#### 10.5 Pin important conversation after finding it
> "Find the project-alpha channel and pin it to my favourites"

**Flow:**
1. Find channel via `teams_find_channel`
2. Add to favourites using `teams_add_favorite`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_find_channel` | ✅ Implemented |
| `teams_add_favorite` | ✅ Implemented |

**Status:** ✅ Implemented.

---

#### 10.6 Acknowledge messages quickly
> "React with thumbs up to the last 3 messages in the support channel"

**Flow:**
1. Find channel via `teams_find_channel`
2. Get recent messages via `teams_get_thread`
3. Add reactions to each using `teams_add_reaction`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_find_channel` | ✅ Implemented |
| `teams_get_thread` | ✅ Implemented |
| `teams_add_reaction` | ✅ Implemented |

**Status:** ✅ Implemented.

---

#### 10.7 Mention someone in a message
> "Send a message to the project channel mentioning Sarah about the deadline"

**Flow:**
1. Find channel via `teams_find_channel`
2. Find Sarah's MRI via `teams_search_people`
3. Compose message with inline mention: `@[Sarah Smith](8:orgid:abc...)`
4. Send using `teams_send_message`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_find_channel` | ✅ Implemented |
| `teams_search_people` | ✅ Implemented |
| `teams_send_message` | ✅ Implemented |

**Status:** ✅ Implemented - use `@[DisplayName](mri)` syntax in message content. The MRI comes from people search results.

---

#### 10.8 Include links in a message
> "Send a message with a link to the project documentation"

**Flow:**
1. Compose message with markdown-style link: `[text](url)`
2. Send using `teams_send_message`

**Required Tools:**
| Tool | Status |
|------|--------|
| `teams_send_message` | ✅ Implemented |

**Status:** ✅ Implemented - use `[display text](https://url)` syntax in message content. Multiple links and combinations with @mentions work.

**Example:**
```
teams_send_message content="Check out [the docs](https://example.com/docs) for details"
```

---

## Implementation Priority

Based on user value and API readiness:

### Phase 1 - Core Messaging ✅ Complete
| Story | Tools | Status |
|-------|-------|--------|
| 1.1 Find & reply | `teams_search`, `teams_send_message` | ✅ Done |
| 1.2 Search with filters | `teams_search` operators | ✅ Done |
| 4.1 Find and message someone | `teams_search_people`, `teams_get_chat`, `teams_send_message` | ✅ Done |
| 4.3 Get my profile | `teams_get_me` | ✅ Done |
| 7.3 Send notes to yourself | `teams_send_message` default | ✅ Done |

### Phase 2 - Catch-Up & Context ✅ Complete
| Story | Tools | Status |
|-------|-------|--------|
| 2.1 Review questions asked of me | `teams_search`, `teams_get_thread` | ✅ Done |
| 2.2 Catch up on unread | `teams_get_unread`, `teams_mark_read` | ✅ Done |
| 2.3 Channel catchup | `teams_find_channel`, `teams_get_thread` | ✅ Done |
| 2.4 Check for replies | `teams_get_thread` | ✅ Done |
| 2.5 Morning catch-up | `teams_get_activity`, `teams_get_unread` | ✅ Done |
| 5.1 Activity feed | `teams_get_activity` | ✅ Done |
| 5.2 Find @mentions | Search operators | ✅ Done |

### Phase 3 - Organisation & Engagement ✅ Complete
| Story | Tools | Status |
|-------|-------|--------|
| 3.1 Favourites | `teams_get_favorites`, `teams_add_favorite`, `teams_remove_favorite` | ✅ Done |
| 6.1 React to messages | `teams_add_reaction`, `teams_remove_reaction` | ✅ Done |
| 6.3 Search emojis | `teams_search_emoji` | ✅ Done |
| 7.1 Edit messages | `teams_edit_message` | ✅ Done |
| 7.2 Delete messages | `teams_delete_message` | ✅ Done |
| 7.4 Save messages | `teams_save_message`, `teams_unsave_message` | ✅ Done |

### Phase 4 - Meetings & Calendar ✅ Complete
| Story | Tools | Status |
|-------|-------|--------|
| 9.1 Check next meeting | `teams_get_meetings` | ✅ Done |
| 9.2 Count meetings | `teams_get_meetings` | ✅ Done |
| 9.3 Meeting summary | `teams_get_meetings`, `teams_get_thread` | ✅ Done |
| 9.4 Meetings with person | `teams_get_meetings`, `teams_search_people` | ✅ Done (organiser only) |
| 9.5-9.7, 9.9 Other meeting stories | `teams_get_meetings` | ✅ Done |
| 9.8 Meeting transcript | `teams_get_meetings`, `teams_get_transcript` | ✅ Done |

### Phase 5 - Remaining Gaps
| Story | Tools Needed | Effort |
|-------|-------------|--------|
| 3.2 Recent chats list | No dedicated API | Blocked |
| 4.2 Presence/availability | WebSocket (not HTTP) | Very High |
| 8.1 Find shared files | AllFiles API | Medium |

---

## Implementation Status

### All Core Features Complete ✅

The following tools are implemented:

**Search & Discovery:**
- `teams_search` - Full-text search with operators
- `teams_find_channel` - Find channels by name
- `teams_search_people` - Find people by name/email
- `teams_get_frequent_contacts` - Ranked frequent contacts

**Messaging:**
- `teams_send_message` - Send messages (chats, channels, self-notes, thread replies via `replyToMessageId`)
- `teams_get_thread` - Read conversation messages
- `teams_edit_message` - Edit your own messages
- `teams_delete_message` - Delete your own messages

**Reactions:**
- `teams_add_reaction` - React to messages
- `teams_remove_reaction` - Remove reactions
- `teams_search_emoji` - Find emoji keys

**Organisation:**
- `teams_get_favorites` / `teams_add_favorite` / `teams_remove_favorite`
- `teams_save_message` / `teams_unsave_message` / `teams_get_saved_messages`
- `teams_get_followed_threads`
- `teams_get_unread` / `teams_mark_read`

**Profile & Status:**
- `teams_get_me` - Current user profile
- `teams_get_activity` - Activity feed
- `teams_status` - Auth status check
- `teams_login` - Manual login

**Chat Management:**
- `teams_get_chat` - Get 1:1 conversation ID

**Calendar & Meetings:**
- `teams_get_meetings` - Get meetings from calendar
- `teams_get_transcript` - Get meeting transcript (speakers, timestamps, text)

### Remaining Gaps

| Feature | Blocker | Notes |
|---------|---------|-------|
| `teams_get_files` | Medium effort | AllFiles API discovered, implementation pending |
| Recent chats list | No dedicated API | Use `teams_get_favorites` + `teams_get_frequent_contacts` as workaround |
| Presence/availability | WebSocket only | Real-time presence not available via HTTP |

---

## Notes

### Search Operators (Supported)
```
from:john.smith@company.com    # Messages from person (use actual email)
in:general                     # Messages in channel
sent:2026-01-20                # Messages from specific date
sent:>=2026-01-15              # Messages since date
hasattachment:true             # Messages with files
"Display Name"                 # Find @mentions (use actual display name)
NOT from:email                 # Exclude results
```

**⚠️ Does NOT work:** `@me`, `from:me`, `to:me`, `mentions:me` - use `teams_get_me` first to get actual email/displayName. Also `sent:lastweek` and `sent:thisweek` do NOT work - use explicit dates or omit (results sorted by recency). `sent:today` works.

Combine operators: `from:sarah@co.com sent:>=2026-01-18 hasattachment:true`

### Conversation IDs
- `48:notes` - Self-chat (notes to yourself)
- `48:notifications` - Activity feed
- `19:xxx@thread.tacv2` - Channel conversation
- `19:xxx@unq.gbl.spaces` - 1:1 or group chat
