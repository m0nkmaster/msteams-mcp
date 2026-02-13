# [SPIKE] Drop Internal APIs — Move to Graph + Outlook REST

> **Status: Investigation only.** This documents findings from testing Microsoft Graph API scopes. The current architecture still uses Substrate/chatsvc/Skype internal APIs.

Migrate from Substrate/chatsvc/Skype internal APIs to Microsoft Graph + Outlook REST, using `outlook.office.com` SSO (client ID `9199bf20-a13f-4107-85dc-02114787ef48`) which gives `Chat.Read` + `Chat.ReadWrite` on Graph.

## SSO Target Comparison (All Live-Tested)

| Graph Scope | Teams (`5e3ce6c0`) | Outlook (`9199bf20`) | M365 Chat (`c0ab8ce9`) | M365 Apps (`4765445b`) |
|---|---|---|---|---|
| **Chat.ReadWrite** | ❌ | ✅ | ❌ | ❌ |
| **Chat.Read** | ❌ | ✅ | ChatReadBasic | ❌ |
| **ChatMessage.Send** | ✅ | ❌ | ✅ | ❌ |
| **Calendars** | ReadWrite | ❌ | Read | Read |
| **Mail** | Read/ReadWrite | ❌ | MailboxSettings | ❌ |
| **Presence** | ❌ | ❌ | ❌ | Read.All ✅ |
| **Files** | ReadWrite.All | ReadWrite.All | ReadWrite.All | ReadWrite.All |
| **People** | Read | Read | Read | Read |
| **OnlineMeetings** | ❌ | ReadWrite | ❌ | ❌ |

**Winner: Outlook** — only SSO target with `Chat.Read` + `Chat.ReadWrite` on Graph.

Outlook web client ID (`9199bf20-a13f-4107-85dc-02114787ef48`) provides two tokens:
- **Graph token** — `Chat.Read`, `Chat.ReadWrite`, `Files.ReadWrite.All`, `People.Read`, `OnlineMeetings.ReadWrite`, etc.
- **Outlook REST token** (`outlook.office.com` audience) — `Calendars.ReadWrite`, `Mail.ReadWrite`, `Contacts.ReadWrite`, `SubstrateSearch-Internal.ReadWrite`, `ChannelMessage.Read.All`, etc.

## Full Feasibility Matrix (All Live-Tested)

### ✅ Can Replace with Graph/Outlook REST (17 tools)

| Current Tool | Current API | Graph/Outlook Replacement | Tested |
|---|---|---|---|
| `teams_send_message` | chatsvc | `POST /chats/{id}/messages` | ✅ 201 |
| `teams_get_thread` | chatsvc | `GET /chats/{id}/messages` | ✅ 200 |
| `teams_edit_message` | chatsvc | `PATCH /chats/{id}/messages/{id}` | ✅ 204 |
| `teams_add_reaction` | chatsvc | `POST .../setReaction` | ✅ 204 |
| `teams_remove_reaction` | chatsvc | `POST .../unsetReaction` | ✅ 204 |
| `teams_get_chat` | chatsvc | `POST /chats` (oneOnOne) | ✅ 201 |
| `teams_create_group_chat` | chatsvc | `POST /chats` (group) | ✅ 201 |
| `teams_search_people` | Substrate | `GET /me/people?$search=` | ✅ 200 |
| `teams_get_me` | Substrate | `GET /me` | ✅ 200 |
| `teams_get_frequent_contacts` | Substrate | `GET /me/people` | ✅ 200 |
| `teams_find_channel` | Substrate | `GET /me/joinedTeams` + `/channels` | ✅ 200 |
| `teams_get_meetings` | Calendar API | Outlook REST `calendarview` | ✅ 200 |
| `teams_get_shared_files` | Substrate | `GET /chats/{id}/messages` (attachments) | ✅* |
| `teams_get_favorites` | CSA | `GET /me/chats?$select=viewpoint` (isPinned) | ✅ 200 |
| `teams_get_unread` | chatsvc | `GET /me/chats` (viewpoint.lastMessageReadDateTime) | ✅ 200 |
| `teams_login` | Browser | Same (Outlook SSO instead of Teams) | ✅ |
| `teams_status` | Internal | Same (check Outlook tokens) | ✅ |

*Files: Graph returns messages with attachments; need to parse differently than Substrate AllFiles.

### ❌ Cannot Replace — Must Keep or Drop (7 tools)

| Current Tool | Current API | Graph Status | Recommendation |
|---|---|---|---|
| `teams_search` | Substrate | 403 (chatMessage search needs `ChannelMessage.Read.All` on Graph token — we have it on Outlook token but not Graph) | **Keep Substrate** OR use Outlook REST token for search |
| `teams_delete_message` | chatsvc | 405 (softDelete not working) | **Keep chatsvc** |
| `teams_mark_read` | chatsvc | 405 (markChatReadForUser) | **Keep chatsvc** |
| `teams_get_activity` | chatsvc | No Graph equivalent | **Keep chatsvc** |
| `teams_get_saved_messages` | chatsvc | No Graph equivalent | **Keep chatsvc** |
| `teams_get_followed_threads` | chatsvc | No Graph equivalent | **Keep chatsvc** |
| `teams_save_message` / `unsave` | chatsvc | No Graph equivalent | **Keep chatsvc** |
| `teams_search_emoji` | Client-side | No Graph equivalent | **Keep client-side** |
| `teams_get_transcript` | Substrate | Graph needs meeting ID lookup | **Investigate** |
| `teams_add_favorite` / `remove` | CSA | No pin/unpin API | **Keep CSA** |

### Search: The Biggest Gap

Graph Search (`POST /search/query` with `chatMessage`) returns 403 even with `Chat.Read` — it additionally requires `ChannelMessage.Read.All` **on the Graph token**. Our Outlook Graph token doesn't have that scope (it's on the Outlook REST token instead).

**Options**:
1. Keep Substrate search (proven, fast) — requires Teams SSO
2. Use Outlook REST token for Substrate search (has `SubstrateSearch-Internal.ReadWrite`) — **needs testing**
3. Iterate through chats + filter messages client-side (slow, impractical)

### Transcript: Needs Investigation

Currently uses Substrate `WorkingSetFiles` API. Graph has `GET /me/onlineMeetings/{id}/transcripts` but requires the meeting ID, which needs a lookup from `threadId`. May be feasible but needs work.

## Verdict: Can We Drop Teams SSO Entirely?

**Yes — with one caveat.** The Outlook refresh token can acquire tokens for **all** APIs:

| Token | Outlook Refresh → | Tested |
|---|---|---|
| Graph (`Chat.ReadWrite`) | ✅ | Confirmed |
| Substrate | ✅ acquired | Search returns 400 — **needs investigation** |
| Skype Spaces | ✅ | Confirmed |
| skypetoken_asm | ✅ (via Skype exchange) | Confirmed |
| chatsvcagg | ✅ | Confirmed |

This means **single Outlook SSO** can power everything — Graph for chat ops, chatsvc for the 7 features with no Graph equivalent, and Substrate for search. The only open issue is Substrate search returning 400 with the Outlook-refreshed token (may be a request format or scope difference).

## Architecture: Outlook-Only SSO

```
Outlook SSO (outlook.office.com) — client ID 9199bf20
├── Outlook refresh token (captured during SSO via window.msal or token endpoint)
│   ├── → Graph token (Chat.ReadWrite) → send/list/edit messages, reactions, create chat, people, teams/channels
│   ├── → Outlook REST token → calendar/meetings, contacts
│   ├── → Substrate token → search, transcripts (needs 400 fix)
│   ├── → Skype Spaces token → skypetoken_asm exchange
│   └── → chatsvcagg token → chatsvc auth
├── skypetoken_asm (derived from Skype Spaces)
│   └── → delete, mark read, activity feed, saved msgs, save/unsave, favorites
└── No Teams SSO needed!
```

### Migration Strategy

**Phase 1**: Migrate chat tools to Graph (send, list, edit, reactions, create chat) — biggest impact, uses Outlook Graph token directly.

**Phase 2**: Switch SSO from `teams.microsoft.com` to `outlook.office.com`. Capture Outlook refresh token from `window.msal` during login. Use it to refresh all token types.

**Phase 3**: Investigate and fix Substrate search with Outlook-refreshed token. If unfixable, keep Substrate search on Teams token as fallback.

**Phase 4**: Remove Teams SSO dependency (if Phase 3 succeeds).

## Implementation Steps (if proceeding)

### Phase 1: Outlook SSO + Token Management
1. Switch login target from `teams.microsoft.com` to `outlook.office.com`
2. Capture Graph + Outlook REST tokens from network requests (not localStorage)
3. Add Outlook client ID to HTTP token refresh (`Origin: https://outlook.office.com`)
4. Add `requireOutlookGraphAuth()` and `requireOutlookRestAuth()` guards

### Phase 2: Graph API Module
5. Create `graph-api.ts` — send, list, edit messages, reactions, create chat
6. Create `graph-people.ts` — search people, get me, frequent contacts
7. Create `graph-teams.ts` — list teams, channels, find channel

### Phase 3: Outlook REST Module
8. Create `outlook-calendar.ts` — meetings, calendar view
9. Evaluate if Outlook REST token works for Substrate search

### Phase 4: Tool Migration
10. Migrate all 17 tools to Graph/Outlook REST
11. Keep chatsvc fallback for the 7 remaining features
12. Update tool descriptions

### Phase 5: Cleanup
13. Remove Substrate API calls (if search migrated)
14. Remove chatsvc calls for migrated tools
15. Update AGENTS.md, tests
16. Build + test end-to-end
