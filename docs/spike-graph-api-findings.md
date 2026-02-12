# Spike: Microsoft Graph API for Teams MCP

**Date:** 2026-02-12
**Branch:** `spike/graph-api-send-message`
**Status:** Spike complete — Graph API confirmed working

## Summary

We investigated whether the Microsoft Graph API can be used with tokens extracted from the Teams browser session (no Azure App registration required). The answer is **yes** — the Teams SPA client ID (`5e3ce6c0-2b1f-4285-8d4b-75ee78787346`) has delegated Graph API permissions that allow sending messages and likely much more.

## What We Built

Added `graph.microsoft.com` as a 4th scope to the HTTP token refresh mechanism. The MSAL refresh token is exchanged for a Graph access token alongside the existing Substrate, Skype Spaces, and chatsvcagg tokens.

### Files Changed

- `src/auth/token-refresh-http.ts` — Added `graph.microsoft.com` to `REFRESH_SCOPES`
- `src/auth/token-extractor.ts` — Added `extractGraphToken()`, `getValidGraphToken()`, `getGraphTokenStatus()`
- `src/utils/auth-guards.ts` — Added `requireGraphAuth()` guard and `GraphAuthInfo` type
- `src/api/graph-api.ts` *(new)* — `graphSendMessage()` and `graphSendChannelMessage()`
- `src/tools/graph-tools.ts` *(new)* — Two spike tools: `teams_graph_send_message`, `teams_graph_token_status`
- `src/tools/registry.ts` — Registered graph tools
- `src/tools/index.ts` — Re-exported graph tools
- `src/auth/token-refresh-http.test.ts` — Updated test expectations for 4 scopes (was 3)

## Test Result

Successfully sent a message to self-notes (`48:notes`) via Graph API:

```
POST https://graph.microsoft.com/v1.0/chats/48%3Anotes/messages
Authorization: Bearer <graph-token>
Content-Type: application/json

{ "body": { "contentType": "text", "content": "Test message from Graph API spike" } }
```

Response confirmed: message ID returned, `from.user` populated with display name and tenant ID, full `chatMessage` resource returned.

## Graph API Coverage vs. Current APIs

### ✅ Graph API CAN Replace (single token)

| Current Tool | Current API | Graph API Equivalent |
|---|---|---|
| `teams_send_message` | chatsvc (skypetoken_asm) | `POST /chats/{id}/messages` — **confirmed working** |
| `teams_edit_message` | chatsvc | `PATCH /chats/{id}/messages/{id}` |
| `teams_delete_message` | chatsvc | `POST /chats/{id}/messages/{id}/softDelete` |
| `teams_get_thread` | chatsvc | `GET /chats/{id}/messages` |
| `teams_add_reaction` | chatsvc | `POST /chats/{id}/messages/{id}/setReaction` |
| `teams_remove_reaction` | chatsvc | `POST /chats/{id}/messages/{id}/unsetReaction` |
| `teams_get_chat` | chatsvc (ID construction) | `POST /chats` (create) or `GET /me/chats` (list) |
| `teams_create_group_chat` | chatsvc | `POST /chats` with members |
| `teams_get_meetings` | mt/part (Skype Spaces token) | `GET /me/calendarView` |
| `teams_search` | Substrate v2/query | `POST /search/query` with `entityTypes: ["chatMessage"]` |
| `teams_search_people` | Substrate v1/suggestions | `GET /me/people` or `POST /search/query` with `entityTypes: ["person"]` |
| `teams_get_frequent_contacts` | Substrate v1/suggestions | `GET /me/people` (ranked by relevance) |
| `teams_get_me` | JWT parsing | `GET /me` |
| `teams_get_transcript` | Substrate WorkingSetFiles | `GET /me/onlineMeetings/{id}/transcripts` |
| `teams_find_channel` | Substrate v1/suggestions | `GET /me/joinedTeams` + `GET /teams/{id}/channels` |
| `teams_get_shared_files` | Substrate AllFiles | `GET /chats/{id}/messages` (filter attachments) or OneDrive API |

### ⚠️ Partial or Different Coverage

| Current Tool | Issue with Graph |
|---|---|
| `teams_get_favorites` | No direct "pinned conversations" API. `GET /me/chats` returns all chats but no pin/favourite status. Would need CSA still. |
| `teams_add/remove_favorite` | No Graph equivalent — pinning is a client-side concept. CSA only. |
| `teams_save/unsave_message` | No Graph equivalent — bookmarks are a Teams client feature. chatsvc only. |
| `teams_get_saved_messages` | Same — no Graph API for saved/bookmarked messages. |
| `teams_get_followed_threads` | No Graph equivalent. |
| `teams_get_unread` | Partial — `GET /me/chats` returns `chatViewpoint.lastMessageReadDateTime` but no per-conversation unread count. |
| `teams_mark_read` | No direct Graph equivalent for consumption horizons. |
| `teams_get_activity` | No Graph equivalent for the Teams activity feed (mentions, reactions, replies). |
| `teams_search_emoji` | Client-side only — no API for this anywhere. |

### ❌ Graph API Cannot Replace

| Feature | Why |
|---|---|
| Favourites/pinned conversations | Teams client concept, not in Graph |
| Saved/bookmarked messages | Teams client concept, not in Graph |
| Activity feed | Teams-specific notification system |
| Followed threads | Teams client concept |
| Custom emoji metadata | CSA-specific |

## Key Findings

1. **No app registration needed.** The Teams SPA client ID already has delegated Graph permissions. The `.default` scope returns whatever permissions the app has consented to.

2. **Single token for ~70% of tools.** Graph can replace all core messaging, search, calendar, people, and transcript operations. Only Teams-specific client features (favourites, saved messages, activity feed, unread tracking) still need chatsvc/CSA.

3. **Auth simplification potential.** Could reduce from 4 tokens (Substrate, Skype Spaces, chatsvcagg, skypetoken_asm) to 2 tokens (Graph + chatsvc/CSA) for the features Graph can't cover.

4. **Better API surface.** Graph API is well-documented, stable, versioned, and returns richer response data (e.g., `from.user` with display name and tenant ID on messages).

5. **Graph Search supports `chatMessage` entity type.** This means we could potentially replace Substrate search for Teams message search, though Substrate may still be needed for file search and other entity types.

6. **Transcript API exists in Graph** (`GET /me/onlineMeetings/{id}/transcripts`) but requires `OnlineMeetingTranscript.Read.All` permission — need to verify if the Teams SPA client has this.

## Potential Migration Strategy

### Phase 1: Dual-mode (current spike)
- Keep existing chatsvc/Substrate implementations
- Add Graph API alternatives for A/B testing
- Validate Graph API works for all replaceable tools

### Phase 2: Graph-first
- Migrate core messaging tools to Graph API
- Migrate search to Graph Search API
- Migrate calendar to Graph Calendar API
- Keep chatsvc/CSA only for Teams-specific features

### Phase 3: Token simplification
- Remove Substrate and Skype Spaces token refresh if no longer needed
- Reduce to Graph + chatsvc tokens only
- Simplify auth-guards and token-extractor

## All Tokens in the MSAL Cache

The Teams Web Client (`Microsoft Teams Web Client`) has access to **10 distinct token audiences**. These were extracted by decoding the JWT `scp` claims from all access tokens in the session state.

### 1. Microsoft Graph API (`https://graph.microsoft.com`)

```
AppCatalog.Read.All        Calendars.Read              Calendars.Read.Shared
Calendars.ReadWrite        Calendars.ReadWrite.Shared   Channel.ReadBasic.All
ChatMember.Read            ChatMessage.Send             Files.ReadWrite.All
FileStorageContainer.Selected  Group.Read.All           InformationProtectionPolicy.Read
Mail.Read                  Mail.ReadWrite               MailboxSettings.ReadWrite
Notes.ReadWrite.All        Organization.Read.All        People.Read
Place.Read                 Place.Read.All               Place.Read.Shared
Sites.ReadWrite.All        Tasks.ReadWrite              Team.ReadBasic.All
TeamsAppInstallation.ReadForTeam  TeamsTab.Create        User.ReadBasic.All
```

### 2. Substrate (`https://substrate.office.com`)

```
ActivityFeed-Internal.Post          ActivityFeed-Internal.Read
ActivityFeed-Internal.ReadWrite     Files.Read.All
Files.ReadWrite                     Files.ReadWrite.Shared
Grammars-Internal.ReadWrite         LicenseAssignment.Read.All.Sdp
Notes-Internal.ReadWrite            OfficeFeed-Internal.ReadWrite
PeoplePredictions-Internal.Read     PlaceDevice.Read.All
Signals.Read                        Signals.ReadWrite
Signals-Internal.Read.Shared        Signals-Internal.ReadWrite
Sites.Read.All.Sdp                  SubstrateSearch-Internal.ReadWrite
Tasks.ReadWrite                     User.ReadWrite
```

### 3. Outlook/Exchange (`https://outlook.office.com/`)

```
Calendars.ReadWrite                 Collab-Internal.ReadWrite
Contacts.ReadWrite                  EWS.AccessAsUser.All
Files.Read.All                      Files.ReadWrite.Shared
Group.Read.All.Sdp                  Group.ReadWrite.All
Group.ReadWrite.All.Sdp             LicenseAssignment.Read.All.Sdp
Mail.ReadWrite                      Mail.Send
OWA.AccessAsUser.All                Place.Read.All
Place.ReadWrite.All                 Policy.Read.All.Sdp
Signals.ReadWrite                   TailoredExperiences-Internal.ReadWrite
User.Invite.All.Sdp                 User.Read
User.Read.Sdp                       User.ReadBasic.All
User.ReadWrite
```

### 4. SharePoint (`https://{tenant}.sharepoint.com`)

```
Container.Selected    MyFiles.Write    Sites.FullControl.All
Sites.Manage.All      User.ReadWrite.All
```

### 5. Substrate Search (`https://outlook.office.com/search`)

```
SubstrateSearch-Internal.ReadWrite
```

### 6. Teams Presence (`https://presence.teams.microsoft.com/`)

```
user_impersonation
```

### 7. IC3 Teams (`https://ic3.teams.office.com`)

```
Teams.AccessAsUser.All
```

### 8. Loki/Delve (`394866fc-eedb-4f01-8536-3ff84b16be2a`)

```
LLM.Read    User.Read.All
```

### 9. Unknown Service (`6bc3b958-689b-49f5-9006-36d165f30e00`)

```
User.Read    user_impersonation
```

### 10. chatsvcagg (cookie-based)

The `skypetoken_asm` cookie — not a standard MSAL access token but used for messaging APIs.

## New Tool Opportunities

Beyond our current tool set, these tokens unlock entirely new capabilities:

| Capability | API | Permission | Notes |
|---|---|---|---|
| **Read/send email** | Graph or Outlook | `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` | Full Outlook email access |
| **Manage calendar (write)** | Graph | `Calendars.ReadWrite` | Create/update/delete events |
| **Read/write OneNote** | Graph | `Notes.ReadWrite.All` | Notebooks, sections, pages |
| **Manage Planner/To Do tasks** | Graph | `Tasks.ReadWrite` | Create/update/complete tasks |
| **Read/write SharePoint sites** | Graph or SharePoint | `Sites.ReadWrite.All` | Lists, documents, pages |
| **Read/write OneDrive files** | Graph | `Files.ReadWrite.All` | Upload, download, share |
| **Get user presence/availability** | Presence API | `user_impersonation` | Online/busy/away/offline status |
| **Read/write contacts** | Outlook | `Contacts.ReadWrite` | Outlook contacts |
| **Manage mailbox settings** | Graph | `MailboxSettings.ReadWrite` | Auto-replies, timezone, etc. |
| **Read room/building info** | Graph | `Place.Read.All` | Meeting room availability |
| **Read org structure** | Graph | `Organization.Read.All` | Org directory info |
| **Full Exchange Web Services** | Outlook | `EWS.AccessAsUser.All` | Legacy but powerful |
| **Full SharePoint control** | SharePoint | `Sites.FullControl.All` | Admin-level SharePoint access |

## Open Questions

- Does Graph Search for `chatMessage` perform as well as Substrate v2/query?
- Does the Graph Calendar API return the same meeting detail (threadId, joinUrl) as the current mt/part API?
- Are there rate limits on Graph API that differ from the internal Teams APIs?
- Does Graph API work for government clouds (GCC, GCC-High, DoD) with the same endpoint?
- Can we refresh tokens for all 10 audiences, or only the ones we explicitly request scopes for?
- The Outlook token has `Mail.Send` — could we build an email tool alongside Teams?
- The Presence API token — can we read other users' presence status for "is X available?" queries?
- Are these scopes consistent across tenants, or do they depend on admin consent policies?
