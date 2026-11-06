# msteams-mcp codebase audit

- **Date:** 2026-11-06
- **Commit:** `b87053a` (v0.30.3)
- **Scope:** all of `src/` (about 15,000 lines of source, 5,400 lines of tests), plus `docs/` and the repo root.
- **Checks run:** `vitest run` (26 files, 559 tests, all pass), `tsc --noEmit` (clean), `eslint src/` (clean). Nothing was run against live Teams.

This is a snapshot. Once the fixes land, delete it, or it will go stale like the two February reviews (see S7).

---

## 1. What this repo does

It's an MCP server and a CLI that let an AI assistant use Microsoft Teams. It signs in once through a real Chrome or Edge window and saves the session (cookies and Microsoft sign-in tokens) to an encrypted file. After that it calls Teams' own internal APIs directly: Substrate for search, chatsvc for messages, CSA for favourites, mt/part for calendar and profiles. Each tool's definition drives the MCP server and the CLI alike.

**Load assumed:** one person, but often several processes at once. Claude Desktop, Claude Code and the `msteams` CLI each start their own server process, and they all share one session file (`~/.teams-mcp-server/session-state.json`) and one browser profile. Several of the findings only show up under that kind of sharing.

## 2. The short answer

The codebase hasn't sprawled into a mess. The layers are sensible (tools → API modules → auth guards → session store). The parsers are split by job and well tested. The CLI is generated from the tool definitions, and there's very little dead code.

The fragility sits in one place: **the rules about tokens, errors and retries are written several times, in several files, and the copies disagree.** Git history confirms it: the most-changed and most-fixed files are `auth-guards.ts`, `browser/auth.ts`, `token-refresh*.ts` and `server.ts`.

### The four big steps

| Step | What it means | Findings |
|---|---|---|
| **A. Make every write safe** | One rule for what may be retried. Copy what `assignments-api.ts` already does: don't replay a write after a timeout, and never guess. | 2, 3, 4 |
| **B. One source of truth for auth** | One way to pick a token from the session, one status report, one cache (the session file itself), and every browser launch through one queue. | 5, 7, 8, 14, 15 |
| **C. Safe with several processes** | Write the session file atomically, never delete a live browser lock, and use a key that doesn't change when the hostname does. | 6, 9, 10 |
| **D. Consolidate duplicated rules** | One "is this a real message" filter, one conversation-type check, one auth-plus-region guard, one result type. Move the pure message-building code out of the HTTP module. | 16, 17, 18, 19 |

**1** was a security issue and is already fixed, in the same branch as this document.

### Suggested order

1. **Today (small, high value):** 2, 3, 4. Each is a few lines. (1 is already fixed.)
2. **Next:** 5, 6, 9, plus the test in 20, so the auto-login path can't regress.
3. **Then the auth consolidation:** 7, 8, 14, 15. Afterwards auth is one rule per question, and the next Microsoft host change is a one-line fix.
4. **Clean-up when convenient:** 11–13 (real bugs, but low impact), 16–19, and the smaller items.

---

## 3. Findings

Each finding is marked **Confirmed** (I traced it in the code) or **Plausible** (the code does this, but it depends on how Microsoft's service behaves, which I couldn't test).

### Must fix

#### 1. Transcript fallback sends your Teams token to any link in the meeting chat
`src/api/transcript-api.ts:203-217` · Security · Plausible · **Fixed:** `findAmsTranscriptLink` (`src/utils/parsers-virtual.ts`) only returns HTTPS links on the hosts the skype token is issued for (`*.asyncgw.teams.microsoft.com`, `*.asm.skype.com`), with tests for planted and look-alike hosts. The message-type check below was left out, because the exact type string couldn't be confirmed without a live recording.

- **What this is:** If the recording isn't in your own OneDrive index, `teams_get_transcript` reads the meeting chat, finds a message containing `<item type="amsTranscript" uri="...">`, and downloads that URI with `Authorization: skype_token <your token>`.
- **Problem:** It doesn't check who sent the message, what type the message is, or what host the URI points to. Anyone in the meeting chat who can post HTML with that `<item>` tag (for example through the chatsvc API, which this project itself uses) can point the URI at their own server. When you ask for the transcript, your skype token goes to them. That token reads and sends your chats for up to about 24 hours.
- **Fix:** Only accept the message when its type is the system recording type (`RichText/Media_CallRecording`). Only send the token when the URI's host is the Teams async media gateway; check a real recording message to confirm the exact host suffix. Two conditions on lines 211–212.
- **If we skip it:** Anyone in a shared meeting can take over your Teams account for a day. Whether chatsvc actually stores an arbitrary `<item>` tag hasn't been tested.

#### 2. A permission error is treated as "logged out"
`src/types/errors.ts:138` · Bug · Confirmed

- **What this is:** It maps each HTTP status to an error code.
- **Problem:** 403 ("forbidden") becomes `AUTH_REQUIRED`. The server treats that as an expired login: it refreshes tokens (possibly in a headless browser), re-runs the tool, gets 403 again, and tells the AI to call `teams_login`. Deleting someone else's message, or posting in a read-only channel, ends in login advice. `assignments-api.ts:48` and `graph-files-api.ts:102` already undo this by hand.
- **Fix:** Map 403 to `ACCESS_DENIED` (the code already exists), then delete the two workarounds.
- **If we skip it:** Permission problems look like login loops.

#### 3. Any failed auto-refresh tells the AI to wipe everything
`src/server.ts:256` · Bug · Confirmed

- **What this is:** When a tool returns an auth error, the server refreshes once and retries the tool.
- **Problem:** If the refresh fails for *any* reason (offline, a timeout, the browser profile locked by another process), the reply is "IMMEDIATELY call teams_login with forceNew: true". `forceNew` deletes the whole config folder: browser profile, cookies, session. A wifi drop then ends in a full sign-in with MFA.
- **Fix:** Only give the forceNew advice when the refresh error is itself `AUTH_EXPIRED` or `AUTH_REQUIRED`. Otherwise return the refresh error as it is; network errors are already marked retryable.
- **If we skip it:** Users get fully logged out after short network failures.

#### 4. Writes are retried, so a message can be sent twice
`src/utils/http.ts:45` · Bug · Plausible

- **What this is:** Every API call goes through `httpRequest`, which retries up to 3 times on timeouts (30s), network errors and 5xx errors.
- **Problem:** That includes the POST behind `send_message`, plus edits, reactions, favourites and creating group chats. If Teams accepts the message but the response is lost or slow, the same message is sent again. `assignments-api.ts` already avoids this by hand (`maxRetries: 1`, plus "the action may have completed" handling). The other writes don't.
- **Fix:** In `httpRequest`, default `maxRetries` to 1 when the method isn't `GET`. One line.
- **If we skip it:** Duplicate messages on slow networks. It's only "plausible" because chatsvc might drop repeats with the same `clientmessageid`. I didn't verify that, and it wouldn't cover edits or reactions.

#### 5. Login can open a second browser on the same profile
`src/tools/auth-tools.ts:133` vs `src/auth/token-refresh.ts:36` · Bug · Confirmed

- **What this is:** Token refreshes go through a queue (`serializeRefresh`) so only one of them uses the browser profile and session file at a time. AGENTS.md says the same.
- **Problem:** `teams_login` opens its browsers outside that queue. AI clients often call tools in parallel. If one tool hits a 401 and starts a headless refresh while `teams_login` runs, two Chromes try to open the same profile, and both write the session file.
- **Fix:** Export `serializeRefresh` and run the body of `handleLogin` through it.
- **If we skip it:** Random "could not launch browser" errors and lost sessions.

#### 6. A browser launch deletes a lock that is still in use
`src/browser/context.ts:183` · Bug · Confirmed

- **What this is:** Chrome puts a `SingletonLock` in the profile so two Chromes can't share it. `cleanupStaleSingletonLock()` already removes the lock safely when its owner process is dead.
- **Problem:** If the launch still fails with a lock error, the catch block deletes the lock *anyway* and launches again, even when the owner is alive. With two processes (say Claude Desktop and the CLI) refreshing at once, the second one takes over the first one's live profile. Two Chromes writing one cookie database can corrupt it.
- **Fix:** Delete the forced unlink at lines 180–185. Return a clear, retryable "browser profile is in use by another process" error instead.
- **If we skip it:** Rare profile corruption, which then needs `forceNew`.

#### 7. Messaging ignores the Teams host list
`src/auth/token-extractor.ts:348` (lines 353 and 392) · Bug · Plausible

- **What this is:** It reads the `skypetoken_asm` and `authtoken` cookies that every chat API needs.
- **Problem:** It only looks at cookies whose domain contains `teams.microsoft.com`. AGENTS.md says every host check reads `TEAMS_HOSTS`, but this one doesn't. On GCC-High (`teams.microsoft.us`), or if Microsoft moves these cookies to a newer host, messaging auth is never found, and every chat tool loops through auto-login. It also takes the *first* matching cookie rather than the freshest, so after a refresh an old cookie can still win.
- **Fix:** Pick the cookie by name across all domains, and take the one with the latest JWT expiry.
- **If we skip it:** Government-cloud messaging is broken, and stale cookies cause one-off failures after refresh. "Plausible" because I can't see which cookie domains GCC-High actually uses. The HTTP refresh also hardcodes `authsvc.teams.microsoft.com` and `Origin: https://teams.microsoft.com`, so GCC refresh is probably broken as well.

### Should fix

#### 8. The rule for picking a token is written five different ways
`token-extractor.ts` (`bestJwt`, `extractCsaToken:439`), `session-store.ts` (`hasSubstrateSearchToken`), `token-refresh-http.ts` (`findAccessToken`), `browser/auth.ts` (`checkBrowserTokenExpiry`) · Bug · Plausible

- **What this is:** Each one searches the saved Microsoft token cache for a token.
- **Problem:** They disagree. `bestJwt` picks the newest unexpired token from the main Teams origin. `extractCsaToken` picks the *first* match from *any* origin and never checks expiry. `findAccessToken` picks the first match. Since the move to `teams.cloud.microsoft`, sessions often contain both origins. If the expired CSA token comes first, the refresh writes a fresh one into the other origin, the expired one still wins, and favourites keep failing (each failure triggers another refresh).
- **Fix:** `extractCsaToken = bestJwt(state, e => e.target?.includes('chatsvcagg'))?.token ?? null`, and make `hasSubstrateSearchToken` and `findAccessToken` use the same helper. The in-page browser check has to stay a separate copy because it runs inside the page.
- **If we skip it:** Every Microsoft host change brings another bug like #47 and #51.

#### 9. The session file is written in place while other processes read it
`src/auth/session-store.ts:73` · Risk · Confirmed

- **What this is:** Every refresh rewrites `session-state.json`.
- **Problem:** `writeFileSync` empties the file, then writes it. Another server process reading at that moment gets broken JSON, `readSecure` returns null, it decides you're logged out, and it starts auto-login (headless browser, maybe then visible). A crash mid-write leaves a broken file, which also means "logged out".
- **Fix:** Write to `filePath + '.tmp'`, then `fs.renameSync` over the real file. Two lines.
- **If we skip it:** Rare, unexplained re-login prompts when several clients are running.

#### 10. The encryption key changes when the computer's name changes
`src/auth/crypto.ts:21` · Risk · Plausible

- **What this is:** The session file is encrypted with a key built from `os.hostname()` and your username.
- **Problem:** On macOS the hostname can change, for example when you rename the Mac, or on some networks and VPNs. Afterwards the file can't be decrypted, the server thinks there's no session, and you have to sign in again. The key also gives little protection: anyone who can read the file can work out your hostname and username.
- **Fix:** On first run, create a random 32-byte key in `CONFIG_DIR/key` (mode 0600) and use that. It's just as safe and doesn't change. Keep reading the old key once so existing sessions move over.
- **If we skip it:** Occasional "logged out for no reason" on Macs.

#### 11. Unread count for one chat can count everything as unread
`src/api/chatsvc-readstatus.ts:192-206` · Bug · Confirmed

- **What this is:** `teams_get_unread` with a `conversationId` counts the messages newer than your last-read position.
- **Problem:** The loop skips system messages and deleted messages *before* checking whether it has reached your last-read message. It also compares by exact ID. If your read position is a system message (someone added to the chat) or a message that was later deleted, the loop never stops, and every message on the page (up to 50) is counted as unread. The all-chats mode does this differently and correctly: it compares timestamps.
- **Fix:** At the top of the loop, before the skips: `if (lastReadId && Number(id) <= Number(lastReadId)) break;`.
- **If we skip it:** Wrong unread counts, and the AI reports old messages as new.

#### 12. Email search pages repeat results
`src/tools/search-tools.ts:352` with `src/api/substrate-api.ts:239` · Bug · Confirmed

- **What this is:** Email search drops calendar replies ("Accepted: …") from each page, then tells the AI where the next page starts (`nextFrom`).
- **Problem:** `nextFrom` is `from + returned`, and `returned` counts only what's left *after* filtering. If 5 of 25 emails were calendar replies, the next page starts at 20 instead of 25, so 5 emails come back twice. `hasMore` in the same flow uses `from + size`, so the two numbers disagree.
- **Fix:** Have the API return the server's own item count and use `from + serverCount` for `nextFrom`.
- **If we skip it:** Paging through mail shows duplicates, and the AI may summarise the same email twice.

#### 13. Adding or removing one favourite fetches every favourite's details
`src/api/csa-api.ts:149` (and `:90`) · Speed · Confirmed

- **What this is:** `modifyFavorite` needs the favourites folder ID and version, so it calls `getFavorites()`.
- **Problem:** `getFavorites()` also looks up the display name of *every* favourite: one request each, plus a second for any without a name, all at once. With 40 favourites, adding one costs 41–81 requests. A burst like that can easily trigger a 429, and one 429 currently blocks *every* API (see S1).
- **Fix:** Split out a small `getFavoritesFolder()` that only reads the folder. `getFavorites()` adds the names on top, and `modifyFavorite` uses only the folder.
- **If we skip it:** Slow favourite changes, plus rate limits that spill over to other tools.

#### 14. Delete the second token cache
`token-cache.json`: `session-store.ts`, `token-extractor.ts:133`, `auth-guards.ts` (`handleSubstrateError`), 7 call sites · Lean · Confirmed

- **What this is:** A separate encrypted file holding a copy of the Substrate token.
- **Problem:** Reading it costs the same as reading the session file (one file read, one decrypt), so it saves nothing. But it's a second source of truth that every refresh path has to remember to clear: four places do, and the next new path probably won't.
- **Fix:** `getValidSubstrateToken = () => extractSubstrateToken()?.token ?? null`. Delete `TokenCache`, `readTokenCache`, `writeTokenCache`, `clearTokenCache`, `handleSubstrateError` and their test mocks.
- **If we skip it:** About 80 lines kept around for nothing, and one more thing that can disagree after a refresh.

#### 15. Status is implemented twice, and the two disagree
`src/server.ts:179` (`teams://status`) vs `src/tools/auth-tools.ts:179` (`handleStatus`) · Bug · Confirmed

- **Problem:** The resource says messaging is available whenever the cookies exist. The tool checks whether they've expired. With an expired skype token, one says yes and the other says no, exactly when someone is debugging auth.
- **Fix:** Export the body of `handleStatus` and have the resource return it.

#### 16. "Is this a real message?" and "what kind of chat is this?" are each written several times
Message filter: `chatsvc-messaging.ts:476`, `chatsvc-readstatus.ts:192`, `chatsvc-activity.ts:162`, `parsers-virtual.ts:47`. Chat type: `parsers-identifiers.ts:18` (`getConversationType`), `chatsvc-conversations.ts:38` (`classifyChatType`), `chatsvc-readstatus.ts:288`, `chatsvc-messaging.ts:793`. · Duplication · Confirmed

- **Problem:** The four message filters skip different things: some skip `DeleteMember` and some don't, and some check `deletetime` and some don't. So one message can appear in `get_thread`, be missing from the activity feed, and still count as unread. The chat-type checks also differ: one knows about 1:1 chats, another doesn't.
- **Fix:** One `isUserMessage(msg)` and one `getConversationType(id)` in `parsers-identifiers.ts`. Every caller uses them.
- **If we skip it:** Each new Teams message type gets fixed in one place and missed in the others.

#### 17. The auth-plus-region guard is copied by hand
`src/api/calendar-api.ts:193`, `src/api/tags-api.ts:59` · Lean · Confirmed

- **Problem:** Both call `requireSkypeSpacesAuth()` and then `getRegionConfig()`, each with its own copy of the error. `requireSkypeSpacesAuthWithConfig()` already does exactly this, and `profile-api.ts` uses it.
- **Fix:** Use the helper in both. About 20 fewer lines.

#### 18. Tool handlers have their own copy of the result type
`src/tools/index.ts:13` (`ToolResult`), 21 hand-written unwraps in `message-tools.ts`, 3 in `search-tools.ts` · Lean · Confirmed

- **Problem:** `ToolResult` (`{success, data} | {success, error}`) is a second copy of `Result<T, McpError>`. There are also two styles: `file-`, `people-`, `meeting-`, `tag-` and `assignment-tools` use `handleApiResult`, and `message-tools.ts` mostly writes `if (!result.ok) return { success: false, error: result.error }` by hand.
- **Fix:** Let handlers return `Result<Record<string, unknown>>` and have `invokeTool` turn it into the MCP shape. Delete `ToolResult` and `handleApiResult`.
- **If we skip it:** About 150 lines of boilerplate, and two patterns for new contributors to choose between.

#### 19. `chatsvc-messaging.ts` does two unrelated jobs
`src/api/chatsvc-messaging.ts:1071-1470` (of 1,567 lines) · Split · Confirmed

- **Problem:** About 400 lines of the biggest file are pure functions with no HTTP: mention HTML, link parsing, scheduled-draft bodies, 1:1 thread bodies, and the wait-for-reply helpers. They live in an HTTP module, while their nearest relative, `markdownToTeamsHtml`, is in `utils/parsers-markdown.ts`. This file is most of why the code *feels* sprawling.
- **Fix:** Move them into a `utils/parsers-compose.ts` (or into `parsers-markdown.ts`), and export them from the `parsers.ts` barrel. They already have tests, so this is a move with no logic changes.

#### 20. The riskiest code has the fewest tests
· Missing test · Confirmed

| Code | Why it's risky | Test today |
|---|---|---|
| `server.ts` auto-login retry | Decides refresh, re-running a tool, and the forceNew advice (findings 2, 3) | none |
| `auth/token-extractor.ts` | Every API's token choice (findings 7, 8) | none of its own |
| `api/chatsvc-readstatus.ts` | Unread counting (finding 11) | none |
| `api/transcript-api.ts` fallback | Sends a token to a URL taken from a message (finding 1) | none |
| `api/csa-api.ts`, `substrate-api.ts` | Favourites, search paging (findings 12, 13) | none |

- **Fix:** One small test each. The first should be a `server.test.ts` with mocked `invokeTool` and `refreshTokensViaBrowser`: check that a 403 doesn't refresh and that a network failure doesn't advise forceNew. The rest are pure-function tests once 11, 12 and 1 are fixed.

---

## 4. Smaller items (nice to have)

- **S1. One 429 blocks every API** (`src/utils/http.ts:34`). `rateLimitedUntil` is a single global, so a rate limit from Substrate search also blocks chat calls. Key it by URL host.
- **S2. Dead code.** `sendNoteToSelf` (`chatsvc-messaging.ts:317`) and `clearSessionState` (`session-store.ts:118`) are unused everywhere. `calculateTokenStatus` (`parsers-people.ts:159`) is only used by its tests, while the real status functions do the same maths inline; either use it or delete it. The fixtures `threadMessage` and `favoritesFolderResponse` are unused.
- **S3. Near-copies.** `searchPeople` and `getFrequentContacts` (`substrate-api.ts`) differ only by URL and query string. `SavedMessage` and `FollowedThread` are the same shape, and `VirtualConversationItem` is declared twice (`parsers-virtual.ts:14`, `chatsvc-virtual.ts:84`).
- **S4. Two paging rules for search.** Message search's `hasMore` uses `from + results.length`, email's uses `from + size` (`substrate-api.ts:133` vs `:239`). Pick one when fixing 12.
- **S5. Transcript filter input.** `threadId` is put inside an OData string unescaped (`transcript-api.ts:95`); escape `'` as `''`. An invalid `meetingDate` throws "Invalid time value"; validate it with `z.string().datetime()` in `meeting-tools.ts:28`.
- **S6. Invented timestamps.** `timestampFromIdOrNow` (`chatsvc-messaging.ts:1071`) gives a message *the current time* when it can't read one, so an old message can look new. Return no timestamp instead.
- **S7. Stale docs.** `docs/2026-02-12-full-code-review.md` and `docs/2026-02-15-clean-code-review.md` (600 lines) describe code that has since changed, and agents reading `docs/` will trust them. Delete them, and delete this file once it's done its job.
- **S8. Hardcoded hosts in refresh.** `token-refresh-http.ts` hardcodes `authsvc.teams.microsoft.com`, the `Origin` header and the cookie domains. That's the same root cause as 7; fix it in the same change if government clouds matter.

---

## 5. What's already good (keep it)

- **`assignments-api.ts` write safety** is the model for the rest: no replay of a write after a timeout, a "may have completed" message instead of a retry, and a read-back to confirm the new state.
- **`graph-files-api.ts` downloads:** a strict SharePoint host allowlist, refusal to overwrite existing files (`open(..., 'wx')`), absolute paths only, and a SHA-256 of the result.
- **The CLI is generated from the tool definitions,** so it can't drift from the MCP server.
- **The parsers** are split by job, pure, and cover most of the 559 tests.
- **Optional auth (Assignments, Graph) can never block core Teams,** and repeat refusals are remembered for 30 minutes.
- **Typecheck and lint are clean,** and every tool name mentioned in a description matches a registered tool.

## 6. Size of the clean-up

**Lean:** about −370 lines of source (token cache 80, result type 150, guard copies 20, dead code 60, near-copies 60) and −600 lines of stale docs. No dependencies to remove. Finding 19 moves about 400 lines without deleting any.

## 7. Not checked

- Nothing was run against live Teams. Every **Plausible** finding (1, 4, 7, 8, 10) depends on how Microsoft's service behaves, which I couldn't test: whether chatsvc stores arbitrary `<item>` tags, whether it ignores repeat `clientmessageid`s, which cookie domains GCC-High uses, and how often macOS hostnames change.
- Read in full: auth, server, HTTP client, errors, every API module, browser login, the tool handlers, the CLI entry point.
- Skimmed only: `cli/resolve.ts`, `cli/render.ts`, `cli/help.ts`, `parsers-search.ts`, `parsers-people.ts`, `parsers-channels.ts`, `parsers-markdown.ts`, `constants.ts`, `api-config.ts`, `logger.ts`, and the wording of the ~40 tool descriptions (checked for broken tool references, not for accuracy).
- Not read: `docs/API-REFERENCE.md`, `SESSION-DATA-REFERENCE.md`, `USER-STORIES.md`, `MANUAL-TEST-SCRIPT.md`, `ROADMAP.md`, the CI workflow.
