# Backlog

Work items not yet scheduled. Strategic direction lives in [ROADMAP.md](ROADMAP.md).

## Send a chat message with file attachments

**Added:** 2026-11-24 (George Murga).

**Status (2026-11-25):** implemented on branch `feat/send-message-attachments` (`teams_send_message` `files`). Verified live in self-notes only: two files uploaded, the read-back message carries both `properties.files` entries, and both download byte-identical. Open: the manual recipient-access test (1:1, group, meeting chat) and checking the file chiclets render in the Teams client.

**Status (2026-11-26):** 1:1 chats verified end to end. Verified live 2026-11-26 in two real 1:1 chats (server messages `1795690515745`, `1795691066392`): Graph shows each file's `view`/`users` link with role `read` granted to exactly the other member's object ID, and a colleague opened both attachments in Teams. File cards render and open in the sender's Teams client. Still untested live: group chats and meeting chats (same code path except members come from the chatsvc threads endpoint); test them on the next real case.

**Problem:** `teams_send_message` sends text only. Posting a message with files (for example an ARD as `.md` and `.html`) today means driving Teams web in Edge over CDP: the attach button differs by chat type and compose mode (compact `sendMessageCommands-FilePicker` opens a menu with "Upload from this device"; expanded `newMessageCommands-FilePicker` opens a native file dialog that CDP file-chooser interception does not catch and that blocks the page). On 2026-11-24 that route failed in a meeting chat and left a native "Open" dialog over Edge.

**Browser fallback that worked (2026-11-24, meeting chat):** click the expanded-compose paperclip (`newMessageCommands-FilePicker`) once and wait: its panel loads slowly (spinner), then shows "Upload from this device" and an `input[type=file]` appears in the page; `DOM.setFileInputFiles` on that input attaches the files. Files with names already shared in the chat are uploaded as copies with " 2" appended. Clearing the compose text with synthetic Ctrl+A/Delete did not work there.

**Proposal:** a tool (`teams_send_message` with a `files` parameter, or a separate `teams_send_message_with_files`) that:

1. uploads each file with the Teams client's Graph token (`Files.ReadWrite.All`, already reused for downloads) to the sender's OneDrive "Microsoft Teams Chat Files" folder, the folder the Teams client itself uses;
2. grants the chat's members access to each file (Graph sharing / invite), so recipients can open them;
3. sends the message through chatsvc with the files in the message's `files` property, the way the Teams client attaches files.

**Unknowns to settle first:**

- The exact chatsvc `files` property shape: capture it from one real message that has an attachment (read-only) and copy it.
- Graph token handling: use the existing optional, on-demand Graph token (`requireGraphTokenAsync()`, as downloads do), not a core refresh scope, so a tenant that refuses Graph cannot break the rest of Teams. Upstream PR #26 added Graph to `REFRESH_SCOPES`; do not copy that.
- Recipient access: confirm a chat member, not the sender, can open and download the file.

**Sharing decision (2026-11-25, evidence from a real attachment the Teams web client sent):** the client grants access with a sharing link, not a direct invite. Graph `GET /me/drive/items/{id}/permissions` on that file shows one `view` link of scope `users`, role `read`, granted to exactly the chat's other members; its permission `id` is the message's `fileInfo.shareId` and its `link.webUrl` is `fileInfo.shareUrl` (`permissionScope: SpecificPeople`). Implementation mirrors it: `POST /me/drive/items/{id}/createLink` `{type: view, scope: users}`, then `POST /shares/u!{base64url(link.webUrl)}/permission/grant` `{recipients: [{objectId}], roles: [read]}`. Member object IDs come from chatsvc `GET /v1/threads/{id}?view=msnp24Equivalent` (`members[].id` = `8:orgid:{oid}`); the conversations endpoint returns no members. Self-notes (`48:notes`) get the link but no grant. Channels are out of scope (the client stores channel files in the team's SharePoint site, not the sender's OneDrive). CONFIRMED 2026-11-26: `grant` accepts `objectId` recipients (1:1 chats).

**Follow-ups (out of scope):** upload sessions for files over 4 MB (simple upload limit); channel attachments; attachments on scheduled messages.

**Pre-mortem (2026-11-25, after three review rounds). Assume the feature shipped and failed. Ranked causes and decisions:**

- P0 `grant` refuses `objectId` recipients, so members cannot open the files. Checkpoint: the manual recipient test. Fallback if it fails: resolve member emails through `profile-api` and grant by `email`.
- P0 chatsvc accepts `properties.files` from this client but Teams does not render the file chiclet (for example because it expects extra properties). Checkpoint: the live self-notes send, read back and opened in Teams.
- P1 Tenant sharing policy blocks "specific people" links, or members are guests from another tenant, so `createLink`/`grant` fail. Decision: fail closed (nothing is sent, the error names the file). Revisit if users report it.
- P1 Large meeting chats exceed an undocumented per-call recipient limit on `grant`. Checkpoint: the manual test in a meeting chat. Fallback: batch the recipients.
- P1 Prompt injection asks to attach an ordinary-looking file inside the home directory (for example `Documents\passwords.xlsx`). Path guards cannot tell; the AGENTS.md rule (attach only files the user named) and a narrow `TEAMS_MCP_UPLOAD_DIR` are the controls. Decision: accepted; recommend setting `TEAMS_MCP_UPLOAD_DIR` in shared setups.
- P2 An ambiguous send failure (timeout, 5xx) followed by a retry posts the message twice. The error tells the caller to check the chat first. Decision: accepted.
- P2 Failed sends leave orphan copies in "Microsoft Teams Chat Files" (reused for 10 minutes, then left). Decision: accepted; the error says where they are.
- P2 OneDrive localises "Microsoft Teams Chat Files" for non-English users, so files land in a second, English-named folder. Unverified. Decision: accepted until reported.
- P2 Editing a file without changing its size or mtime within 10 minutes reuses the earlier upload. Decision: accepted.
- P2 The 8.3-shape rule refuses a genuine file named like `ABCDEF~1.TXT`. Decision: accepted; the error says to pass the long path.

**Done when:** one call posts text plus N files to a 1:1, a group and a meeting chat; recipients can open every file; tests cover the upload, sharing and message body; README and AGENTS.md list the tool.

**Related:** ROADMAP.md, "Future Consideration: Microsoft Graph API (hybrid)", file upload bullet (its "being trialled in PR #26" wording is stale). [Upstream PR #26](https://github.com/m0nkmaster/msteams-mcp/pull/26) (`teams_upload_file`, by dkattan) was closed unmerged and unreviewed on 2026-09-25: reference code only. Worth reusing: `buildFilesProperty()`, which maps the Graph DriveItem to the chatsvc `properties.files` shape (verify against a real message). Gaps: Graph added to core refresh, no sharing/permission call for recipients, 4 MB simple-upload limit, no new tests.

## Stop default auth-error advice from recommending `forceNew`

**Added:** 2026-11-25 (George Murga).

**Problem:** `getDefaultSuggestions()` in `src/types/errors.ts` (lines 78-89) tells the AI, for `AUTH_REQUIRED` and `AUTH_EXPIRED`, to "IMMEDIATELY call teams_login with { "forceNew": true }". `forceNew` deletes the whole config directory (browser profile, cookies, every signed-in account, session state) and forces an interactive sign-in. That contradicts the AGENTS.md rule "`forceNew` is destructive ... use it only when the user asks or approves" and the Troubleshooting order (plain `teams_login` first, which tries headless SSO). Seen in practice on 2026-11-25: an expired Graph token was fixed by a plain `teams_login` with no prompt; `forceNew` would have wiped the profile for nothing.

**Proposal:** change both suggestion lists to: call `teams_login` (without `forceNew`) first; use `forceNew: true` only if that fails and the user approves (wrong account or corrupted session). Check other places that recommend `forceNew` (tool descriptions in `src/tools/auth-tools.ts`, README, `docs/`) and align them.

**Done when:** a test pins the new suggestions for `AUTH_REQUIRED` and `AUTH_EXPIRED` (written first, seen failing); no default suggestion recommends `forceNew` without user approval; README/docs wording matches AGENTS.md.

**Status (2026-11-26):** done on branch `fix/no-forcenew-default-advice`, PR: [george-cm/msteams-mcp#3](https://github.com/george-cm/msteams-mcp/pull/3). Default suggestions, the auto-login-failed error and the `teams_login` description point to a plain `teams_login` first; `forceNew` only after that fails or signs in the wrong account, with the user's approval. Tests: pinned suggestion lists, a sweep over every `ErrorCode`, a source scan of every string literal in `src/` (TypeScript AST), and a server test for the auto-login-failed branch.

**Finding (2026-11-26):** the old server error claimed a plain `teams_login` "will return early without fixing the problem". False with dead tokens: auto-login is HTTP refresh plus headless SSO only, while plain login falls back to a visible sign-in (`src/tools/auth-tools.ts`, `src/auth/token-refresh.ts`). True in one case: Substrate token still valid but messaging cookies missing or expired, where the fast path returned "Already authenticated". Fixed: the fast path runs one serialized HTTP refresh and returns early only if messaging is usable afterwards; otherwise it continues to headless SSO and then a visible browser.

**Decisions (second-order review, 2026-11-26), what an AI client does with the new wording when plain login keeps failing:**

- Login times out or cannot open a window (user away, no display, profile locked): `teams_login` returns an error, so the escalation clause fires and the AI asks the user to approve `forceNew`. The user may approve a wipe they did not need; the suggestion names what gets deleted, so the choice is informed. Decision: accepted.
- Login reports success but the request still fails (for example an optional Graph token refused while the Substrate token has 10+ minutes left): login says "Already authenticated", the escalation clause ("if that teams_login fails") never fires, and the AI can loop login, retry, login. Before this change the same case wiped the profile on every pass. Decision: not fixed here; see "Plain `teams_login` does not refresh optional tokens on the fast path" below, which also proposes a loop-breaking suggestion.
- Several tool calls fail at once and the AI calls `teams_login` in parallel: two browsers on one profile hit the Chromium lock, one login fails, and the AI asks to approve `forceNew`. `handleLogin` is not serialized (pre-existing). Decision: accepted for now; follow-up below.
- "IMMEDIATELY" was dropped, so the AI could be less eager to re-login. "Do NOT skip this step or tell the user Teams is unavailable" is kept in every list. Decision: accepted.
- A plain login can now open a visible window without asking first (as `forceNew` always did). The auto-login-failed suggestion tells the AI to warn the user that a sign-in window may open. Decision: accepted.

**Decisions (second-order review of the post-sign-in `API_ERROR` and the sign-in guards, 2026-11-26):**

- The `API_ERROR` after sign-in says "do not call teams_login again in a loop", but an AI client can ignore suggestions and retry anyway. Each retry costs one headless SSO (~8s, invisible) plus one serialized HTTP refresh, then fails again — a bounded, invisible loop, far cheaper than the old login->auth-fail->auto-login cycle. Decision: accepted; the wording is the mitigation.
- While any sign-in is in progress, `refreshTokensViaBrowser` and the on-demand token refreshes fail fast at entry — including their HTTP leg, which writes session state. Consequence: every Teams tool call fails auth for the whole sign-in window (up to ~5 min for a visible login). Trade-off accepted: a session-state write landing mid-`forceNew` wipe can resurrect the wiped account, which is worse than a temporary outage with a clear "wait" error.
- A second `teams_login` is refused (`API_ERROR`, "already in progress") rather than joined. Joining would return the first call's result — wrong for `forceNew`, which must wipe — and refusal gives the client an actionable signal. Consequence: the user must wait for the open window or close it; the visible login's own timeout bounds the wait. Decision: accepted.
- `forceNew` is refused while any sign-in is in progress: wiping the profile under an open sign-in is the exact leak `forceNew` exists to prevent. Decision: accepted.
- The "forceNew may help" hint is only offered after headless SSO. After any interactive (visible) sign-in — including `forceNew` itself — the wording is always "will not fix", because the fresh sign-in already rebuilt the auth state a wipe would recreate. Decision: accepted.
- `teams_login` can block ~105s behind a wedged refresh (90s deadline + wedged-browser cleanup) before headless SSO starts, then up to 5 min in the visible browser. An MCP client may time out the call while the server keeps running; a retry then gets "already in progress" until the first attempt's `finally` clears the guard — self-correcting. Decision: accepted.

## Plain `teams_login` does not refresh optional tokens on the fast path

**Added:** 2026-11-26 (George Murga), from the `forceNew` advice review.

**Problem:** when the Substrate token has 10+ minutes left and messaging is usable, `teams_login` returns "Already authenticated" without touching the optional Graph or Assignments tokens. An expired SPA refresh token for Graph (AADSTS700084) therefore survives a plain login until the Substrate token is nearly expired, and file tools tell the AI to call `teams_login` and retry: a loop. `teams_login` calls are also not serialized, so parallel calls can collide on the browser profile lock.

**Proposal:** let the fast path continue to headless SSO when an optional-resource refusal is remembered (`token-refresh-http.ts`, 30-minute refusal memory); serialize `handleLogin`; add an `AUTH_*` suggestion "If teams_login reports success but the same request fails again, stop and tell the user instead of retrying".

**Done when:** a test shows a remembered Graph refusal sends plain login to headless SSO; concurrent `teams_login` calls share one attempt; the loop-breaking suggestion is pinned by a test.

## CLI drops everything after the first newline under `npm run` on Windows

**Added:** 2026-11-26 (George Murga).

**Problem:** `npm run cli -- send_message "$(cat msg.md)" --files ...` on Windows: npm launches the script through `cmd.exe`, which cuts the argument list at the first newline. The message was sent with only its first paragraph and without `--files`, and the CLI reported `ok` (self-notes message `1795690349453`, 2026-11-26). Calling `node node_modules/tsx/dist/cli.mjs src/cli.ts ...` (or the built `node dist/cli.js`) directly passes multi-line arguments intact.

**Proposal:** add a `--content-file <path>` (or `-` for stdin) option for `content`-style parameters in `src/cli/resolve.ts`, and document in AGENTS.md (CLI as test harness) that multi-line arguments must not go through `npm run` on Windows.

**Done when:** a test covers reading content from a file; AGENTS.md documents the trap; a multi-line message with flags after it sends intact on Windows.
