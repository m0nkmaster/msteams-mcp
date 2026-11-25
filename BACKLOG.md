# Backlog

Work items not yet scheduled. Strategic direction lives in [ROADMAP.md](ROADMAP.md).

## Send a chat message with file attachments

**Added:** 2026-11-24 (George Murga).

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

**Sharing decision (2026-11-25, evidence from a real attachment the Teams web client sent):** the client grants access with a sharing link, not a direct invite. Graph `GET /me/drive/items/{id}/permissions` on that file shows one `view` link of scope `users`, role `read`, granted to exactly the chat's other members; its permission `id` is the message's `fileInfo.shareId` and its `link.webUrl` is `fileInfo.shareUrl` (`permissionScope: SpecificPeople`). Implementation mirrors it: `POST /me/drive/items/{id}/createLink` `{type: view, scope: users}`, then `POST /shares/u!{base64url(link.webUrl)}/permission/grant` `{recipients: [{objectId}], roles: [read]}`. Member object IDs come from chatsvc `GET /v1/threads/{id}?view=msnp24Equivalent` (`members[].id` = `8:orgid:{oid}`); the conversations endpoint returns no members. Self-notes (`48:notes`) get the link but no grant. Channels are out of scope (the client stores channel files in the team's SharePoint site, not the sender's OneDrive). UNCONFIRMED until the manual recipient test: that `grant` accepts `objectId` recipients.

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
