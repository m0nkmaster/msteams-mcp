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
- Whether the Graph refresh must stay non-fatal for restricted tenants (see ROADMAP "Hybrid opportunity" and PR #26, which trials `teams_upload_file`).
- Recipient access: confirm a chat member, not the sender, can open and download the file.

**Done when:** one call posts text plus N files to a 1:1, a group and a meeting chat; recipients can open every file; tests cover the upload, sharing and message body; README and AGENTS.md list the tool.

**Related:** ROADMAP.md, "Future Consideration: Microsoft Graph API (hybrid)", file upload bullet; PR #26.
