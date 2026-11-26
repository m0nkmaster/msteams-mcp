---
type: Decision
title: Chat attachments share like the Teams client
description: Upload to OneDrive "Microsoft Teams Chat Files", createLink view/users, grant read to the chat's other members
tags: [attachments, graph, sharing, chatsvc]
status: stable
stale_after: 2027-05-25
---

# Chat attachments share like the Teams client

**Decided 2026-11-25** (branch `feat/send-message-attachments`). Verified in 1:1 chats on 2026-11-26.

## Decision

`teams_send_message` `files` uploads each file with Graph to the sender's OneDrive folder "Microsoft Teams Chat Files", then:

1. `POST /me/drive/items/{id}/createLink` `{type: view, scope: users}`
2. `POST /shares/u!{base64url(link.webUrl)}/permission/grant` `{recipients: [{objectId}], roles: [read]}` for the chat's other members (skipped in self-notes)
3. sends the message with a chatsvc `properties.files` entry (JSON string) whose `fileInfo.shareId` = permission id and `fileInfo.shareUrl` = `link.webUrl`.

Members: 1:1 chats derive both ids from `19:<a>_<b>@unq.gbl.spaces` (the caller must be exactly one of them); other chats read chatsvc `GET /v1/threads/{id}?view=msnp24Equivalent` (`members[].id` = `8:orgid:{oid}`). Channels are refused.

## Evidence

- A real message sent by the Teams web client (meeting chat, 2026-11-24): Graph `GET /permissions` on its files showed one `view` link, scope `users`, role `read`, granted to exactly the 2 other members; its id and `webUrl` equal the message's `shareId` / `shareUrl`.
- Graph `invite` (the original hypothesis) was not chosen: its result shape is unverified, `createLink` + `grant` reproduces the client's permission exactly.
- Live self-notes send 2026-11-25: both entries read back, both files downloaded byte-identical.

## Open

- CONFIRMED 2026-11-26: `grant` accepts `objectId` recipients. Verified live 2026-11-26 in two real 1:1 chats (server messages `1795690515745`, `1795691066392`): Graph shows each file's `view`/`users` link with role `read` granted to exactly the other member's object ID, and a colleague opened both attachments in Teams.
- Untested live: group chats, meeting chats, guests, very large meeting chats. Fallback if a grant is refused: grant by email via `profile-api`.
- Windows CLI trap: multi-line content via `npm run cli` loses everything after the first newline (BACKLOG.md). See BACKLOG.md pre-mortem.
