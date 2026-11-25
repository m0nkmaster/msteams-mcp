---
type: Decision
title: Chat attachments share like the Teams client
description: Upload to OneDrive "Microsoft Teams Chat Files", createLink view/users, grant read to the chat's other members
tags: [attachments, graph, sharing, chatsvc]
status: draft
stale_after: 2027-05-25
---

# Chat attachments share like the Teams client

**Decided 2026-11-25** (branch `feat/send-message-attachments`). Status `draft` until the manual recipient test passes.

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

- UNCONFIRMED: `grant` accepts `objectId` recipients (needs a second person). Fallback: grant by email via `profile-api`.
- Guests and very large meeting chats untested. See BACKLOG.md pre-mortem.
