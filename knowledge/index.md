---
okf_version: "0.2"
type: Index
title: msteams-mcp (fork) knowledge
---

# msteams-mcp knowledge

- [Chat attachments share like the Teams client](decisions/chat_attachments_sharing.md) - createLink view/users + grant read to chat members; evidence from a real message
- [Chat attachment upload path guards](decisions/chat_attachment_path_guards.md) - upload root, refused paths, and why (prompt-injection exfiltration, 8.3 short names)
- [Graph token expired: plain teams_login, never forceNew](playbooks/graph_token_expired_teams_login.md) - AADSTS700084 SPA refresh-token 24h limit; headless SSO fixes it
- [teams_login failure semantics and the forceNew policy](decisions/teams_login_failure_and_forcenew_policy.md) - plain login first, forceNew with approval only, API_ERROR instead of fake success, sign-in guards
