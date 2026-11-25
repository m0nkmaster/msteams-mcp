---
type: Decision
title: Chat attachment upload path guards
description: Attachments only from the upload root; dot paths, credential files, config and app-data folders, UNC, 8.3 and ADS refused
tags: [attachments, security, windows]
status: stable
stale_after: 2027-05-25
---

# Chat attachment upload path guards

**Decided 2026-11-25** after the AGENTS.md security audit (George: "Add it") and three adversarial review rounds.

## Why

Teams content is untrusted. A prompt-injected message can ask the agent to attach `~/.ssh/id_rsa` or `session-state.json`; uploading shares it with the chat. Guards limit what an injected path can reach.

## Rules (`resolveUploadFile` in `src/api/chat-attachments-api.ts`)

- Absolute path, inside the upload root: `TEAMS_MCP_UPLOAD_DIR` (must be absolute) else the home directory. Lexical check first, then `fs.realpathSync.native`, then containment again (symlinks).
- Windows, before any fs call: UNC/device paths refused (an SMB probe leaks NTLM), 8.3-shaped segments refused, `:` alternate data streams refused.
- Refused after realpath: any dot segment relative to the root, `_netrc` / `_git-credentials`, the server config dir, `home/AppData` and `%APPDATA%`/`%LOCALAPPDATA%` (Windows), `home/Library` case-folded (macOS).
- Regular file, at most 4 MB.

## Lessons

- Plain `fs.realpathSync` keeps Windows 8.3 short names (`ADO-MC~1`); `.native` returns `.ado-mcp`. Confirmed live 2026-11-25. Always use `.native` for security checks.
- A worktree under `.worktrees/` is itself a dot path: set `TEAMS_MCP_UPLOAD_DIR` when testing from one.
