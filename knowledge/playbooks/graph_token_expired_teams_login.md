---
type: Playbook
title: "Graph token expired: plain teams_login, never forceNew"
description: AADSTS700084 (SPA refresh token 24h limit) on Graph features; fix with headless teams_login
tags: [auth, graph, troubleshooting]
status: stable
stale_after: 2027-05-25
---

# Graph token expired: plain teams_login, never forceNew

**Seen 2026-11-25.** Attachment send failed with `AUTH_INTERACTION_REQUIRED`: "AADSTS700084: The refresh token was issued to a single page app (SPA) ... fixed, limited lifetime of 1.00:00:00".

## Fix

1. `npm run cli -- login` (or `teams_login` without `forceNew`). It tried headless SSO and printed "Login completed silently via SSO" in about 17 s, no window.
2. Retry the call.

Plain `teams_login` reaches headless SSO whenever an optional resource (Graph, Assignments) was flagged by a failed on-demand refresh, re-mints it after sign-in, and returns a warning that forbids another login loop if it still fails. The flag is per process: the long-lived MCP server sees it, but a fresh `msteams login` CLI process does not — in the CLI, if login says "Already authenticated" the Graph token can still be dead until the Substrate token is near expiry.

## Do not

- Do not use `forceNew: true`: it deletes the whole config directory and forces an interactive sign-in. Only with George's approval.
