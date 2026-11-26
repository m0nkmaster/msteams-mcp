---
type: Decision
title: teams_login failure semantics and the forceNew policy
description: Plain teams_login first; forceNew only with user approval; login fails (API_ERROR) instead of faking success; sign-in guards
tags: [auth, login, forceNew, token-refresh]
status: stable
stale_after: 2027-05-26
---

# teams_login failure semantics and the forceNew policy

**Decided 2026-11-26**, merged in [PR #3](https://github.com/george-cm/msteams-mcp/pull/3) (`0c54057`). Covered by unit tests only; no live run of the new paths yet.

## Policy

- Default advice for `AUTH_REQUIRED` / `AUTH_EXPIRED` and the auto-login-failed error: call plain `teams_login` first.
- `forceNew: true` only if that fails or signs in the wrong account, and only with the user's approval. It deletes the whole config directory (browser profile, cookies, every account).
- A guard test (`src/force-new-advice.test.ts`) scans every string in `src/` with the TypeScript parser. It fails on any forceNew recommendation without an un-negated approval phrase in the same clause. Template literals, `+` chains and same-file constants are joined; imported constants and `OBJ.KEY` are not.

## Login behaviour

- **Fast path** (Substrate token >= 10 min): returns "Already authenticated" only if messaging (`skypetoken_asm`) is usable after one serialized HTTP refresh. Otherwise headless SSO, then a visible browser.
- **After any sign-in:** if the session is still unusable after one serialized HTTP refresh, `teams_login` returns `API_ERROR` (not retryable) instead of success. Reason: a fake success loops login -> retry forever.
  - After headless SSO with an auth-state refresh error: says forceNew *may* help, with approval.
  - After a visible sign-in: says forceNew will not help.
- **Sign-in guard:** while a `teams_login` browser is open, automatic and on-demand refreshes fail fast, and a second `teams_login` is refused ("already in progress"). Calls with valid tokens keep working.
- **forceNew** waits for any in-flight refresh before deleting.
- **Session epoch:** an HTTP refresh that read the session before a wipe skips its write, so a wiped account cannot come back.

## Why plain login, not forceNew

The old advice claimed plain login "returns early without fixing the problem". Code reading showed that was only true when the Substrate token was valid but messaging was broken; that case is now fixed. Plain login fixed expired tokens silently three times (2026-11-25/26).

## Open (BACKLOG.md)

- Plain login does not refresh optional tokens (Graph) on the fast path: possible login -> retry loop.
- The browser's own session save has no epoch check (narrow window, deferred).
