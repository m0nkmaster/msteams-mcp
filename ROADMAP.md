# Roadmap

| Priority | Feature | Description | Difficulty | Notes |
|----------|---------|-------------|------------|-------|
| **P1** | **Browserless token refresh** | **Refresh MSAL tokens via direct HTTP instead of opening a headless browser** | **Medium** | **See details below** |
| P3 | Find team | Search/discover teams by name | Easy | Teams List API |
| P3 | Get person details | Detailed profile info (working hours, OOO status) | Easy | Delve API |
| P3 | Get message by URL | Fetch a specific message from a Teams link with surrounding context | Medium | Parse URL to extract conversation/message IDs, return target message plus N messages before/after |
| P4 | Meeting attendees | Filter meetings by attendee (not just organiser) | Medium | Need to research attendee list in calendar API response |
| ~~P3~~ | ~~Reduce chatsvc-api boilerplate~~ | ~~Extract `requireMessageAuthWithConfig()` helper returning `{ auth, region, baseUrl }`~~ | ~~Easy~~ | ✅ Done — `requireMessageAuthWithConfig()` in `auth-guards.ts`, applied to all 13 chatsvc functions |
| ~~P3~~ | ~~Substrate error response wrapper~~ | ~~Extract helper that clears token cache on `AUTH_EXPIRED` after Substrate/files API calls~~ | ~~Easy~~ | ✅ Done — `handleSubstrateError()` in `auth-guards.ts`, applied to all 5 call sites |
| ~~P4~~ | ~~Split chatsvc-api.ts~~ | ~~Break 1,600+ line file into sub-modules (messaging, activity, reactions, virtual-conversations)~~ | ~~Medium~~ | ✅ Done — Split into 5 sub-modules: `chatsvc-messaging.ts`, `chatsvc-activity.ts`, `chatsvc-reactions.ts`, `chatsvc-virtual.ts`, `chatsvc-readstatus.ts` + `chatsvc-common.ts`. Barrel file (`chatsvc-api.ts`) re-exports everything, so existing `import { ... } from './chatsvc-api.js'` statements continue to work unchanged |
| ~~P4~~ | ~~Typed API response interfaces~~ | ~~Define lightweight interfaces for common raw API response shapes~~ | ~~Easy~~ | ✅ Done — `types/api-responses.ts` with `RawChatsvcMessage`, `RawConversationResponse`, `RawAllFilesResponse`, `RawFileItem`, etc. Applied to chatsvc, files, and readstatus modules |
| ~~P4~~ | ~~Additional test coverage~~ | ~~Add tests for new helpers and Result utilities~~ | ~~Medium~~ | ✅ Done — Added tests for `handleSubstrateError`, `formatHumanReadableDate`, `Result` utilities (ok/err/unwrap/map/andThen). 200 tests total (was 117) |

## Browserless Token Refresh

### Problem

Token refresh currently opens a headless Chrome/Edge browser with the persistent profile (~8 seconds). This is slow, resource-heavy (spawns a full browser process), and requires the Chromium profile lock — meaning it can't run concurrently with a visible login.

### Goal

Replace the browser-based refresh with a direct HTTP call to Azure AD's OAuth2 token endpoint. This would make token refresh instant (~100ms), eliminate the Playwright/browser dependency for refresh, and remove the profile lock contention issue.

### How It Works

MSAL (Microsoft Authentication Library) stores tokens in the browser's localStorage using a well-documented cache schema. The key entries are:

1. **Refresh tokens**: Stored under keys like `{homeAccountId}-login.microsoftonline.com-refreshtoken-{clientId}--{scopes}`. These are long-lived (days/weeks) and can be exchanged for fresh access tokens.
2. **Access tokens**: Stored under keys matching `{homeAccountId}-login.microsoftonline.com-accesstoken-{clientId}-{tenantId}-{scopes}`. These expire after ~1 hour.
3. **Client ID**: The Teams SPA's public client ID, found in the MSAL cache metadata.

The OAuth2 token endpoint accepts refresh token grants from public clients (no client secret needed):

```
POST https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

grant_type=refresh_token
&client_id={teams_client_id}
&refresh_token={refresh_token}
&scope={target_scopes}
```

This returns a new access token + refresh token pair.

### Implementation Plan

1. **Extract MSAL cache from session state**: Parse `session-state.json` localStorage entries to find refresh tokens, client ID, tenant ID, and account info. The MSAL cache schema is documented at https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/cache-schema.md

2. **New module `src/auth/token-refresh-http.ts`**: Direct HTTP token refresh:
   - Extract refresh token + client ID + tenant ID from session state
   - POST to `https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token`
   - Parse response for new access token + refresh token
   - Update the session state's localStorage entries with the new tokens (maintaining MSAL's cache format so the browser profile stays in sync)
   - Write updated session state back to encrypted storage

3. **Scopes to refresh**: We need tokens for multiple scopes:
   - `https://substrate.office.com/.default` (Substrate/search)
   - `https://api.spaces.skype.com/.default` (Calendar/Skype Spaces)
   - The `skypetoken_asm` cookie may need separate handling (it's a cookie, not an MSAL token)

4. **Update `token-refresh.ts`**: Try HTTP refresh first, fall back to browser-based refresh if it fails (e.g., refresh token expired, MSAL cache format changed).

5. **Update session state after HTTP refresh**: The tricky part — we need to write the new tokens back into the session state's localStorage in MSAL's exact cache format, so that:
   - `token-extractor.ts` can find them (it reads from session state)
   - The persistent browser profile stays consistent if a browser is opened later

### Risks & Unknowns

- **MSAL cache format stability**: The localStorage schema may change between MSAL versions. Need to handle gracefully.
- **Refresh token scope**: A single refresh token may or may not work for all required scopes. May need multiple refresh calls.
- **`skypetoken_asm` cookie**: This is used for messaging/chatsvc APIs and isn't an MSAL token. It may require a separate refresh mechanism or may be derivable from one of the MSAL tokens.
- **Conditional Access / MFA**: Some tenants may require additional claims that only a browser-based flow can satisfy.
- **Rate limiting**: Azure AD may rate-limit token endpoint calls.

### Research Steps

1. Examine the MSAL localStorage entries in a real session state to map out the exact cache structure
2. Identify the Teams client ID and required scopes from the cache
3. Test a manual `curl` call to the token endpoint with an extracted refresh token
4. Determine how `skypetoken_asm` is obtained (is it exchanged from an MSAL token, or set by a separate auth flow?)