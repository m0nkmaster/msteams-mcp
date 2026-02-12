# Clean Code Review: msteams-mcp

**Date**: 2026-02-12
**Branch**: `refactor/clean-code-review`
**Scope**: ~7,500 lines across 30 source files
**Tests**: 117 → 179 passing

---

## Phase 1: Scope

MCP server enabling AI assistants to interact with Microsoft Teams via direct API calls (Substrate, chatsvc, CSA) with browser-based authentication.

### Directory Structure

```
src/
├── index.ts              # Entry point
├── server.ts             # MCP server (TeamsServer class)
├── constants.ts          # Shared constants
├── tools/                # Tool handlers (6 files, ~3,200 lines)
├── api/                  # API clients (7 files, ~3,000 lines)
├── auth/                 # Authentication (5 files, ~1,300 lines)
├── browser/              # Playwright login (2 files, ~620 lines)
├── utils/                # Parsers, HTTP, config (4 files, ~1,600 lines)
├── types/                # Interfaces (4 files, ~270 lines)
└── __fixtures__/         # Test data (1 file)
```

---

## Phase 2: Readability & Clarity

### Strengths

- **Excellent module organisation** — clear separation into `api/`, `auth/`, `browser/`, `tools/`, `types/`, `utils/` with barrel exports
- **Consistent naming** — functions, types, and files follow clear conventions (`extractX`, `parseX`, `requireX`, `getX`)
- **Result types** — `Result<T, McpError>` forces explicit error handling everywhere
- **Error taxonomy** — machine-readable `ErrorCode` enum with LLM-friendly `suggestions` arrays
- **Section dividers** — consistent `// ───────` separators in larger files aid navigation
- **JSDoc** — thorough on public APIs with `@param`, `@returns`, `@example`, `@see` where appropriate

### Issues

#### R1. Reaction functions used inline headers instead of shared helper ✅ FIXED

`addReaction` and `removeReaction` in `chatsvc-api.ts` manually constructed headers identical to `getSkypeAuthHeaders()`. Every other chatsvc function used the shared helpers.

**Fix**: Unified into `setReaction` helper using `getSkypeAuthHeaders()`.

#### R2. `chatsvc-api.ts` is 1,680 lines — the largest file by far

Handles messaging, threads, saved messages, followed threads, chat creation, consumption horizons, activity feed, and reactions. Consider splitting into sub-modules (e.g., `chatsvc-messaging.ts`, `chatsvc-activity.ts`, `chatsvc-reactions.ts`).

**Status**: Deferred — added to ROADMAP.md as P4.

#### R3. `message-tools.ts` is 1,079 lines

Contains 18 tool definitions with handlers. Could be split along the same lines as the API layer.

**Status**: Deferred — grouped with R2 in ROADMAP.md.

#### R4. Inconsistent `as` casting style in API response parsing

Throughout `chatsvc-api.ts` and `files-api.ts`, raw API responses are parsed with extensive `as Record<string, unknown>` casts. Some functions (e.g., `getSharedFiles`) have 20+ casts. Consider defining lightweight response interfaces for common shapes.

**Status**: Deferred — added to ROADMAP.md as P4.

#### R5. `index.ts` entry point comment was stale ✅ FIXED

Said "search Microsoft Teams messages using browser automation" — the server does far more and uses direct API calls.

**Fix**: Updated to describe the full scope.

---

## Phase 3: Duplication

#### D1. `getApiConfig()` was defined identically in two files ✅ FIXED

Identical private function in `chatsvc-api.ts` (line 46) and `csa-api.ts` (line 23):

```typescript
function getApiConfig() {
  return { region: getRegion(), baseUrl: getTeamsBaseUrl() };
}
```

**Fix**: Exported from `auth-guards.ts` with `ApiConfig` interface. Both files now import it.

#### D2. `addReaction` / `removeReaction` were ~90% identical ✅ FIXED

Only differences: HTTP method (`PUT` vs `DELETE`) and body (`{ key, value: Date.now() }` vs `{ key }`).

**Fix**: Extracted `setReaction(conversationId, messageId, emojiKey, method)` private helper.

#### D3. Auth-guard + `getApiConfig` boilerplate repeated in every API function

Every function in `chatsvc-api.ts` (16 times) follows this pattern:

```typescript
const authResult = requireMessageAuth();
if (!authResult.ok) return authResult;
const auth = authResult.value;
const { region, baseUrl } = getApiConfig();
```

4 lines × 16 functions = 64 lines of pure boilerplate. A helper like `requireMessageAuthWithConfig()` returning `{ auth, region, baseUrl }` would eliminate this.

**Status**: Deferred — added to ROADMAP.md as P3. Too many call sites to change safely in a single review PR.

#### D4. Substrate token cache-clearing on `AUTH_EXPIRED` repeated 5 times

In `substrate-api.ts` and `files-api.ts`:

```typescript
if (!response.ok) {
  if (response.error.code === ErrorCode.AUTH_EXPIRED) {
    clearTokenCache();
  }
  return response;
}
```

**Status**: Deferred — added to ROADMAP.md as P3. Could be a wrapper or middleware in `http.ts`.

#### D5. `getSavedMessages` and `getFollowedThreads` were structurally identical ✅ FIXED

Both fetched from a virtual conversation, parsed with `parseVirtualConversationMessage`, and returned the same shape. Differed only in conversation ID and regex pattern.

**Fix**: Extracted `fetchVirtualConversation(virtualConversationId, referencePattern, options)` shared helper. Each public function is now a thin wrapper that maps `sourceReferenceId` to the appropriate field name (`sourceMessageId` vs `sourcePostId`).

---

## Phase 4: Simplification

#### S1. `extractObjectId` had repetitive branch structure ✅ FIXED

The GUID-test-then-base64-decode logic was repeated 4 times for different prefix formats.

**Fix**: Extracted `resolveIdPart(idPart)` closure:

```typescript
const resolveIdPart = (idPart: string): string | null => {
  if (guidPattern.test(idPart)) return idPart.toLowerCase();
  if (isLikelyBase64Guid(idPart)) return decodeBase64Guid(idPart);
  return null;
};
```

Each branch now calls `resolveIdPart(identifier.substring(N))` — function went from 56 lines to 30.

#### S2. `Promise.resolve(requireMessageAuth())` in `files-api.ts` ✅ FIXED

`requireMessageAuth()` is synchronous, so wrapping it in `Promise.resolve()` just to use `Promise.all` was unnecessary.

**Fix**: Call it directly first, then `await` the async token check.

#### S3. `getSubstrateTokenStatus` duplicates logic from `calculateTokenStatus`

In `token-extractor.ts`, `getSubstrateTokenStatus()` (line 184) and `getMessageAuthStatus()` (line 546) manually compute `expiresAt` and `minutesRemaining` — the same calculation that `calculateTokenStatus()` in `parsers.ts` already provides.

**Status**: Not fixed — low impact, would require changing the return type signatures.

#### S4. `hasMarkdownFormatting` checks for newlines

```typescript
if (/\n/.test(text)) return true;
```

Any message with a newline is treated as "has markdown formatting." Technically correct for the conversion pipeline, but the function name is slightly misleading — it's really `needsHtmlConversion`.

**Status**: Not fixed — renaming would touch callers and the function works correctly as-is.

---

## Phase 5: Documentation Accuracy

#### Doc1. `index.ts` module comment was outdated ✅ FIXED

Updated from "search Microsoft Teams messages using browser automation" to describe the full scope.

#### Doc2. `server.ts` module comment was outdated ✅ FIXED

Updated from "Teams search" to "Microsoft Teams" with full capability list.

#### Doc3. `AGENTS.md` referenced non-existent `test/` directory ✅ FIXED

Removed from the directory tree.

#### Doc4. `createGroupChat` JSDoc had stale `@param region` ✅ FIXED

The function no longer accepts a `region` parameter — removed from JSDoc.

#### Doc5. Unused interfaces in `types/teams.ts` ✅ FIXED

Removed 7 interfaces only referenced in their own definition file:
- `TeamsMessage`
- `TeamsReaction`
- `InterceptedRequest`
- `InterceptedResponse`
- `SearchApiEndpoint`
- `SearchPaginationOptions`
- `TeamsSearchResultsWithPagination`

#### Doc6. `AGENTS.md` had duplicate item 7 ✅ FIXED

Two items numbered "7" in the Implementation Patterns list. Renumbered 7–10 → 7–11.

---

## Phase 6: Test Coverage

### Existing coverage (before review)

`parsers.test.ts` — **117 tests** covering: `stripHtml`, `extractLinks`, `buildMessageLink`, `getConversationType`, `extractMessageTimestamp`, `decodeBase64Guid`, `parsePersonSuggestion`, `parseV2Result`, `parseJwtProfile`, `calculateTokenStatus`, `parseSearchResults`, `parsePeopleResults`, `extractObjectId`, `buildOneOnOneConversationId`, `extractActivityTimestamp`, `markdownToTeamsHtml`, `formatTranscriptText`.

### New tests added (+62) ✅

#### `parsers.test.ts` (+30 tests)

- **`parseChannelSuggestion`** (3 tests): complete suggestion, missing fields → null, default channelType
- **`parseChannelResults`** (3 tests): groups structure, undefined input, non-channel entities skipped
- **`parseTeamsList`** (4 tests): teams with channels, channelType mapping (0→Standard, 1→Private, 2→Shared), undefined/missing data, missing required fields
- **`filterChannelsByName`** (3 tests): case-insensitive partial match, cross-team results, no match
- **`parseVirtualConversationMessage`** (7 tests): saved message parsing, followed thread parsing, Control messages → null, missing id → null, missing timestamp → null, missing secondaryReferenceId, link building with context, HTML link extraction
- **`hasMarkdownFormatting`** (10 tests): bold, italic, strikethrough, inline code, code blocks, unordered lists, ordered lists, newlines, plain text → false

#### `errors.test.ts` (22 tests) — NEW FILE

- **`createError`** (8 tests): default suggestions per code, retryable defaults (RATE_LIMITED/NETWORK_ERROR/TIMEOUT/API_ERROR → true, AUTH_REQUIRED/INVALID_INPUT → false), custom overrides, retryAfterMs
- **`classifyHttpError`** (10 tests): 401→AUTH_EXPIRED, 403→AUTH_REQUIRED, 404→NOT_FOUND, 429→RATE_LIMITED, 400/422→INVALID_INPUT, 500+→API_ERROR, timeout/network messages, unknown status
- **`extractRetryAfter`** (3 tests): numeric seconds→ms conversion, missing header→undefined, HTTP date format

#### `crypto.test.ts` (10 tests) — NEW FILE

- **encrypt/decrypt round-trip** (5 tests): simple string, JSON data, empty string, unicode content, unique IV per encryption
- **`isEncrypted`** (3 tests): encrypted data → true, partial/null/string → false, manual object with all fields → true
- **decrypt error handling** (2 tests): tampered ciphertext → throw, tampered auth tag → throw

### Remaining gaps (not addressed)

- **T7**: No integration/tool-level tests — `invokeTool` with Zod validation in `registry.ts`
- **T8**: No tests for `httpRequest` retry/timeout logic — requires mocking `fetch`

Both added to ROADMAP.md as P4.

---

## Phase 7: Summary

### Overall Assessment: **Good** 🟢

Well-structured, well-documented codebase with strong architectural decisions. The code is clearly written by someone who understands both the problem domain and clean code principles.

### What Was Fixed (this PR)

| # | Category | Item | Lines saved |
|---|----------|------|-------------|
| D1 | Duplication | Shared `getApiConfig()` | ~14 |
| D2 | Duplication | Unified reaction functions | ~50 |
| D5 | Duplication | Unified virtual conversation fetcher | ~60 |
| R1 | Readability | Reaction headers use shared helper | — |
| S1 | Simplification | `resolveIdPart` in `extractObjectId` | ~26 |
| S2 | Simplification | Removed `Promise.resolve` wrapper | ~4 |
| Doc1-6 | Documentation | Fixed stale comments, numbering, unused types | ~55 |
| T1-T6 | Tests | +62 new tests across 3 files | +450 |

**Net**: +706 / -256 lines (mostly tests), 179 tests passing

### What's Deferred (in ROADMAP.md)

| # | Priority | Item | Reason |
|---|----------|------|--------|
| D3 | P3 | Auth+config boilerplate helper | 16 call sites — too invasive for review PR |
| D4 | P3 | Substrate error response wrapper | 5 call sites across 2 files |
| R2/R3 | P4 | Split large files | Structural change, needs careful planning |
| R4 | P4 | Typed API response interfaces | Low urgency, incremental improvement |
| T7/T8 | P4 | Integration & HTTP retry tests | Requires fetch mocking infrastructure |

### What's Done Well (no changes needed)

- **Architecture** — Clean separation of concerns, barrel exports, tool registry pattern
- **Error handling** — Consistent `Result<T>` types with machine-readable error codes
- **Security** — AES-256-GCM encryption at rest, restrictive file permissions, no hardcoded secrets
- **Testing** — Thorough pure-function tests with realistic fixtures
- **Documentation** — AGENTS.md is an excellent onboarding document; JSDoc is thorough on public APIs
- **Constants** — No magic numbers; all thresholds centralised in `constants.ts`
- **Auth flow** — Persistent browser profile with headless-first strategy is elegant
- **Region support** — Dynamic config extraction supports commercial, GCC, GCC-High, DoD environments
