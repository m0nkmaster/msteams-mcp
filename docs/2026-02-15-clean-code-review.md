# Clean Code Review — Full Codebase

**Date:** 2026-02-15
**Scope:** Entire `src/` directory (~30 source files, 7 test files, ~135 exported functions)
**Reviewer:** Cascade

---

## Review Progress

- [x] 1. Scope identification
- [x] 2. Readability & clarity
- [x] 3. Duplication detection
- [x] 4. Simplification opportunities
- [x] 5. Documentation accuracy
- [x] 6. Test coverage gaps
- [x] 7. Summary & recommendations

---

## Overall Assessment

The codebase is **well-structured and professionally maintained**. Architecture is modular, naming is consistent, error handling uses proper Result types, and the code follows a clear pattern throughout. The review found **no critical bugs or security issues**. Findings are primarily quality improvements and minor DRY violations.

---

## Critical Issues (must fix)

None found.

---

## Improvements (should fix)

### 1. Duplicate `escapeHtml` functions

🟡 **IMPROVE** `src/api/chatsvc-messaging.ts:801` + `src/utils/parsers.ts:1114`

Two independent implementations of the same HTML escape function:

**chatsvc-messaging.ts:**
```typescript
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
```

**parsers.ts:**
```typescript
function escapeHtmlChars(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
```

These are byte-identical logic. Export one from `parsers.ts` and import it in `chatsvc-messaging.ts`.

*Reason: DRY violation. Two copies means a fix to one won't propagate to the other.*

---

### 2. `handleGetUnread` sequential API calls for favourites

🟡 **IMPROVE** `src/tools/message-tools.ts:656-744`

The aggregate unread handler iterates over favourites **sequentially** with `for...of` + `await`:

```typescript
for (const fav of favorites.slice(0, maxToCheck)) {
  const unreadResult = await getUnreadStatus(fav.conversationId);
  // ...
}
```

This means checking 30 favourites makes 30 serial HTTP calls. Using `Promise.allSettled` with a concurrency limit (e.g., 5-10 at a time) would significantly reduce latency.

*Reason: Performance. Users with many pinned chats will experience slow response times.*

---

### 3. `parsers.ts` is 1422 lines — consider splitting

🟡 **IMPROVE** `src/utils/parsers.ts`

This single file contains 28 exported functions spanning unrelated concerns:
- HTML/link parsing
- Search result parsing (Teams + Email)
- People/channel parsing
- JWT profile parsing
- Token status calculation
- Markdown-to-HTML conversion
- Transcript formatting
- GUID decoding/extraction
- Virtual conversation parsing

Suggested split:
- `parsers/html.ts` — `stripHtml`, `extractLinks`, `escapeHtmlChars`
- `parsers/search.ts` — `parseV2Result`, `parseSearchResults`, `parseEmailResult`, `parseEmailSearchResults`, `classifyEmailType`
- `parsers/people.ts` — `parsePersonSuggestion`, `parsePeopleResults`, `extractObjectId`, `decodeBase64Guid`
- `parsers/channels.ts` — `parseChannelSuggestion`, `parseChannelResults`, `parseTeamsList`, `filterChannelsByName`
- `parsers/markdown.ts` — `markdownToTeamsHtml`, `hasMarkdownFormatting`
- `parsers/index.ts` — barrel re-exports

The test file mirrors this and could split correspondingly.

*Reason: The file has grown large enough that navigation and cognitive load are affected. Each submodule would be under 200 lines.*

---

### 4. `token-extractor.ts` — repeated `getTeamsLocalStorage` + null check pattern

🟡 **IMPROVE** `src/auth/token-extractor.ts`

10 functions all begin with the same boilerplate:

```typescript
const localStorage = getTeamsLocalStorage(state);
if (!localStorage) return null;
```

Consider a higher-order function wrapper:

```typescript
function withLocalStorage<T>(
  state: SessionState | undefined,
  fn: (localStorage: LocalStorageEntry[]) => T | null
): T | null {
  const localStorage = getTeamsLocalStorage(state);
  if (!localStorage) return null;
  return fn(localStorage);
}
```

Then each function becomes:
```typescript
export function extractSubstrateToken(state?: SessionState): SubstrateTokenInfo | null {
  return withLocalStorage(state, (localStorage) => {
    // ... core logic only
  });
}
```

*Reason: Reduces 10 instances of identical 2-line boilerplate.*

---

### 5. Missing return type annotations on some public functions

🟡 **IMPROVE** Various files

Most public functions have explicit return types, but some in `auth-guards.ts` and `token-extractor.ts` rely on inference for complex return types. For API contract clarity, all exported functions should have explicit return types.

Examples:
- `getSubstrateTokenStatus()` — return type is inferred as a complex object literal
- `getMessageAuthStatus()` — same pattern
- `discoverConfig()` — returns `DiscoveredConfig | null` but relies on inference

*Reason: Explicit return types act as documentation and catch accidental signature changes.*

---

### 6. `README.md` references non-existent CLI commands

🟡 **IMPROVE** `README.md:179-221`

The "CLI Tools" section references `npm run cli` commands, but the actual harness is at `src/test/mcp-harness.ts`. The README also mentions `npm run test:mcp` but `package.json` defines `"cli": "tsx src/test/mcp-harness.ts"` — there is no `test:mcp` script.

*Reason: Users following the README would get "missing script" errors.*

---

### 7. `README.md` session files section mentions `.user-data/` but actual directory is `browser-profile/`

🟡 **IMPROVE** `README.md:238`

```
Contents: `session-state.json`, `token-cache.json`, `.user-data/`
```

The actual persistent browser profile directory is `browser-profile/` (set in `context.ts`), not `.user-data/`.

*Reason: Misleading documentation for users troubleshooting session issues.*

---

## Suggestions (nice to have)

### 8. `registry.ts` — the 3 `eslint-disable` comments could be eliminated

🟢 **SUGGEST** `src/tools/registry.ts:22,33,52`

The `any` type usage could be replaced with a base `RegisteredTool<ZodTypeAny>` type or a `RegisteredToolBase` interface that avoids the need for `any`. Since these are the only `eslint-disable` comments in production code, eliminating them would enforce zero-suppression policy.

---

### 9. Tool handler files have very long tool definition objects

🟢 **SUGGEST** `src/tools/message-tools.ts`, `src/tools/search-tools.ts`

Each tool has its JSON Schema `inputSchema` defined inline in the tool definition object, making files like `message-tools.ts` (1079 lines) very long. Consider extracting schemas to a shared `schemas/` directory or co-locating them in a separate file per group.

*Reason: The tool definition objects are largely static metadata that adds visual noise to the handler logic.*

---

### 10. `buildMessageLink` has a deprecated overload that could be removed

🟢 **SUGGEST** `src/utils/parsers.ts:124-132`

```typescript
/**
 * @deprecated Use the options object overload instead.
 */
export function buildMessageLink(
  conversationId: string,
  messageTimestamp: string | number,
  parentMessageId?: string,
  teamsBaseUrl?: string
): string;
```

All internal callers already use the options object form. The deprecated positional overload could be removed to simplify the implementation (the multi-signature resolution logic in lines 133-170 would collapse to a single clean function).

---

### 11. Console logging could use a lightweight logger abstraction

🟢 **SUGGEST** Various files (23 `console.error` calls in production code)

Production code uses bare `console.error()` for logging. A minimal `logger.ts` module with level control (e.g., `logger.debug()`, `logger.warn()`, `logger.error()`) would allow:
- Silencing debug output in production
- Consistent `[module]` prefixing (some calls already do `[login:headless]`)
- Future structured logging

---

### 12. `chatsvc-messaging.ts:270-274` — IIFE for timestamp fallback is unusual

🟢 **SUGGEST** `src/api/chatsvc-messaging.ts:269-274`

```typescript
const timestamp = msg.originalarrivaltime ||
  msg.composetime ||
  (() => {
    const parsed = parseInt(id, 10);
    return !isNaN(parsed) && parsed > 0 ? new Date(parsed).toISOString() : new Date().toISOString();
  })();
```

A small helper like `parseTimestampOrNow(id)` would be clearer than the inline IIFE.

---

## Test Gaps

### High Priority

| Gap | File | Notes |
|-----|------|-------|
| **No tests for any API module** | `substrate-api.ts`, `chatsvc-messaging.ts`, `chatsvc-activity.ts`, `chatsvc-reactions.ts`, `chatsvc-virtual.ts`, `chatsvc-readstatus.ts`, `csa-api.ts`, `calendar-api.ts`, `transcript-api.ts`, `files-api.ts` | These are the core business logic modules. Would need mocked HTTP/auth but are the highest-value test targets. |
| **No tests for tool handlers** | `search-tools.ts`, `message-tools.ts`, `people-tools.ts`, `meeting-tools.ts`, `file-tools.ts` | Tool handlers contain non-trivial logic (unread aggregation, pagination, thread ordering). |
| **No tests for `server.ts`** | `server.ts` | Auto-login retry logic, concurrent auth deduplication, and error formatting are complex and untested. |

### Medium Priority

| Gap | File | Notes |
|-----|------|-------|
| **No tests for `token-extractor.ts`** | `auth/token-extractor.ts` | 15 exported functions, none tested directly. Pure functions like `extractSubstrateToken` could be tested with mock session state. |
| **No tests for `session-store.ts`** | `auth/session-store.ts` | 11 exported functions. Filesystem operations would need mocking but encryption round-trip and path logic are testable. |
| **No tests for `http.ts`** | `utils/http.ts` | Retry logic, timeout handling, and error classification are critical paths. |
| **No tests for `api-config.ts`** | `utils/api-config.ts` | URL construction with region/base URL parameterisation could have subtle bugs. |
| **No tests for `browser/auth.ts`** | `browser/auth.ts` | Harder to test (Playwright dependency), but auth detection regexes could be extracted and tested. |

### Existing Test Quality

The 7 existing test files are well-written:
- Good use of fixtures (`__fixtures__/api-responses.ts`)
- Edge cases covered (empty input, invalid formats, boundary values)
- Proper mocking in `token-refresh-http.test.ts`
- ~130+ test cases total with good coverage of parser logic

---

## Documentation Gaps

| Issue | Location | Notes |
|-------|----------|-------|
| `test:mcp` script referenced but doesn't exist | `README.md:211` | Should be `npm run cli` |
| `.user-data/` directory name is wrong | `README.md:238` | Should be `browser-profile/` |
| `AGENTS.md` is comprehensive and accurate | `AGENTS.md` | Well maintained — no issues found |
| `ROADMAP.md` is current | `ROADMAP.md` | Accurate reflection of planned features |
| No JSDoc on `types/server.ts` interface methods | `types/server.ts` | `ITeamsServer` interface methods lack descriptions |
| Tool `inputSchema` descriptions are excellent | All tool files | Every parameter has clear, actionable descriptions with examples |

---

## Architecture Observations (no action needed)

These are positive patterns worth preserving:

- **Result<T, E> everywhere** — consistent error handling across all API boundaries
- **Auth guard pattern** — `requireMessageAuth()`, `requireSubstrateTokenAsync()` etc. provide clean auth checks
- **Tool registry** — clean separation between MCP protocol, tool definitions, and API calls
- **Encrypted credentials at rest** — AES-256-GCM with machine-specific key derivation
- **HTTP-first token refresh** — ~1s vs ~8s browser fallback, transparent to callers
- **Constants centralisation** — no magic numbers in business logic
- **Barrel exports** — clean module boundaries via index.ts files
