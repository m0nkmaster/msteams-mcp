# Microsoft Teams Session Data Reference

Complete reference for the session data extracted from the Teams web client. This includes cookies, localStorage entries, and configuration discovered during browser authentication.

## Table of Contents

1. [Overview](#overview)
2. [Session State Structure](#session-state-structure)
3. [Cookies](#cookies)
4. [localStorage - MSAL Tokens](#localstorage---msal-tokens)
5. [localStorage - Discovery Config](#localstorage---discovery-config)
6. [localStorage - User Details](#localstorage---user-details)
7. [localStorage - Feature Flags](#localstorage---feature-flags)
8. [Token Types](#token-types)
9. [Multi-Tenant Support](#multi-tenant-support)
10. [Debugging Session Issues](#debugging-session-issues)

---

## Overview

The Teams MCP server extracts authentication tokens and configuration from the browser's session state after login. This avoids using Microsoft Graph API in favour of the same internal APIs the Teams web client uses.

**Session data is stored in:**
- `~/.teams-mcp-server/session-state.json` (encrypted)
- `~/.teams-mcp-server/token-cache.json` (encrypted)

**Key principle:** All tenant-specific configuration (region, partition, base URLs) is extracted from the session, not hardcoded. This enables support for different Teams environments (commercial, GCC, GCC-High, DoD).

---

## Session State Structure

Playwright's `storageState()` captures the following structure:

```typescript
interface SessionState {
  cookies: Array<{
    name: string;
    value: string;
    domain?: string;
    path?: string;
    expires?: number;
    httpOnly?: boolean;
    secure?: boolean;
    sameSite?: 'Strict' | 'Lax' | 'None';
  }>;
  origins: Array<{
    origin: string;  // e.g., "https://teams.microsoft.com"
    localStorage: Array<{ name: string; value: string }>;
  }>;
}
```

**Known Teams Origins:**
| Origin | Environment |
|--------|-------------|
| `https://teams.microsoft.com` | Commercial |
| `https://teams.microsoft.us` | GCC-High |
| `https://dod.teams.microsoft.us` | DoD |
| `https://teams.cloud.microsoft` | New Teams URL |

---

## Cookies

### skypetoken_asm

**Purpose:** Authentication for chatsvc messaging APIs

**Domain:** `teams.microsoft.com` (or government cloud equivalent)

**Format:** JWT token

**Used by:** 
- Messaging APIs (`/api/chatsvc/{region}/...`)
- Calendar APIs (`/api/mt/part/{region}/...`)
- CSA APIs (`/api/csa/{region}/...`)

**Header format:**
```
Authentication: skypetoken={value}
```

**JWT Claims:**
```json
{
  "skypeid": "orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
  "exp": 1738454636,
  "iat": 1738368236
}
```

**Expiry:** ~24 hours (longer than MSAL tokens)

### authtoken

**Purpose:** Bearer token for certain API calls

**Domain:** `teams.microsoft.com`

**Format:** URL-encoded, may have `Bearer=` prefix

**Used by:** Some chatsvc APIs alongside skypetoken_asm

**Processing:**
```typescript
let authToken = decodeURIComponent(rawAuthToken);
if (authToken.startsWith('Bearer=')) {
  authToken = authToken.substring(7);
}
```

---

## localStorage - MSAL Tokens

MSAL (Microsoft Authentication Library) stores OAuth tokens in localStorage with complex key names containing the user ID, tenant ID, and scope.

### Token Entry Structure

Each MSAL token entry is JSON with this structure:

```json
{
  "target": "https://substrate.office.com/SubstrateSearch-Internal.ReadWrite",
  "secret": "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6...",
  "cachedAt": "1738368236",
  "expiresOn": "1738372136",
  "extendedExpiresOn": "1738375736"
}
```

### Token Types by Target/Scope

| Target Contains | Token Type | Used For |
|-----------------|------------|----------|
| `substrate.office.com/SubstrateSearch` | Substrate Search | Message search, people search |
| `chatsvcagg.teams.microsoft.com` | CSA Token | Favourites, teams list |
| `api.spaces.skype.com` | Spaces Token | Calendar/meetings API |

### Finding Tokens

```typescript
// Substrate search token
for (const item of localStorage) {
  const entry = JSON.parse(item.value);
  if (entry.target?.includes('substrate.office.com') && 
      entry.target?.includes('SubstrateSearch')) {
    // entry.secret is the JWT token
  }
}

// CSA token
for (const item of localStorage) {
  if (item.name.includes('chatsvcagg.teams.microsoft.com')) {
    const entry = JSON.parse(item.value);
    // entry.secret is the token
  }
}

// Spaces token (for calendar)
for (const item of localStorage) {
  const entry = JSON.parse(item.value);
  if (entry.target?.includes('api.spaces.skype.com')) {
    // entry.secret is the JWT token
  }
}
```

### JWT Token Claims

All MSAL tokens are JWTs with standard claims:

```json
{
  "aud": "https://substrate.office.com",
  "iss": "https://sts.windows.net/{tenant-id}/",
  "iat": 1738368236,
  "exp": 1738372136,
  "oid": "ab76f827-27e2-4c67-a765-f1a53145fa24",
  "upn": "user@company.com",
  "name": "John Smith",
  "tid": "56b731a8-a2ac-4c32-bf6b-616810e913c6"
}
```

**Key claims for our use:**
- `oid` - User's object ID (used to construct MRI: `8:orgid:{oid}`)
- `upn` - User's email address
- `name` - Display name
- `tid` - Tenant ID
- `exp` - Expiry timestamp (seconds since epoch)

---

## localStorage - Discovery Config

### DISCOVER-REGION-GTM

**Key pattern:** `tmp.auth.v1.{userId}.Discover.DISCOVER-REGION-GTM`

**Purpose:** Contains all regional API endpoint URLs

**Structure:**
```json
{
  "item": {
    "middleTier": "https://teams.microsoft.com/api/mt/part/amer-02",
    "chatServiceAfd": "https://teams.microsoft.com/api/chatsvc/amer",
    "chatSvcAggAfd": "https://teams.microsoft.com/api/csa/amer",
    "chatServiceAggregator": "https://chatsvcagg.teams.microsoft.com",
    "unifiedPresence": "https://presence.teams.microsoft.com",
    "search": "https://us-prod.asyncgw.teams.microsoft.com/msgsearch",
    "substrateSyncS2S": "https://amer.substratesync.teams.microsoft.com",
    "calling_trouterUrl": "https://go.trouter.teams.microsoft.com/v3/c",
    "ams": "https://us-api.asm.skype.com",
    "amsV2": "https://us-prod.asyncgw.teams.microsoft.com",
    // ... 80+ more URLs
  },
  "shouldRefresh": false,
  "hitCount": 0
}
```

**Key URLs we extract:**

| Field | Purpose | Example |
|-------|---------|---------|
| `middleTier` | Calendar API base URL | `https://teams.microsoft.com/api/mt/part/amer-02` |
| `chatServiceAfd` | Messaging API base URL | `https://teams.microsoft.com/api/chatsvc/amer` |
| `chatSvcAggAfd` | CSA API base URL | `https://teams.microsoft.com/api/csa/amer` |

**Partition extraction:**

- **Partitioned tenants:** `middleTier` ends with `/part/{region}-{partition}` (e.g., `amer-02`)
- **Non-partitioned tenants:** `middleTier` ends with `/{region}` (e.g., `emea`)

```typescript
// Partitioned format
const partitionMatch = middleTierUrl.match(/\/api\/mt\/part\/([a-z]+)-(\d+)$/);
// partitionMatch[1] = "amer", partitionMatch[2] = "02"

// Non-partitioned format
const simpleMatch = middleTierUrl.match(/\/api\/mt\/([a-z]+)$/);
// simpleMatch[1] = "emea"
```

**Base URL extraction:**

The Teams base URL (for GCC/GCC-High support) is extracted from any of the full URLs:

```typescript
const url = new URL(chatServiceAfd); // "https://teams.microsoft.com/api/chatsvc/amer"
const teamsBaseUrl = `${url.protocol}//${url.host}`; // "https://teams.microsoft.com"
```

---

## localStorage - User Details

### DISCOVER-USER-DETAILS

**Key pattern:** `tmp.auth.v1.{userId}.Discover.DISCOVER-USER-DETAILS`

**Purpose:** User-specific configuration and license information

**Structure:**
```json
{
  "item": {
    "id": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
    "region": "amer",
    "userRegion": "amer",
    "partition": "amer02",
    "userPartition": "amer01",
    "ocdiRedirect": "MicrosoftDefault",
    "licenseDetails": {
      "isFreemium": false,
      "isTrial": false,
      "isTeamsEnabled": true,
      "isCopilot": true,
      "isM365CopilotBusinessChat": true,
      "isCopilotApps": true,
      "isTranscriptEnabled": true,
      "isWebinarEnabled": true,
      "isAudioConf": true,
      "isFrontline": false,
      "isAvatarsEnabled": true,
      "isTeamsDeviceManagementEnabled": true,
      "experimentationTeamsAddOnPlans": ["Copilot"]
    },
    "regionSettings": {
      "isUnifiedPresenceEnabled": true,
      "isOutOfOfficeIntegrationEnabled": true,
      "isContactMigrationEnabled": true,
      "isAppsDiscoveryEnabled": true,
      "isFederationEnabled": true
    }
  },
  "shouldRefresh": false,
  "hitCount": 0
}
```

**Key fields:**

| Field | Purpose | Notes |
|-------|---------|-------|
| `id` | User's MRI | `8:orgid:{guid}` format |
| `region` | User's region | `amer`, `emea`, `apac` |
| `partition` | Tenant partition | `amer02` (no hyphen, different from URL format) |
| `userPartition` | User's partition | May differ from tenant partition |
| `licenseDetails.isCopilot` | Copilot enabled | Boolean |
| `licenseDetails.isTranscriptEnabled` | Transcription enabled | Boolean |
| `licenseDetails.isFrontline` | Frontline worker | Boolean |

**Note:** The `partition` here (`amer02`) differs from the URL format (`amer-02`). Always use the URL from `DISCOVER-REGION-GTM` for API calls.

---

## localStorage - Feature Flags

### experience-loader-ecs-flags

**Key pattern:** `tmp.react-web-client.experience-loader-ecs-flags`

**Purpose:** Feature flags and redirect configuration

**Structure:**
```json
{
  "enableLazyLoadedWorker": false,
  "workerChunkLoadMaxRetries": 10,
  "isTeams2025IconPackEnabled": false,
  "enableSplashScreen2025": true,
  "enableOCDIRedirects": false,
  "ocdiRedirectBaseUri": "https://teams.cloud.microsoft",
  "ocdiRedirectPath": "/",
  "ocdiRedirectCookieName": "ocdiRedirect"
}
```

**Key field:** `ocdiRedirectBaseUri` shows the new Teams URL format (`teams.cloud.microsoft`).

---

## Token Types

### Summary Table

| Token | Source | Expiry | Used For | Header |
|-------|--------|--------|----------|--------|
| **Substrate** | MSAL localStorage | ~1 hour | Search, People | `Authorization: Bearer {token}` |
| **CSA** | MSAL localStorage | ~1 hour | Favourites, Teams list | `Authorization: Bearer {token}` |
| **Spaces** | MSAL localStorage | ~1 hour | Calendar/Meetings | `Authorization: Bearer {token}` |
| **Skype** | Cookie `skypetoken_asm` | ~24 hours | Messaging, Threads | `Authentication: skypetoken={token}` |
| **Auth** | Cookie `authtoken` | ~24 hours | Some APIs | Combined with Skype token |

### Token Refresh

- **MSAL tokens:** Refresh automatically when Teams makes API calls requiring them. We trigger refresh by loading Teams in a headless browser and making a search request.
- **Cookies:** Longer-lived, typically don't need refresh during a session.

---

## Multi-Tenant Support

### Supported Environments

| Environment | Teams URL | Login URL |
|-------------|-----------|-----------|
| Commercial | `teams.microsoft.com` | `login.microsoftonline.com` |
| GCC | `teams.microsoft.com` | `login.microsoftonline.com` |
| GCC-High | `teams.microsoft.us` | `login.microsoftonline.us` |
| DoD | `dod.teams.microsoft.us` | `login.microsoftonline.us` |

### How It Works

1. **Login:** User navigates to `teams.microsoft.com`. Microsoft redirects to correct environment based on tenant.
2. **Session capture:** After login, we save the session state including the correct origin.
3. **Config extraction:** `DISCOVER-REGION-GTM` contains full URLs with the correct base domain.
4. **API calls:** We use the extracted base URL, not hardcoded `teams.microsoft.com`.

### Finding the Right Origin

```typescript
const TEAMS_ORIGINS = [
  'https://teams.microsoft.com',
  'https://teams.microsoft.us',
  'https://dod.teams.microsoft.us',
  'https://teams.cloud.microsoft',
];

function getTeamsOrigin(state: SessionState) {
  for (const knownOrigin of TEAMS_ORIGINS) {
    const origin = state.origins.find(o => o.origin === knownOrigin);
    if (origin) return origin;
  }
  // Fallback: find any origin containing 'teams.microsoft'
  return state.origins.find(o => 
    o.origin.includes('teams.microsoft') || o.origin.includes('teams.cloud')
  ) ?? null;
}
```

---

## Debugging Session Issues

### CLI Commands

```bash
# Check current session status
npm run cli -- status

# Dump all configuration from session
npm run cli -- dump-config

# Output as JSON for analysis
npm run cli -- dump-config --json

# Force re-login
npm run cli -- login --force
```

### Common Issues

#### "No valid token" for Search

The Substrate token expires after ~1 hour. Solutions:
1. Run `npm run cli -- login` to refresh
2. The server attempts automatic refresh when tokens are near expiry

#### "No region configuration found"

The session is missing `DISCOVER-REGION-GTM`. This can happen if:
1. Session is corrupted - delete `~/.teams-mcp-server/` and re-login
2. Login was interrupted before Teams fully loaded

#### Wrong region/partition

If API calls fail with 404 or redirect errors:
1. Run `npm run cli -- dump-config` to see extracted config
2. Check that region matches your tenant location
3. Delete session and re-login if incorrect

### Session File Locations

| OS | Location |
|----|----------|
| macOS/Linux | `~/.teams-mcp-server/` |
| Windows | `%APPDATA%\teams-mcp-server\` |

**Files:**
- `session-state.json` - Encrypted session (cookies + localStorage)
- `token-cache.json` - Encrypted token cache
- `.user-data/` - Browser profile directory

### Encryption

Session files are encrypted at rest using AES-256-GCM with a key derived from:
- Hostname
- Username

This means session files are machine-specific and cannot be copied between computers.

---

## Known Limitations

### Substrate URL

The Substrate search API URL (`substrate.office.com`) is not found in any localStorage config. It appears to be hardcoded in the Teams client. For government cloud tenants, this may need to be different.

**Current hardcoded value:**
```typescript
const DEFAULT_SUBSTRATE_BASE_URL = 'https://substrate.office.com';
```

If GCC/GCC-High users report search issues, we may need to:
1. Find an alternative config source
2. Add environment detection
3. Make it configurable

### Client Version

The `X-Ms-Client-Version` header is hardcoded:
```typescript
'X-Ms-Client-Version': '1415/1.0.0.2025010401'
```

This is extracted from observed network traffic. If Microsoft starts enforcing version checks, we may need to extract this dynamically.
