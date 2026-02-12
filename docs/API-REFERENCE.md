# Microsoft Teams Web API Reference

Complete reference for the undocumented Microsoft Teams APIs used by this MCP server. These are internal APIs discovered through browser network inspection, not the official Microsoft Graph API.

## Table of Contents

1. [Authentication](#authentication)
2. [Regional Variations](#regional-variations)
3. [Search APIs](#search-apis)
4. [People APIs](#people-apis)
5. [Virtual Conversations](#virtual-conversations)
6. [Messaging APIs](#messaging-apis)
7. [Conversation APIs](#conversation-apis)
8. [Activity & Notifications](#activity--notifications)
9. [Reactions & Emoji](#reactions--emoji)
10. [Calendar & Scheduling](#calendar--scheduling)
11. [Transcripts](#transcripts)
12. [Files & Attachments](#files--attachments)
13. [Common Gotchas](#common-gotchas)

---

## Authentication

Teams uses multiple authentication mechanisms depending on the API surface:

| Auth Type | Header | Source | Used By |
|-----------|--------|--------|---------|
| **Bearer (Substrate)** | `Authorization: Bearer {token}` | MSAL localStorage, `SubstrateSearch-Internal.ReadWrite` scope | Search, People |
| **Bearer (CSA)** | `Authorization: Bearer {csaToken}` | MSAL, `chatsvcagg.teams.microsoft.com` audience | Teams list, Favorites |
| **Bearer (Spaces)** | `Authorization: Bearer {spacesToken}` | MSAL, `api.spaces.skype.com` audience | Calendar/Meetings |
| **Skype Token** | `Authentication: skypetoken={token}` | Cookie `skypetoken_asm` | Messaging, Threads, Calendar |
| **Bearer (Substrate + Prefer)** | `Authorization: Bearer {token}` + `Prefer` header | Same as Substrate search token | Transcripts (WorkingSetFiles) |

### Required Headers

Most endpoints require these headers:

```
Origin: https://teams.microsoft.com
Referer: https://teams.microsoft.com/
Content-Type: application/json
```

### Token Storage (MSAL Cache)

Tokens are stored in browser localStorage under MSAL cache keys. Tokens typically expire after ~1 hour.

**MSAL Cache Key Formats:**

| Entry Type | Key Format | Example |
|------------|-----------|---------|
| **Refresh Token** | `{homeAccountId}-{environment}-refreshtoken-{clientId}----` | `ab76f827-...56b731a8-...-login.windows.net-refreshtoken-5e3ce6c0-...----` |
| **Access Token** | `{homeAccountId}-{environment}-accesstoken-{clientId}-{tenantId}-{scopes}` | `ab76f827-...-login.windows.net-accesstoken-5e3ce6c0-...-56b731a8-...-https://substrate.office.com/.default` |
| **Account** | `{homeAccountId}-{environment}-{tenantId}` | `ab76f827-...-login.windows.net-56b731a8-...` |

**Key Values:**
- **Client ID**: `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Microsoft Teams Web Client, registered as a SPA)
- **Environment**: `login.windows.net`
- **homeAccountId**: `{userObjectId}.{tenantId}`

**Access Token Entry Structure:**
```json
{
  "credentialType": "AccessToken",
  "homeAccountId": "{userOid}.{tenantId}",
  "environment": "login.windows.net",
  "clientId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
  "realm": "{tenantId}",
  "target": "https://substrate.office.com/SubstrateSearch-Internal.ReadWrite https://substrate.office.com/.default ...",
  "tokenType": "Bearer",
  "secret": "eyJ0eXAi...",
  "expiresOn": "1770938426",
  "extendedExpiresOn": "1770942075",
  "cachedAt": "1770934778"
}
```

**Refresh Token Entry Structure:**
```json
{
  "credentialType": "RefreshToken",
  "homeAccountId": "{userOid}.{tenantId}",
  "environment": "login.windows.net",
  "clientId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
  "secret": "{opaque_refresh_token}",
  "expiresOn": "1770943425"
}
```

The refresh token is opaque (not a JWT), ~1800 chars, and long-lived (days/weeks).

### Token Refresh (Browserless)

Tokens can be refreshed without a browser by POSTing directly to Azure AD's OAuth2 token endpoint:

```
POST https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded
Origin: https://teams.microsoft.com

grant_type=refresh_token
&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346
&refresh_token={refresh_token}
&scope=https://substrate.office.com/.default offline_access
```

**Critical:** The `Origin: https://teams.microsoft.com` header is **required**. The Teams client ID is registered as a Single-Page Application (SPA), and Azure AD validates that refresh token grants from SPA clients include a cross-origin `Origin` header. Without it, Azure AD returns error `AADSTS9002327`.

**Scopes to refresh:**

| Scope | Used For |
|-------|----------|
| `https://substrate.office.com/.default offline_access` | Search, People, Transcripts, Files |
| `https://api.spaces.skype.com/.default offline_access` | Calendar, Meetings + skypetoken_asm derivation |
| `https://chatsvcagg.teams.microsoft.com/.default offline_access` | Favorites, Teams list (CSA) |

Azure AD may rotate the refresh token on each use — always store the new `refresh_token` from the response.

### Skype Token Exchange

The `skypetoken_asm` cookie (used for messaging/chatsvc APIs) is obtained by exchanging a Skype Spaces access token:

```
POST https://authsvc.teams.microsoft.com/v1.0/authz
Authorization: Bearer {api.spaces.skype.com_access_token}
Content-Type: application/json

{}
```

**Response:**
```json
{
  "tokens": {
    "skypeToken": "eyJ0eXAi...",
    "expiresIn": 86400
  },
  "regionGtms": {
    "middleTier": "https://teams.microsoft.com/api/mt/part/amer-02",
    "chatServiceAfd": "https://teams.microsoft.com/api/chatsvc/amer",
    ...
  }
}
```

The `skypeToken` is set as the `skypetoken_asm` cookie on `.asyncgw.teams.microsoft.com` and `.asm.skype.com`. The `regionGtms` response also provides regional API URLs.

The `authtoken` cookie (on `teams.microsoft.com`) is the Skype Spaces access token itself, URL-encoded with a `Bearer=` prefix.

### Session Persistence

Session state includes:
- Cookies (`skypetoken_asm`, `authtoken` — for messaging APIs)
- localStorage (contains MSAL token cache)
- sessionStorage

---

## Regional Variations

API URLs include regional identifiers based on the user's tenant location:

| Region | Code |
|--------|------|
| Americas | `amer` |
| Europe/Middle East/Africa | `emea` |
| Asia Pacific | `apac` |

**Usage patterns:**
- `/api/csa/{region}/api/v1/...`
- `/api/chatsvc/{region}/v1/...`
- `/api/mt/part/{region}/beta/...`
- `nam.loki.delve.office.com` (North America Delve APIs)

---

## Search APIs

### Full-Text Message Search

**Endpoint:** `POST https://substrate.office.com/searchservice/api/v2/query`

**Auth:** Bearer (Substrate)

**Request:**
```json
{
  "entityRequests": [
    {
      "entityType": "Message",
      "contentSources": ["Teams"],
      "fields": [
        "Extension_SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId_String",
        "Extension_SkypeSpaces_ConversationPost_Extension_FileData_String",
        "Extension_SkypeSpaces_ConversationPost_Extension_ThreadType_String"
      ],
      "propertySet": "Optimized",
      "query": {
        "queryString": "search term AND NOT (isClientSoftDeleted:TRUE)",
        "displayQueryString": "search term"
      },
      "from": 0,
      "size": 25,
      "topResultsCount": 5
    }
  ],
  "QueryAlterationOptions": {
    "EnableAlteration": true,
    "EnableSuggestion": true,
    "SupportedRecourseDisplayTypes": ["Suggestion"]
  },
  "cvid": "uuid",
  "logicalId": "uuid",
  "scenario": {
    "Dimensions": [
      {"DimensionName": "QueryType", "DimensionValue": "Messages"},
      {"DimensionName": "FormFactor", "DimensionValue": "general.web.reactSearch"}
    ],
    "Name": "powerbar"
  }
}
```

**Response:**
```json
{
  "EntitySets": [
    {
      "ResultSets": [
        {
          "Total": 4307,
          "Results": [
            {
              "Id": "AAMkA...",
              "ReferenceId": "uuid.1000.1",
              "HitHighlightedSummary": "Message with <c0>highlights</c0>...",
              "Source": {
                "Summary": "Plain text content",
                "From": {
                  "EmailAddress": {
                    "Name": "Smith, John",
                    "Address": "john.smith@company.com"
                  }
                }
              }
            }
          ]
        }
      ]
    }
  ]
}
```

**Pagination:**
- `from`: Starting offset (0, 25, 50, ...)
- `size`: Page size (default 25, max ~50)
- Response includes `Total` count

**Search Operators:**

| Operator | Example | Description |
|----------|---------|-------------|
| `from:` | `from:john.smith@company.com` | Messages from a person |
| `in:` | `in:general` | Messages in a channel |
| `sent:` | `sent:2026-01-20`, `sent:>=2026-01-15` | By date (explicit dates only) |
| `subject:` | `subject:budget` | In message subject |
| `"Name"` | `"Smith, John"` | Find @mentions (name in quotes) |
| `hasattachment:true` | - | Messages with files |
| `NOT` | `NOT from:user@co.com` | Exclude results |

**Finding @mentions:**
```
"Macdonald, Rob"              # Find mentions of you
"Macdonald, Rob" from:diego   # Mentions from a specific person
```

---

### People Search (Autocomplete)

**Endpoint:** `POST https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar`

**Auth:** Bearer (Substrate)

**Request:**
```json
{
  "EntityRequests": [
    {
      "Query": {
        "QueryString": "rob",
        "DisplayQueryString": "rob"
      },
      "EntityType": "People",
      "Size": 5,
      "Fields": ["Id", "MRI", "DisplayName", "EmailAddresses", "JobTitle", "Department"]
    }
  ]
}
```

**Response:**
```json
{
  "Groups": [
    {
      "Suggestions": [
        {
          "Id": "uuid@tenant",
          "DisplayName": "Smith, John",
          "GivenName": "John",
          "Surname": "Smith",
          "EmailAddresses": ["user@company.com"],
          "CompanyName": "Company Name",
          "Department": "Engineering",
          "JobTitle": "Senior Engineer"
        }
      ]
    }
  ]
}
```

---

### Frequent Contacts

**Endpoint:** `POST https://substrate.office.com/search/api/v1/suggestions?scenario=peoplecache`

**Auth:** Bearer (Substrate)

Same as people search, but with empty `QueryString`. Returns ranked list of frequently contacted people.

**Request:**
```json
{
  "EntityRequests": [
    {
      "Query": { "QueryString": "", "DisplayQueryString": "" },
      "EntityType": "People",
      "Size": 500,
      "Fields": ["Id", "MRI", "DisplayName", "EmailAddresses", "GivenName", "Surname", "CompanyName", "JobTitle"]
    }
  ]
}
```

---

### Channel Search (Organisation-wide)

**Endpoint:** `POST https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar&setflight=TurnOffMPLSuppressionTeams,EnableTeamsChannelDomainPowerbar&domain=TeamsChannel`

**Auth:** Bearer (Substrate)

**Request:**
```json
{
  "EntityRequests": [
    {
      "Query": {
        "QueryString": "testing",
        "DisplayQueryString": "testing"
      },
      "EntityType": "TeamsChannel",
      "Size": 10
    }
  ],
  "cvid": "uuid",
  "logicalId": "uuid"
}
```

**Response:**
```json
{
  "Groups": [
    {
      "Suggestions": [
        {
          "Name": "AI In Testing",
          "ThreadId": "19:ca554e7ce33b4a2f8099765fba3079bf@thread.tacv2",
          "TeamName": "AI Team",
          "GroupId": "df865310-bf69-4f1b-8dc7-ebd0cbfa090f",
          "EntityType": "ChannelSuggestion",
          "ChannelType": "Standard",
          "Text": "testing",
          "PropertyHits": ["Name"]
        }
      ]
    }
  ]
}
```

**Key Fields:**
- `Name`: Channel display name
- `ThreadId`: Conversation ID for use with messaging APIs
- `TeamName`: Parent team's display name
- `GroupId`: Team's Azure AD group ID
- `ChannelType`: `"Standard"`, `"Private"`, or `"Shared"`

---

## People APIs

### Person Profile (Delve)

**Endpoint:** `POST https://nam.loki.delve.office.com/api/v2/person?smtp={email}&personaType=User&locale=en-gb`

**Auth:** Bearer

**Response:**
```json
{
  "person": {
    "names": [
      {
        "value": {
          "displayName": "Smith, John",
          "givenName": "John",
          "surname": "Smith"
        },
        "source": "Organisation"
      }
    ],
    "emailAddresses": [
      {
        "value": {
          "name": "user@company.com",
          "address": "user@company.com"
        }
      }
    ]
  }
}
```

---

### Batch User Lookup

**Endpoint:** `POST https://teams.microsoft.com/api/mt/part/{region}/beta/users/fetch`

**Auth:** Bearer

**Request Body:**
```json
["8:orgid:{userId1}", "8:orgid:{userId2}"]
```

**Response:**
```json
{
  "value": [
    {
      "alias": "USERNAME",
      "mail": "user@domain.com",
      "displayName": "Display Name",
      "objectType": "User"
    }
  ]
}
```

---

### Batch Profile Resolution (fetchShortProfile)

**Endpoint:** `POST https://teams.microsoft.com/api/mt/part/{region}/beta/users/fetchShortProfile?isMailAddress=false&enableGuest=true&skypeTeamsInfo=true&canBeSmtpAddress=false&includeIBBarredUsers=true&includeDisabledAccounts=true`

**Auth:** Skype Token + Bearer

This is the primary API Teams uses to resolve MRIs to display names. It supports batching multiple MRIs in a single request.

**Request Body:**
```json
[
  "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
  "8:orgid:0166b018-6b8a-4352-9c64-004738067307",
  "28:fd931076-bbfb-4a38-a85c-1f0fb5b61bee"
]
```

**Response:**
```json
{
  "type": "Microsoft.SkypeSpaces.MiddleTier.Models.IUserIdentity",
  "value": [
    {
      "userPrincipalName": "john.smith@company.com",
      "givenName": "John",
      "surname": "Smith",
      "displayName": "Smith, John",
      "jobTitle": "Senior Engineer",
      "department": "Engineering",
      "email": "john.smith@company.com",
      "userType": "Member",
      "isShortProfile": true,
      "tenantName": "Company Name",
      "companyName": "Company Inc",
      "mri": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24"
    }
  ]
}
```

**Supported MRI Formats:**

| Format | Example | Description |
|--------|---------|-------------|
| Organisation user | `8:orgid:{guid}` | Standard Teams/AAD user |
| Bot/App | `28:{guid}` | Teams bot or application |

**Use Cases:**
- Resolving sender MRIs from activity feed to display names
- Bulk resolving user identities for message display
- Getting profile details (job title, department) for users

**Notes:**
- Disabled/deleted accounts return minimal data with `accountEnabled: false`
- Guests may have limited profile information
- The response `mri` field confirms which MRI each result corresponds to

---

### Profile Picture

**Endpoint:** `GET https://teams.microsoft.com/api/mt/part/{region}/beta/users/{userId}/profilepicturev2/{mri}?size=HR96x96`

Available sizes: `HR64x64`, `HR96x96`, `HR196x196`

---

## Virtual Conversations

Teams uses special "virtual conversation" IDs that act as aggregated views across all conversations. These follow the same messaging API pattern but return consolidated data.

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{virtualId}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=200&startTime=1`

**Auth:** Skype Token + Bearer

### Available Virtual Conversations

| Virtual ID | Purpose | Content Quality | Notes |
|------------|---------|-----------------|-------|
| `48:mentions` | @mentions | **Full content (RichText/Html)** | Messages where you were @mentioned - includes complete message text |
| `48:annotations` | Messages with reactions | **Full content (RichText/Html)** | Messages that received reactions - includes the reacted-to message content |
| `48:notifications` | Activity feed (aggregate) | Stubs with references | All notifications combined - often has empty content, use `clumpId` to fetch source |
| `48:saved` | Saved/bookmarked messages | Stubs with references | Messages you've bookmarked - use `secondaryReferenceId` to get original |
| `48:threads` | Followed threads | Stubs with references | Threads you're following for updates |
| `48:notes` | Personal notes | Full content | Self-chat / notes to self |
| `48:calllogs` | Call history | Unknown | Record of calls made/received (endpoint exists, message format unconfirmed) |
| `48:drafts` | Draft messages | Full content | Unsent scheduled messages (uses different endpoint pattern) |

**Content Quality Explanation:**

- **Full content**: The `content` field contains the actual message HTML. Ready to use directly.
- **Stubs with references**: The `content` field is often empty. Use `clumpId` (source conversation ID) and `secondaryReferenceId` to fetch the actual message from the source conversation.

**Recommendation for Activity Data:**

For the richest activity data, prefer `48:mentions` and `48:annotations` over `48:notifications`. The mentions and annotations endpoints return complete message content, while notifications returns stub records that require additional API calls to resolve.

### Response Structure

Virtual conversation messages include additional fields to identify the source:

```json
{
  "messages": [
    {
      "sequenceId": 55,
      "conversationid": "48:saved",
      "conversationLink": "https://teams.microsoft.com/api/chatsvc/amer/v1/users/ME/conversations/48:saved",
      "contenttype": "text",
      "type": "Message",
      "s2spartnername": "skypespaces",
      "clumpId": "19:QsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1@thread.tacv2",
      "secondaryReferenceId": "T_19:QsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1@thread.tacv2_M_1769464929223",
      "id": "1769470012345",
      "originalarrivaltime": "2026-01-26T18:30:00.000Z",
      "content": "Message content here...",
      "from": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
      "imdisplayname": "Smith, John"
    }
  ]
}
```

**Key Fields:**

| Field | Description |
|-------|-------------|
| `clumpId` | The original conversation ID where the message lives |
| `secondaryReferenceId` | Composite key: `T_{conversationId}_M_{messageId}` for messages, `T_{conversationId}_P_{postId}_Threads` for followed threads |
| `id` | Message ID within the virtual conversation (not the original message ID) |
| `originalarrivaltime` | Original timestamp from source conversation |

### Drafts Endpoint (Different Pattern)

Drafts use a slightly different endpoint:

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/drafts?view=msnp24Equivalent&pageSize=200&startTime=1`

**Response:**
```json
{
  "drafts": [
    {
      "sequenceId": 1,
      "conversationid": "48:drafts",
      "draftType": "ScheduledDraft",
      "innerThreadId": "19:abc_def@unq.gbl.spaces",
      "draftDetails": {
        "sendAt": "1755475200000"
      },
      "content": "Scheduled message content..."
    }
  ]
}
```

---

## Messaging APIs

### Send Message

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages`

**Auth:** Skype Token + Bearer

**Request:**
```json
{
  "id": "-1",
  "type": "Message",
  "conversationid": "{conversationId}",
  "conversationLink": "blah/{conversationId}",
  "from": "8:orgid:{userId}",
  "fromUserId": "8:orgid:{userId}",
  "composetime": "2026-01-23T18:03:50.335Z",
  "originalarrivaltime": "2026-01-23T18:03:50.335Z",
  "content": "<p>Message content here</p>",
  "messagetype": "RichText/Html",
  "contenttype": "Text",
  "imdisplayname": "Display Name",
  "clientmessageid": "{uniqueId}"
}
```

**Response:**
```json
{
  "OriginalArrivalTime": 1769191432285
}
```

**Special Conversation IDs:** See [Virtual Conversations](#virtual-conversations) for the full list (`48:notes`, `48:saved`, `48:threads`, etc.)

---

### Get Messages

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages?pageSize=50&view=msnp24Equivalent`

**Auth:** Skype Token + Bearer

**Response:**
```json
{
  "messages": [
    {
      "id": "1769189921704",
      "originalarrivaltime": "2026-01-23T17:54:43.263Z",
      "messagetype": "RichText/Html",
      "contenttype": "text",
      "content": "<p>Message content</p>",
      "from": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
      "imdisplayname": "Smith, John"
    }
  ]
}
```

---

### Reply to Thread (Channel)

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{channelId};messageid={threadRootId}/messages`

**Auth:** Skype Token + Bearer

The `;messageid=` suffix indicates a thread reply. URL-encoded as `%3Bmessageid%3D`.

**URL Pattern Differences:**

| Action | URL Path |
|--------|----------|
| New channel post | `conversations/{channelId}/messages` |
| Reply to thread | `conversations/{channelId};messageid={threadRootId}/messages` |
| Chat message | `conversations/{chatId}/messages` |

**Note:** Chats (1:1, group, meeting) don't use threading. All messages go to the flat conversation.

---

### Edit Message

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}`

**Auth:** Skype Token + Bearer

**Request:**
```json
{
  "id": "{messageId}",
  "type": "Message",
  "conversationid": "{conversationId}",
  "content": "<p>Updated message content</p>",
  "messagetype": "RichText/Html",
  "contenttype": "text",
  "imdisplayname": "Display Name"
}
```

**Response:** `200 OK` (empty or minimal body)

You can only edit your own messages. Returns `403 Forbidden` for others' messages.

---

### Delete Message

**Endpoint:** `DELETE https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}?behavior=softDelete`

**Auth:** Skype Token + Bearer

**Response:** `200 OK` with `null` body

This is a soft delete. Channel owners/moderators can delete others' messages.

---

### Typing Indicator

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages`

**Request:**
```json
{
  "content": "",
  "contenttype": "Application/Message",
  "messagetype": "Control/Typing"
}
```

---

## Conversation APIs

### Create Group Chat

**Endpoint:** `POST https://teams.microsoft.com/api/chatsvc/{region}/v1/threads`

**Auth:** Skype Token + Bearer

**Request:**
```json
{
  "members": [
    { "id": "8:orgid:user1-guid", "role": "Admin" },
    { "id": "8:orgid:user2-guid", "role": "Admin" },
    { "id": "8:orgid:user3-guid", "role": "Admin" }
  ],
  "properties": {
    "threadType": "chat",
    "topic": "Optional chat name"
  }
}
```

**Response (201):**
```json
{
  "threadResource": {
    "id": "19:5bf2c81dc44b4a60a181bf9170953912@thread.v2",
    "tenantId": "56b731a8-...",
    "type": "Thread",
    "properties": {
      "creator": "8:orgid:user1-guid",
      "threadType": "chat",
      "historydisclosed": "false"
    }
  }
}
```

**Notes:**
- All members get `"role": "Admin"` for group chats
- The `topic` property is optional - sets the chat name
- **Response body may be empty `{}`** - extract conversation ID from `Location` header instead
- Location header format: `https://amer.ng.msg.teams.microsoft.com/v1/threads/19:xxx@thread.v2`
- Use the extracted ID with the messages endpoint to send messages

---

### Get User's Teams & Channels

**Endpoint:** `GET https://teams.microsoft.com/api/csa/{region}/api/v3/teams/users/me?isPrefetch=false&enableMembershipSummary=true`

**Auth:** Skype Token + Bearer (CSA)

**Response:**
```json
{
  "conversationFolders": {
    "folderHierarchyVersion": 1769200822270,
    "conversationFolders": [
      {
        "id": "folder-guid",
        "sortType": "UserDefinedCustomOrder",
        "name": "Folder Name",
        "folderType": "UserCreated",
        "conversationFolderItems": [
          {
            "conversationId": "19:channelId@thread.tacv2",
            "createdTime": 1753172521981
          }
        ]
      }
    ]
  },
  "teams": [
    {
      "threadId": "19:teamId@thread.tacv2",
      "displayName": "Team Name",
      "description": "Team description",
      "isFavorite": false,
      "channels": [
        {
          "id": "19:channelId@thread.tacv2",
          "displayName": "General",
          "description": "Channel description",
          "isFavorite": true,
          "membershipType": "standard"
        }
      ]
    }
  ]
}
```

---

### Get Conversation Details

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}?view=msnp24Equivalent`

**Auth:** Skype Token + Bearer

**Response:**
```json
{
  "id": "19:abc@thread.tacv2",
  "threadProperties": {
    "threadType": "topic",
    "productThreadType": "TeamsStandardChannel",
    "groupId": "guid",
    "topic": "Channel topic",
    "topicThreadTopic": "Channel Name",
    "spaceThreadTopic": "Team Name",
    "spaceId": "19:teamroot@thread.tacv2"
  },
  "members": [...],
  "lastMessage": {...}
}
```

**Conversation Type Identification:**

| Type | `threadType` | `productThreadType` | Name Source |
|------|--------------|---------------------|-------------|
| Standard Channel | `topic` | `TeamsStandardChannel` | `topicThreadTopic` |
| Team (General/Root) | `space` | `TeamsTeam` | `spaceThreadTopic` |
| Private Channel | `space` | `TeamsPrivateChannel` | `topicThreadTopic` |
| Meeting Chat | `meeting` | `Meeting` | `topic` |
| Group Chat | `chat` | `Chat` | `topic` or members |
| 1:1 Chat | `chat` | `OneOnOne` | Other participant |

---

### Channel Posts

**Endpoint:** `GET https://teams.microsoft.com/api/csa/{region}/api/v1/containers/{containerId}/posts`

**Query Parameters:**
- `threadedPostsOnly=true` - Only top-level posts
- `pageSize=20`
- `teamId={teamId}`
- `includeRcMetadata=true` - Include read/saved metadata

**Response:**
```json
{
  "posts": [
    {
      "containerId": "19:channelId@thread.tacv2",
      "id": "1769189921704",
      "latestMessageTime": "2026-01-23T17:54:43.263Z",
      "message": {
        "messageType": "RichText/Html",
        "content": "<p>Message content</p>",
        "fromFamilyNameInToken": "Smith",
        "fromGivenNameInToken": "John"
      }
    }
  ]
}
```

---

### Favourites (Get/Modify)

**Endpoint:** `POST https://teams.microsoft.com/api/csa/{region}/api/v1/teams/users/me/conversationFolders?supportsAdditionalSystemGeneratedFolders=true`

**Auth:** Skype Token + Bearer (CSA)

**Get all folders:**
```json
{}
```

**Add to Favorites:**
```json
{
  "actions": [
    {
      "action": "AddItem",
      "folderId": "{tenantId}~{userId}~Favorites",
      "itemId": "{conversationId}"
    }
  ]
}
```

**Remove from Favorites:**
```json
{
  "actions": [
    {
      "action": "RemoveItem",
      "folderId": "{tenantId}~{userId}~Favorites",
      "itemId": "{conversationId}"
    }
  ]
}
```

**Response:**
```json
{
  "folderHierarchyVersion": 1769200822270,
  "conversationFolders": [
    {
      "id": "{tenantId}~{userId}~Favorites",
      "name": "Favorites",
      "folderType": "Favorites",
      "conversationFolderItems": [
        {
          "conversationId": "19:abc@thread.tacv2",
          "createdTime": 1750768187119
        }
      ]
    }
  ]
}
```

---

### Save/Unsave Message (Bookmarks)

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/rcmetadata/{rootMessageId}`

**Auth:** Skype Token + Bearer

**Two-ID System:**

The rcmetadata API uses two different message IDs:
- **URL path** (`rootMessageId`): The thread root post ID for channel threaded replies, or the message ID itself for top-level posts
- **Body** (`mid`): The actual message being saved/unsaved

For **1:1 chats, group chats, meetings, and channel top-level posts**: `rootMessageId` = `messageId` (same value)

For **channel threaded replies**: `rootMessageId` = parent post ID ≠ `messageId`

**Save:**
```json
{ "s": 1, "mid": 1769200192761 }
```

**Unsave:**
```json
{ "s": 0, "mid": 1769200192761 }
```

**Response:**
```json
{
  "conversationId": "19:abc@thread.v2",
  "rootMessageId": 1769200192761,
  "rcMetadata": {
    "lu": 1769200800298,
    "s": 1
  }
}
```

---

### Mark as Read

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/properties?name=consumptionhorizon`

**Auth:** Skype Token + Bearer

**Request:**
```json
{ "consumptionhorizon": "{timestamp1};{timestamp2};{messageId}" }
```

---

### Get Read Position

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/threads/{threadId}/consumptionhorizons`

**Auth:** Skype Token + Bearer

**Response:**
```json
{
  "id": "19:meeting_abc@thread.v2",
  "version": "1769191217184",
  "consumptionhorizons": []
}
```

---

### 1:1 Chat ID Format

Conversation IDs for 1:1 chats are **predictable** - no API call needed:

```
19:{userId1}_{userId2}@unq.gbl.spaces
```

- User IDs are Azure AD object IDs (GUIDs)
- IDs are **sorted lexicographically** (both participants get the same ID)
- The conversation is created implicitly when the first message is sent

**Example:**
- Your ID: `ab76f827-27e2-4c67-a765-f1a53145fa24`
- Other: `5817f485-f870-46eb-bbc4-de216babac62`
- Since `'5' < 'a'`: `19:5817f485-..._ab76f827-...@unq.gbl.spaces`

---

## Activity & Notifications

Teams provides multiple virtual conversation endpoints for activity data. Each has different content quality:

### Mentions Feed (Recommended for @mentions)

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/48%3Amentions/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=200&startTime=1`

**Auth:** Skype Token + Bearer

Returns messages where you were @mentioned with **full content**.

**Response:**
```json
{
  "messages": [
    {
      "sequenceId": 221,
      "conversationid": "48:mentions",
      "contenttype": "RichText/Html",
      "messagetype": "RichText/Html",
      "content": "<p>Hey <span itemtype=\"http://schema.skype.com/Mention\">@Rob</span>, can you review this?</p>",
      "from": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
      "imdisplayname": "Smith, John",
      "clumpId": "19:154a997776b7479cbc27fc34b3e3688a@thread.tacv2",
      "secondaryReferenceId": "T_19:154a997776b7479cbc27fc34b3e3688a@thread.tacv2_M_1769794973794",
      "id": "1769794976533",
      "originalarrivaltime": "2026-01-30T12:30:00.000Z"
    }
  ]
}
```

---

### Annotations Feed (Messages with Reactions)

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/48%3Aannotations/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=200&startTime=1`

**Auth:** Skype Token + Bearer

Returns messages that received reactions/annotations with **full content**.

**Response:**
```json
{
  "messages": [
    {
      "sequenceId": 2267,
      "conversationid": "48:annotations",
      "contenttype": "Text",
      "messagetype": "RichText/Html",
      "content": "<p>FYI the deployment is complete</p>",
      "from": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
      "imdisplayname": "Macdonald, Rob",
      "clumpId": "19:meeting_abc123@thread.v2",
      "id": "1769797291002",
      "originalarrivaltime": "2026-01-30T13:15:00.000Z"
    }
  ]
}
```

---

### Aggregate Activity Feed

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/48%3Anotifications/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=50`

**Auth:** Skype Token + Bearer

Returns all activity types combined but with **limited content** (often empty stubs).

**Response:**
```json
{
  "messages": [
    {
      "sequenceId": 4230,
      "conversationid": "48:notifications",
      "contenttype": "text",
      "s2spartnername": "skypespaces",
      "clumpId": "19:QsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1@thread.tacv2",
      "secondaryReferenceId": "T_19:QsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1@thread.tacv2_P_1769200182753_Threads",
      "content": "",
      "from": "8:orgid:ab76f827-27e2-4c67-a765-f1a53145fa24",
      "id": "1769276832046",
      "originalarrivaltime": "2026-01-24T18:47:12.046Z"
    }
  ],
  "syncState": "base64EncodedState..."
}
```

**Content Limitations:**

The `48:notifications` endpoint often returns empty `content` fields. To get actual message content:

1. Extract `clumpId` (the source conversation ID)
2. Extract message ID from `secondaryReferenceId` (format: `T_{convId}_M_{messageId}` or `T_{convId}_P_{postId}_Threads`)
3. Fetch the message from the source conversation using the messages API

**Activity Types (identified via `messagetype` and content patterns):**

| Type | Identification |
|------|----------------|
| @Mention | Content contains `<span itemtype="http://schema.skype.com/Mention">` |
| Reaction | `messagetype` contains reaction identifier |
| Reply | Standard message in a thread context |

**Recommendation:** For rich activity data, prefer `48:mentions` and `48:annotations` which include full message content. Use `48:notifications` when you need a combined view and are willing to make follow-up API calls for content.

Use `syncState` from response for efficient incremental polling.

---

### Thread Annotations (Reactions, Read Status)

**Endpoint:** `GET https://teams.microsoft.com/api/chatsvc/{region}/v1/threads/{threadId}/annotations?messageIds={id1},{id2}`

**Response:**
```json
{
  "annotations": {
    "1769013008614": {
      "annotations": [
        {
          "mri": "8:orgid:user-guid",
          "time": 1769076365,
          "annotationType": "l2ch",
          "annotationGroup": "userMetaData"
        }
      ]
    }
  }
}
```

---

## Reactions & Emoji

### Add Reaction

**Endpoint:** `PUT https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}/properties?name=emotions`

**Auth:** Skype Token + Bearer

**Request:**
```json
{
  "emotions": {
    "key": "like",
    "value": 1769429691997
  }
}
```

### Remove Reaction

**Endpoint:** `DELETE https://teams.microsoft.com/api/chatsvc/{region}/v1/users/ME/conversations/{conversationId}/messages/{messageId}/properties?name=emotions`

**Request:**
```json
{
  "emotions": {
    "key": "like"
  }
}
```

### Emoji Key Format

**Standard emojis:** Just the name (e.g., `like`, `heart`, `laugh`)

**Custom/org emojis:** `{name};{storage-id}` (e.g., `octi-search;0-wus-d10-66fac2a3b0cda332435c21a14485efe7`)

**Quick Reaction Keys:**

| Emoji | Key |
|-------|-----|
| 👍 | `like` |
| ❤️ | `heart` |
| 😂 | `laugh` |
| 😮 | `surprised` |
| 😢 | `sad` |
| 😠 | `angry` |

**Other Common Keys:**

| Category | Keys |
|----------|------|
| Expressions | `smile`, `wink`, `cry`, `cwl`, `rofl`, `blush`, `speechless`, `wonder`, `sleepy`, `yawn`, `eyeroll`, `worry`, `puke`, `giggle` |
| Affection | `kiss`, `inlove`, `hug`, `lips` |
| Actions | `facepalm`, `sweat`, `dance`, `bow`, `headbang`, `wasntme`, `hungover`, `shivering` |
| Animals | `penguin`, `cat`, `monkey`, `polarbear`, `elephant` |
| Objects | `flower`, `sun`, `star`, `xmastree`, `cake`, `gift`, `cash`, `champagne` |

---

### Custom Emoji Metadata

**Endpoint:** `GET https://teams.microsoft.com/api/csa/{region}/api/v1/customemoji/metadata`

**Response:**
```json
{
  "categories": [
    {
      "id": "customEmoji",
      "title": "Custom Emoji",
      "emoticons": [
        {
          "id": "angrysteam;0-wus-d4-2137e02e9efa1e425eeab4373bbe8827",
          "documentId": "0-wus-d4-2137e02e9efa1e425eeab4373bbe8827",
          "shortcuts": ["angrysteam"],
          "description": "angrysteam"
        }
      ]
    }
  ]
}
```

**Image URL pattern:** `https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/{emoji-id}/default/20_f.png`

---

## Calendar & Scheduling

### Get Meetings (Calendar View)

**Endpoint:** `GET https://teams.microsoft.com/api/mt/part/{region}-{partition}/v2.1/me/calendars/calendarView`

**Auth:** Skype Token + Skype Spaces Bearer Token

This is the primary endpoint for fetching upcoming and recent meetings. Supports OData-style query parameters.

**Region Partitioning:**

The `mt/part` endpoints use partitioned regions (e.g., `amer-02`, `emea-01`). The partition is tenant-specific and assigned when the tenant is provisioned.

**Finding the Correct Partition:**

Teams stores the user's region/partition in localStorage under a key containing `DISCOVER-REGION-GTM`. This contains a `middleTier` URL like:
```
https://teams.microsoft.com/api/mt/part/amer-02
```

Extract the region-partition from this URL to use the correct endpoint without guessing.

**Authentication:**

This endpoint requires a specific token combination:
- `Authentication: skypetoken={skypetoken_asm}` (from cookies)
- `Authorization: Bearer {spacesToken}` (from MSAL localStorage)

The Bearer token must have scope `https://api.spaces.skype.com/Authorization.ReadWrite`. Other tokens (authtoken cookie, CSA token, Substrate token) will return 401 Unauthorized.

**Finding the Spaces Token:**

Look in MSAL localStorage for an accesstoken entry where `target` includes `api.spaces.skype.com`. Extract the `secret` field as the Bearer token.

**Query Parameters:**

| Parameter | Description | Example |
|-----------|-------------|---------|
| `startDate` | Start of date range (ISO 8601) | `2026-02-01T00:00:00.000Z` |
| `endDate` | End of date range (ISO 8601) | `2026-02-11T00:00:00.000Z` |
| `$top` | Max results to return | `100` |
| `$skip` | Pagination offset | `0` |
| `$count` | Include total count | `true` |
| `$orderby` | Sort order | `startTime asc` or `endTime desc` |
| `$filter` | Filter criteria | See below |
| `$select` | Fields to return | See below |

**Common Filters:**

```
isAppointment eq false and isAllDayEvent eq false and isCancelled eq false
```

**Recommended Select Fields:**

```
cleanGlobalObjectId,endTime,eventTimeZone,eventType,hasAttachments,iCalUid,
isAllDayEvent,isAppointment,isCancelled,isOnlineMeeting,isOrganizer,isPrivate,
lastModifiedTime,location,myResponseType,objectId,organizerAddress,organizerName,
schedulingServiceUpdateUrl,showAs,skypeTeamsData,skypeTeamsMeetingUrl,startTime,subject
```

**Example: Upcoming Meetings (next 10 days)**

```
GET /api/mt/part/amer-02/v2.1/me/calendars/calendarView
  ?startDate=2026-02-01T00:00:00.000Z
  &endDate=2026-02-11T00:00:00.000Z
  &$top=100&$count=true&$skip=0
  &$orderby=startTime asc
  &$filter=isAppointment eq false and isAllDayEvent eq false and isCancelled eq false
  &$select=...
```

**Example: Recent Meetings (past 5 days)**

```
GET /api/mt/part/amer-02/v2.1/me/calendars/calendarView
  ?startDate=2026-01-27T23:59:59.999Z
  &endDate=2026-02-01T23:59:59.999Z
  &$top=500&$count=true&$skip=0
  &$orderby=endTime desc
  &$filter=isAppointment eq false and isAllDayEvent eq false and isCancelled eq false
  &$select=...
```

**Response:**

```json
{
  "count": 15,
  "value": [
    {
      "objectId": "AQMkAGU5NDVj...",
      "startTime": "2026-01-30T13:30:00+00:00",
      "endTime": "2026-01-30T14:30:00+00:00",
      "lastModifiedTime": "2026-01-30T13:31:46+00:00",
      "eventTimeZone": "IndianSt",
      "eventType": "Single",
      "subject": "Weekly Standup",
      "location": "Conference Room A",
      "organizerName": "Smith, John",
      "organizerAddress": "john.smith@company.com",
      "isOnlineMeeting": true,
      "skypeTeamsMeetingUrl": "https://teams.microsoft.com/l/meetup-join/...",
      "myResponseType": "Accepted",
      "isOrganizer": false,
      "isCancelled": false,
      "isPrivate": false,
      "showAs": "Busy"
    }
  ]
}
```

**Key Response Fields:**

| Field | Description |
|-------|-------------|
| `subject` | Meeting title |
| `startTime`, `endTime` | Meeting times (ISO 8601 with timezone) |
| `organizerName`, `organizerAddress` | Who created the meeting |
| `skypeTeamsMeetingUrl` | Teams join link (if online meeting) |
| `isOnlineMeeting` | Whether it's a Teams meeting |
| `location` | Location text or room name |
| `myResponseType` | Your RSVP: `None`, `Accepted`, `Tentative`, `Declined` |
| `showAs` | Calendar status: `Free`, `Busy`, `Tentative`, `OutOfOffice` |
| `eventType` | `Single`, `Occurrence`, `Exception`, `SeriesMaster` |
| `skypeTeamsData` | Contains meeting thread ID for fetching chat messages |

**Alternative Endpoint (simpler, less filtering):**

```
GET /api/mt/part/{region}/v2.0/me/calendars/default/calendarView
  ?StartDate=2026-02-01T00:00:00.000Z
  &EndDate=2026-02-07T00:00:00.000Z
  &shouldDecryptData=true
```

---

### Schedule / Availability

**Endpoint:** `POST https://nam.loki.delve.office.com/api/v1/schedule?smtp={email}&personaType=User`

**Response:**
```json
{
  "nextAvailability": {
    "utcDateTime": "0001-01-01T00:00:00",
    "currentStatus": "Free",
    "nextStatus": "Free"
  },
  "workingHoursCalendar": {
    "daysOfWeek": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
    "startTime": "08:00:00",
    "endTime": "17:00:00",
    "timeZone": { "name": "Greenwich Mean Time" }
  }
}
```

---

### Out of Office Status

**Endpoint:** `POST https://nam.loki.delve.office.com/api/v1/oofstatus?smtp={email}&personaType=User`

**Response:**
```json
{
  "outOfOfficeState": "Disabled",
  "externalAudience": "Unknown",
  "emailAddress": "USER@COMPANY.COM"
}
```

---

### Meeting Tabs

**Endpoint:** `GET https://teams.microsoft.com/api/mt/part/{region}/beta/chats/{meetingThreadId}/tabs`

**Response:**
```json
{
  "value": [
    {
      "id": "uuid",
      "name": "Polls",
      "appId": "uuid",
      "configuration": {
        "entityId": "TeamsMeetingPollPage",
        "contentUrl": "https://forms.office.com/..."
      }
    }
  ]
}
```

---

## Transcripts

### Meeting Transcript via WorkingSetFiles

**Endpoint:** `GET https://substrate.office.com/api/beta/me/WorkingSetFiles/`

**Auth:** Bearer (Substrate) — same token as Search

**Critical Header:**
```
Prefer: substrate.flexibleschema,outlook.data-source="Substrate",exchange.behavior="SubstrateFiles"
```

Without this `Prefer` header, the API returns `400 Resource not found for the segment 'WorkingSetFiles'`.

**How It Works:**

The Teams web app stores meeting recordings on SharePoint/OneDrive. The Substrate WorkingSetFiles API indexes these files with meeting metadata, including the **full transcript JSON** embedded in `ItemProperties.Default.TranscriptJson`.

This means the transcript can be retrieved in a **single API call** using the existing Substrate search token — no Graph API permissions or Azure App registration required.

**Query Parameters:**

| Parameter | Description | Example |
|-----------|-------------|--------|
| `$filter` | Filter by meeting thread ID and optional date range | See below |
| `$orderby` | Sort order | `FileCreatedTime desc` |
| `$select` | Fields to return | See below |

**Filter by Meeting Thread ID:**
```
$filter=ItemProperties/Default/MeetingThreadId eq '19:meeting_xxx@thread.v2'
  AND FileCreatedTime gt 2026-02-09T00:00:00.000Z
  AND FileCreatedTime lt 2026-02-11T00:00:00.000Z
```

**Alternative Filter (by ICalUId):**
```
$filter=ItemProperties/Default/MeetingICalUId eq '040000008200E00074C5B7101A82E008...'
```

**Recommended Select Fields:**
```
SharePointItem,Visualization,
ItemProperties/Default/MeetingCallId,
ItemProperties/Default/DriveId,
ItemProperties/Default/RecordingStartDateTime,
ItemProperties/Default/RecordingEndDateTime,
ItemProperties/Default/TranscriptJson,
ItemProperties/Default/DocumentLink
```

**Full Example URL:**
```
GET https://substrate.office.com/api/beta/me/WorkingSetFiles/
  ?$filter=ItemProperties/Default/MeetingThreadId eq '19:meeting_abc@thread.v2'
    AND FileCreatedTime gt 2026-02-09T00:00:00.000Z
    AND FileCreatedTime lt 2026-02-11T00:00:00.000Z
  &$orderby=FileCreatedTime desc
  &$select=SharePointItem,Visualization,ItemProperties/Default/TranscriptJson,...
```

**Response:**
```json
{
  "value": [
    {
      "SharePointItem": {
        "SitePath": "https://tenant-my.sharepoint.com/personal/user_company_com",
        "SiteId": "62a4d642-7d66-4def-bc88-68f1be573afb",
        "UniqueId": "534d9639-9296-4662-bb7e-76cd9f957e55",
        "FileUrl": "https://...Recordings/Meeting-Recording.mp4",
        "MediaDuration": 3134
      },
      "Visualization": {
        "Title": "Weekly Standup",
        "Type": "Video",
        "StaticTeaser": "50d6a4c6-...\r\nSmith, John: Hello everyone...."
      },
      "ItemProperties": {
        "Default": {
          "MeetingCallId": "d10867a1-17f2-4024-880e-1876e23f83a6",
          "DriveId": "b!QtakaGZ97028iGjRvlc6-...",
          "RecordingStartDateTime": "2026-02-10T17:00:26Z",
          "RecordingEndDateTime": "2026-02-10T17:52:40Z",
          "TranscriptJson": "{\"$schema\":\"http://stream.office.com/schemas/transcript.json\",...}",
          "DocumentLink": "https://...Recordings/Meeting-Recording.mp4"
        }
      }
    }
  ]
}
```

**TranscriptJson Format:**

The `TranscriptJson` field is a JSON string containing the full transcript:

```json
{
  "$schema": "http://stream.office.com/schemas/transcript.json",
  "version": "1.0.0",
  "type": "Transcript",
  "entries": [
    {
      "id": "45de8890-da20-4a5f-8a9d-5d4d831402c2/8",
      "text": "Hello everyone.",
      "speakerId": "b71f4d0f-ed13-4f3e-abdf-037e146be579@56b731a8-...",
      "speakerDisplayName": "Smith, John",
      "confidence": 0.54141116,
      "startOffset": "00:00:22.2871418",
      "endOffset": "00:00:23.1671418",
      "hasBeenEdited": false,
      "spokenLanguageTag": "en-gb"
    }
  ]
}
```

**Transcript Entry Fields:**

| Field | Description |
|-------|-------------|
| `text` | Spoken text content |
| `speakerDisplayName` | Speaker's display name |
| `speakerId` | Speaker's MRI (objectId@tenantId) |
| `startOffset` | Start time (HH:MM:SS.mmmmmmm) |
| `endOffset` | End time (HH:MM:SS.mmmmmmm) |
| `confidence` | Speech recognition confidence (0-1) |
| `spokenLanguageTag` | Language (e.g., `en-gb`, `en-us`) |
| `hasBeenEdited` | Whether the entry was manually edited |

**Notes:**
- The transcript is the **full content** — not paginated or progressive
- Only meetings with recording/transcription enabled will have entries
- The recording must be stored on SharePoint/OneDrive (standard for Teams)
- `Visualization.StaticTeaser` contains a short preview of the transcript
- The SharePoint `media/transcripts` endpoint is also available for streaming the transcript separately (used by the Teams web player), but the WorkingSetFiles approach is simpler as it returns everything in one call

---

## Files & Attachments

### Files Shared in Conversation

**Endpoint:** `GET https://substrate.office.com/AllFiles/api/users('OID:{userId}@{tenantId}')/AllShared?ThreadId={conversationId}&ItemTypes=File&ItemTypes=Link&PageSize=25`

**Response:**
```json
{
  "Items": [
    {
      "ItemType": "File",
      "FileData": {
        "FileName": "Meeting Recording.mp4",
        "FileExtension": "mp4",
        "WebUrl": "https://sharepoint.com/..."
      },
      "SharedByDisplayName": "Smith, John",
      "SharedBySmtp": "john.smith@company.com"
    },
    {
      "ItemType": "Link",
      "WeblinkData": {
        "WebUrl": "https://jira.company.com/...",
        "Title": "JIRA Ticket"
      }
    }
  ]
}
```

---

## Common Gotchas

1. **Date operators** - Only explicit dates work (`sent:2026-01-20`) or `sent:today`. Named shortcuts like `sent:lastweek` and `sent:thisweek` return 0 results.

2. **`@me` doesn't exist** - `from:me`, `to:me`, and `mentions:me` don't work. Get your email/name first, then search with those values.

3. **Thread replies require `;messageid=`** - The URL suffix is required for channel thread replies. Chats don't have threading.

4. **Token expiry** - MSAL tokens last ~1 hour. They only refresh when an API call requires them, not on page load.

5. **CSA vs chatsvc auth** - CSA needs the CSA Bearer token; chatsvc uses the skypetoken cookie.

6. **Search won't find all thread replies** - It's full-text search. A reply that doesn't contain your search terms won't appear. Use `teams_get_thread` for full context.

7. **User ID formats vary** - APIs return IDs as raw GUIDs, MRIs (`8:orgid:...`), with tenant suffixes (`...@tenantId`), or base64-encoded. Handle all formats.

8. **Message deep links** - Format: `https://teams.microsoft.com/l/message/{threadId}/{messageTimestamp}`. Use the thread ID, not the channel ID, for threaded messages.

9. **Activity feed content is often empty** - The `48:notifications` endpoint returns stub records with empty `content` fields for many activity types. Use `48:mentions` for @mentions and `48:annotations` for reactions to get full message content.

10. **Sender names not always included** - Activity items may only include the MRI (`8:orgid:...`) without `imdisplayname`. Use the `fetchShortProfile` API to batch-resolve MRIs to display names.

---

## Conducting API Research

To discover new endpoints:

1. Run: `npm run research`
2. Log in to Teams when prompted (uses system Chrome/Edge, reuses existing encrypted session)
3. Navigate to features you want to investigate
4. Press Ctrl+C to stop — findings are saved to `research-findings.json`
5. Analyse the JSON file for request/response patterns

The research script monitors network requests matching interesting patterns (search, messaging, calendar, transcript, etc.) and captures both request headers and response bodies.

For a clean session, clear: `rm -rf ~/.teams-mcp-server/` (or `%APPDATA%\teams-mcp-server\` on Windows).
