# Link types for MCB testing
_Exported on 11/02/2026 at 19:32:46 GMT from Cursor (2.4.31)_

---

**User**

I'm testing links in the team's MCB. Get me a series of links—different things like deep links, thread links, and channel links.

---

**Cursor**

Let me pull some data from Teams to get you a variety of link types.

The channel search needs a re-auth, but favourites worked. Let me grab some messages from a few channels to get thread/message IDs.

Here's a collection of different Teams link types from your live data:

---

**Channel links** (conversation IDs for channels):

| Channel | Conversation ID |
|---------|----------------|
| WeaponX Support | `19:QsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1@thread.tacv2` |
| EEC Leads | `19:154a997776b7479cbc27fc34b3e3688a@thread.tacv2` |
| SQUAD - Engagement and Impact | `19:e2d13501fe304fad93419e58281fc584@thread.tacv2` |
| WeaponX Architecture | `19:bfe5d477291e44269507b3d25d1090ee@thread.tacv2` |
| EEC Fleet Internal | `19:sE0cuDtFB5S2UvlwhhP-Qji8NjI9Q-B1FXgCGImyFH01@thread.tacv2` |

---

**Message deep links** (direct links to specific messages in channels):

1. Philip Kendall in WeaponX Support:
   `https://teams.microsoft.com/l/message/19%3AQsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1%40thread.tacv2/1770837025504`

2. Diana Jerman in EEC Leads:
   `https://teams.microsoft.com/l/message/19%3A154a997776b7479cbc27fc34b3e3688a%40thread.tacv2/1770835522750`

3. Kim Price in SQUAD - Engagement and Impact:
   `https://teams.microsoft.com/l/message/19%3Ae2d13501fe304fad93419e58281fc584%40thread.tacv2/1770836256847`

---

**Thread root IDs** (the root posts that replies hang off):

| Channel | Thread Root ID |
|---------|---------------|
| WeaponX Support | `1770834618730` |
| EEC Leads (new requests thread) | `1770831926446` |
| EEC Leads (holdouts thread) | `1770763256655` |
| SQUAD - E&I (holdout results) | `1770645620803` |
| SQUAD - E&I (peeking thread) | `1770243294251` |

---

**Meeting chat links** (note the `?context=` parameter):

1. AgentX sync meeting chat:
   `https://teams.microsoft.com/l/message/19%3Ameeting_N2ZkYTBhNGEtYzQ0YS00MWI2LWEyMzMtZTdjZGVkMjA0ZTVh%40thread.v2/1769179918988?context=%7B%22contextType%22%3A%22chat%22%7D`

2. Same meeting, different message:
   `https://teams.microsoft.com/l/message/19%3Ameeting_N2ZkYTBhNGEtYzQ0YS00MWI2LWEyMzMtZTdjZGVkMjA0ZTVh%40thread.v2/1769169707059?context=%7B%22contextType%22%3A%22chat%22%7D`

---

**Meeting conversation ID** (meeting thread format):

`19:meeting_N2ZkYTBhNGEtYzQ0YS00MWI2LWEyMzMtZTdjZGVkMjA0ZTVh@thread.v2`

---

That gives you a mix of: raw channel conversation IDs, encoded deep links to specific messages, thread root references, and meeting chat links with context parameters. All from real conversations so they should resolve in the client. Let me know if you need any other link types (e.g. 1:1 chat links, or links with specific `tenantId`/`groupId` context).

---

**User**

All these just go to the channel, not a specific message:

Message deep links (direct links to specific messages in channels):
Philip Kendall in WeaponX Support:
https://teams.microsoft.com/l/message/19%3AQsLXSoyGdLTIChUa-elhfgq_VyIauBGVMBk3-7orc1w1%40thread.tacv2/1770837025504
Diana Jerman in EEC Leads:
https://teams.microsoft.com/l/message/19%3A154a997776b7479cbc27fc34b3e3688a%40thread.tacv2/1770835522750
Kim Price in SQUAD - Engagement and Impact:
https://teams.microsoft.com/l/message/19%3Ae2d13501fe304fad93419e58281fc584%40thread.tacv2/1770836256847


---

**Cursor**

Good catch. Those links are missing the context parameters that Teams needs to scroll to a specific message in a channel. The MCP is generating bare `/l/message/<channelId>/<messageId>` links without `groupId`, `tenantId`, or `parentMessageId` - so Teams just opens the channel.

For channel thread replies, the full format should look something like:

```
https://teams.microsoft.com/l/message/<channelId>/<parentMessageId>?tenantId=<tenantId>&groupId=<groupId>&parentMessageId=<parentMessageId>&messageId=<replyMessageId>&createdTime=<timestamp>
```

And for top-level channel posts:

```
https://teams.microsoft.com/l/message/<channelId>/<messageId>?tenantId=<tenantId>&groupId=<groupId>&createdTime=<timestamp>
```

Notice the meeting chat links *did* work differently - they had `?context={"contextType":"chat"}` appended, which is why those may behave better.

So the issue is in how the Teams MCP constructs `messageLink` for channel messages - it's missing the required query parameters. Worth raising as a bug. Want me to look at the MCP source to see how `messageLink` is being built?

