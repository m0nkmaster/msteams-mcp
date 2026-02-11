# Roadmap

| Priority | Feature | Description | Difficulty | Notes |
|----------|---------|-------------|------------|-------|
| Bug | Message links unreliable | Deep links sometimes fail with "can't find" error when Teams opens | Medium | Investigate link format variations - may be threading/context related |
| Done | Meeting transcripts | Retrieve meeting transcripts via `teams_get_transcript` | Hard | Uses Substrate WorkingSetFiles API: threadId → filter by MeetingThreadId → parse embedded TranscriptJson |
| P2 | Meeting attendees | Filter meetings by attendee (not just organiser) | Medium | Need to research attendee list in calendar API response |
| P2 | Find team | Search/discover teams by name | Easy | Teams List API |
| P2 | Get person details | Detailed profile info (working hours, OOO status) | Easy | Delve API |
| P2 | Get shared files | Files shared in a conversation | Medium | AllFiles API |
| P2 | AI meeting summary | Retrieve Copilot/Intelligent Recap summary for meetings | Hard | Endpoint identified: `substrate.office.com/search/api/v1/recommendations` with scenario `MeetingCatchup.MeetApp`. Needs research to capture full summary content and check Copilot license requirement |
| P3 | Meeting recordings | Locate recordings/transcripts | Medium | Recording URL available in WorkingSetFiles response (`SharePointItem.FileUrl`) — could expose alongside transcript |
| P3 | Region auto-detection | Detect user's API region (amer/emea/apac) from session instead of defaulting to amer | Easy | Could extract from browser session or make configurable via env var |
| P3 | Verify Skype token refresh | Check whether the messaging token (skypetoken_asm) gets refreshed during auto token refresh, or only the Substrate token | Easy | Add before/after logging to `refreshTokensViaBrowser()` to compare both tokens |
| P3 | Get message by URL | Fetch a specific message from a Teams link with surrounding context | Medium | Parse URL to extract conversation/message IDs, return target message plus N messages before/after |