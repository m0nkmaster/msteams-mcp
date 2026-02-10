# Roadmap

| Priority | Feature | Description | Difficulty | Notes |
|----------|---------|-------------|------------|-------|
| ~~Bug~~ | ~~When authenticating through the browser after a period of time the user goes through SSO but as soon as Teams starts to load you are briefly shown the Auth complete model and the page closes. However, the tokens have not been refreshed. Upon trying again and going through the SSO process again the tokens are refreshed and the app works as expected.~~ | ~~Medium~~ | Fixed in v0.2.3 — added post-refresh token validation; headless SSO now falls back to visible browser if tokens weren't actually refreshed |
| P2 | Meeting attendees | Filter meetings by attendee (not just organiser) | Medium | Need to research attendee list in calendar API response |
| P2 | Find team | Search/discover teams by name | Easy | Teams List API |
| P2 | Get person details | Detailed profile info (working hours, OOO status) | Easy | Delve API |
| P2 | Get shared files | Files shared in a conversation | Medium | AllFiles API |
| P3 | Meeting recordings | Locate recordings/transcripts | Hard | Needs research |
| P3 | Region auto-detection | Detect user's API region (amer/emea/apac) from session instead of defaulting to amer | Easy | Could extract from browser session or make configurable via env var |
| P3 | Verify Skype token refresh | Check whether the messaging token (skypetoken_asm) gets refreshed during auto token refresh, or only the Substrate token | Easy | Add before/after logging to `refreshTokensViaBrowser()` to compare both tokens |
| P3 | Get message by URL | Fetch a specific message from a Teams link with surrounding context | Medium | Parse URL to extract conversation/message IDs, return target message plus N messages before/after |
| Bug | Message links unreliable | Deep links sometimes fail with "can't find" error when Teams opens | Medium | Investigate link format variations - may be threading/context related |