# Roadmap

| Priority | Feature | Description | Difficulty | Notes |
|----------|---------|-------------|------------|-------|
| P3 | Find team | Search/discover teams by name | Easy | Teams List API |
| P3 | Get person details | Detailed profile info (working hours, OOO status) | Easy | Delve API |
| P3 | Get message by URL | Fetch a specific message from a Teams link with surrounding context | Medium | Parse URL to extract conversation/message IDs, return target message plus N messages before/after |
| P4 | Meeting attendees | Filter meetings by attendee (not just organiser) | Medium | Need to research attendee list in calendar API response |

## Future Consideration: Microsoft Graph API

The Teams web client's MSAL cache includes a `graph.microsoft.com` token with broad delegated permissions (Calendars.ReadWrite, Mail.ReadWrite, Files.ReadWrite.All, ChatMessage.Send, People.Read, Tasks.ReadWrite, Notes.ReadWrite.All, etc.). There's also an `outlook.office.com` token with EWS.AccessAsUser.All and Mail.Send.

We currently use Teams' internal APIs (Substrate, chatsvc, CSA) because they don't require an Azure App Registration — the project's key USP. Graph API is Microsoft's official, documented, stable API but:

- **Same auth model** — we'd still piggyback on the Teams web client's token, not our own app registration
- **Stricter monitoring** — Graph has better telemetry; Microsoft could flag third-party use of the Teams client ID
- **Rate limits** — Graph rate limits are per-app; heavy usage could impact the user's actual Teams experience
- **No new capability for current features** — the internal APIs already cover search, messaging, calendar, etc.

**Where Graph could add value:** Expanding beyond Teams into broader M365 features — mail, OneDrive/SharePoint files, Planner tasks, OneNote — where no equivalent internal API exists. If the project evolves toward a general "M365 MCP", Graph would be the natural path. The tokens are already available and would just need adding to `REFRESH_SCOPES` in `token-refresh-http.ts`.