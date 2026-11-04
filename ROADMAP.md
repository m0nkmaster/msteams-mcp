# Roadmap

## Pending

| Priority | Feature | Description | Difficulty | Notes |
|----------|---------|-------------|------------|-------|
| P3 | Find team | Search/discover teams by name | Easy | Teams List API |
| P3 | Get person details | Detailed profile info (working hours, OOO status) | Easy | Delve API |
| P4 | Meeting attendees | Filter meetings by attendee (not just organiser) | Medium | Need to research attendee list in calendar API response |

## Future Consideration: Microsoft Graph API (hybrid)

The Teams web client's MSAL cache includes a `graph.microsoft.com` token with broad delegated permissions - confirmed live via a fresh login: `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `Calendars.ReadWrite`, `Mail.ReadWrite`, `ChannelMessage.Read.All`, `ChatMessage.Send`, `People.Read`, `Tasks.ReadWrite`, `Notes.ReadWrite.All`, and more. There's also an `outlook.office.com` token with `EWS.AccessAsUser.All` and `Mail.Send`. Using Graph is the same token-borrowing model we already rely on - we piggyback on the Teams client's token, not our own app registration - so it preserves the no-app-registration USP.

We deliberately keep messaging, search and Teams-internal concepts on the internal APIs (Substrate, chatsvc, CSA), because Graph is weaker or blocked there:

- **Search** - Graph message search is shallow next to Substrate full-text.
- **Chat reads** - the borrowed token has `ChatMessage.Send` but no `Chat.Read`, so 1:1/group message bodies are not even in scope.
- **Channel reads** - `ChannelMessage.Read.All` is present, but Graph Teams-message reads carry licensing/metered-billing and throttling constraints the internal APIs sidestep.
- **Teams-internals** - favourites, saved messages, followed threads, activity feed, read/consumption horizons and reaction semantics are largely absent from Graph.
- **Monitoring/rate limits** - Graph has richer telemetry (Microsoft could flag third-party use of the Teams client ID) and per-app rate limits that could spill into the user's real Teams experience.

**Hybrid opportunity:** keep the above on the internal APIs, but lean on Graph where it is clearly the right tool or fills a gap the internal APIs cannot:

- **File upload / attachments** - Graph OneDrive/SharePoint is how the Teams client itself uploads chat files. Being trialled in PR #26 (adds `graph.microsoft.com` to `REFRESH_SCOPES` plus a `teams_upload_file` tool). Two things to resolve before adopting: make the Graph refresh scope non-fatal so a failed Graph token cannot abort the whole refresh (matters for restricted / GCC-High / DoD tenants), and confirm recipients are actually granted access to the uploaded file.
- **Presence** - `/me/presence` and `/communications/getPresencesByUserId` would close the current "presence not available over HTTP" gap - a genuinely new capability.
- **Calendar / meetings** - `/me/calendarView` and `/me/events` are cleaner and better documented than the reverse-engineered mt/part calendar API; a candidate to simplify existing code.
- **Broader M365** - Planner tasks, OneNote, and mail beyond search, where no equivalent internal API exists.
