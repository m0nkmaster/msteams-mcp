# Roadmap

| Priority | Feature | Description | Difficulty | Notes |
|----------|---------|-------------|------------|-------|
| ~~Bug~~ | ~~Message links unreliable~~ | ~~Fixed: added tenantId, groupId, parentMessageId, createdTime per MS docs~~ | ~~Medium~~ | ~~Was missing required query params for channel and chat deep links~~ |
| P3 | Meeting attendees | Filter meetings by attendee (not just organiser) | Medium | Need to research attendee list in calendar API response |
| P3 | Find team | Search/discover teams by name | Easy | Teams List API |
| P3 | Get person details | Detailed profile info (working hours, OOO status) | Easy | Delve API |
| P2 | Get shared files | Files shared in a conversation | Medium | AllFiles API |requirement |
| P3 | Get message by URL | Fetch a specific message from a Teams link with surrounding context | Medium | Parse URL to extract conversation/message IDs, return target message plus N messages before/after |