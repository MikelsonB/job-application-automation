---
description: Show the applications dashboard, or update a status (e.g. /status ukg interview Nov 3)
argument-hint: [company or search words] [new status] [details]
---
Request: $ARGUMENTS

- If the request is empty: run `node tools/tracker.mjs` and present the dashboard clearly (counts, response rate, action needed, list of applications).
- Otherwise: it's a status update. Find the application, then run `node tools/status.mjs <search> <status> --note "<details>"` (statuses: drafting, ready, applied, assessment, interview, offer, accepted, declined, rejected, ghosted, withdrawn). Put dates, interviewer names and other details in the Notes section of application.md and contacts in `contacts:`. Commit.
- If it moves to `interview`, offer to prepare me for it.
