---
description: Update job_listing.xlsx — new Runway + GitHub-list postings (dev roles), fit scores, still open or closed, deadlines, applied status
argument-hint: [check only | days N]
---
Run the "Job listing" workflow from CLAUDE.md (`/updatejoblisting`).

$ARGUMENTS

- Nothing above → full update: fetch new postings from Runway and the GitHub internship list, review + score the new ones, check every posting
  (open / closed / deadline), rebuild the Excel file, short report, commit.
- "check only" → skip the fetch: check postings + rebuild.
- "days N" → fetch the last N days instead of "since the last update" (`fetch --days N`).

Don't ask me anything during the update; decide hide / fit yourself with the rules in CLAUDE.md and list anything
doubtful in the report. If job_listing.xlsx is open (build exit code 3), ask me to close it, then build again.
