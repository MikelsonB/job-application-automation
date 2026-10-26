---
description: List deadlines and follow-ups due, and draft follow-up emails
---
Run `node tools/tracker.mjs` and look at the "Action needed" section.

- For each follow-up due: draft a short, polite follow-up email (3-5 sentences) for that application, using its application.md (contacts, date applied, position) and job_description.md. Recruiter's name if known, otherwise "Hiring Team".
- For each close deadline on an application that is still `drafting` or `ready`: remind me what is left to do.
- If nothing is due, say so and show when the next follow-up is.

After I confirm I sent a follow-up, add a history note in application.md and set `follow_up` 7 days later. Commit.
