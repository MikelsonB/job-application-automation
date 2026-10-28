# Job Application Workspace

Open this folder with Claude Code and give it a job posting. Claude already knows the whole process (see `CLAUDE.md`).

## Daily use

| You type | What happens |
|---|---|
| `/apply <job link>` (or just paste a link/text) | Claude reads the posting → **you confirm the role and choose the language (English, French or both)** → analysis + questions about missing skills → **you approve resume changes** → resume DOCX + PDF → **you approve the cover letter** → cover letter DOCX + PDF + ready-to-paste form answers |
| `/quick <job link>` (or `/apply quick <link>`) | **No questions, English only:** Claude picks the graduation date, tailors the resume as much as possible with what's already confirmed (resume + experience bank), builds it, and gives you one report at the end (fit, changes, skills to confirm, files to upload) |
| `/apply` with a file in `inbox/` | Same, from a PDF / Word / text file / screenshot you dropped in `inbox/` |
| "applied" | Marks it applied today, sets a follow-up reminder 10 days later |
| `/status` | Dashboard: all applications, response rate, deadlines, follow-ups due |
| `/status ukg interview Nov 3` | Updates a status and saves the details |
| `/followups` | What needs a follow-up + email drafts |
| `/inbox` | Reads your Gmail for job emails (confirmations, tests, interviews, rejections), suggests status updates, adds interviews to Google Calendar after you confirm. Never sends or deletes emails. `/status` runs it automatically. |
| `/fill` (or "fill my current tab") | Fills the application form open in your Chrome with what Claude knows, asks you the new questions (and remembers your answers for next time), stops before Next/Submit. You finish and submit. Needs the Claude in Chrome extension. |

You always apply yourself, with the PDF (or DOCX) files from the application's folder.

## A typical day (example)

> Company names, dates and answers below are examples.

### 1. Start (1 min)
1. Open **Chrome**: be logged in to Gmail, LinkedIn and the job sites you use (Claude in Chrome extension connected).
2. Open the **`Job_application`** folder on the Desktop → right-click inside → **Open in Terminal**.
3. Type **`claude`** and press Enter. Claude reads `CLAUDE.md` automatically.

### 2. What's new: `/status` (2 min)
`/status` first checks Gmail (`/inbox`), then shows the dashboard.
```
You:     /status
Claude:  Gmail since Oct 26: 2 job emails
           Shopify   "Online assessment"           applied → assessment (due Nov 4)
           NovaTech  "Interview invitation Nov 12" applied → interview
         Update these? Add to Google Calendar?
You:     yes, add both to calendar
Claude:  Applications: 8 | applied: 4 | assessment: 1 | interview: 1 | rejected: 2
         Response rate: 25%
         Action needed:
           • FOLLOW UP: Desjardins (applied 10 days ago, no reply)
           • DEADLINE: Ubisoft posting closes Oct 30 (still "ready", not submitted)
         Do you want to prepare for the NovaTech interview?
```
Then: `/followups` for follow-up email drafts, or "prep me for NovaTech".

### 3. New jobs: `/apply` (5–10 min per job)
```
You:     /apply https://careers.novatech.com/jobs/4521
```
**Checkpoint 1 — confirm the role + language**
```
Claude:  NovaTech – Software Developer Intern – Winter 2027 – Montreal hybrid
         Eligibility: ✅ graduation  ✅ work permit  ✅ language
         Fit: 80/100 — recommended.
         Language: English, French or both? (I recommend English)
You:     english
```
**Checkpoint 2 — resume changes**
```
Claude:  Not on your resume: Kubernetes? pytest?
         Proposed changes: [before → after]
You:     kubernetes yes at Avanade, pytest no. changes ok
Claude:  Resume ready (1 page): Mikelson_Benjamin_Resume.pdf
```
**Checkpoint 3 — cover letter**
```
Claude:  Draft: [letter]
You:     ok
Claude:  Cover letter ready. answers.md ready. Status: ready. Files to upload: [paths]
```
Repeat for each job — questions already answered are not asked again.

### 4. Apply in the browser: `/fill` (2–5 min per job)
1. In Chrome, open the application page of each "ready" job, one per tab. Log in where needed.
2. In Claude:
```
You:     /fill all tabs
Claude:  NovaTech: page 1 filled — stopped before "Next".
         Shopify: filled — stopped before "Submit".
         Ubisoft: waiting for you to log in.
         1 new question: "Expected hourly rate?"
You:     $28/h, save as default
```
3. **You check each tab**, upload the PDFs if needed, click **Next / Submit** yourself. Multi-page form → "continue novatech".

### 5. Record it (10 sec)
```
You:     applied novatech and shopify
Claude:  Both marked applied today. Follow-up reminders on Nov 6.
```

### 6. End of day (optional)
`/status` for a last look, then close the terminal — everything is saved (folders, tracker, answer bank, git history).

### Routine summary

| When | You type | Time |
|---|---|---|
| Start of the day | `/status` (checks Gmail too) | 2 min |
| New jobs | `/apply <link>` for each | 5–10 min each |
| Applying | open tabs → `/fill all tabs` → you check & submit | 2–5 min each |
| After submitting | "applied ..." | 10 sec |
| Every few days | `/followups` | 5 min |
| Interview coming | "prep me for <company>" | 15–30 min |

### One-time setup before the first real day
1. Install the **Claude in Chrome** extension, sign in, then type `/chrome` in Claude to connect it.
2. Open a **new Claude session** in this folder so all settings load.
3. Try it on **one simple job** (Greenhouse or Lever site) and watch each step.

## Where things are

- `applications/<Company>/<date>_<Position>/` — everything for one application (job description, analysis, resume, cover letter, form answers, status history)
- `tracker.csv` — all applications in one table (open with Excel; it is regenerated automatically)
- `master/resume.yaml` (English) and `master/resume.fr.yaml` (French) — your master resumes. To change your resume for **all future** applications (new job, new project), ask Claude to update the master.
- File names: English → `Mikelson_Benjamin_Resume` + `Mikelson_Benjamin_Cover_Letter`; French → `Mikelson_Benjamin_CV` + `Mikelson_Benjamin_Lettre_de_presentation` (each as .pdf and .docx)
- `master/experience_bank.md` — skills/experience you confirmed that don't fit on one page (Claude fills it as you answer its questions)
- `master/answer_bank.md` — your answers to application form questions (address, citizenship, graduation date, availability…). Every new question you answer is saved here, so Claude never asks it twice. You can edit it anytime.

## Good to know

- **Cover letters show your interest, not just your experience.** Each one explains why you want this company and this job, where you are now (last-year Computer Engineering student with several internships), and why you are a great fit, in 3 short paragraphs (~150-250 words) that sound like you. Your resume already lists the details. If Claude doesn't know a personal reason you want the job, it will ask you for one.
- The resume style is locked to your originals (`master/Mikelson_Resume_Original.pdf` and `master/Mikelson_CV_2026_FR.pdf`); the content always has to fit on **one page**.
- PDFs are made with Microsoft Word. If a build fails, close the `.docx` in Word and try again.
- The Raleway font is installed for your Windows user (files in `assets/fonts/`). On a new computer, right-click each `.ttf` → **Install**.
- Everything is saved in a local git history (no GitHub needed). Your files contain your phone number and email: if you ever push to GitHub, use a **private** repository.
- Manual build: `node tools/build.mjs applications/<Company>/<folder>` · dashboard: `node tools/tracker.mjs`
