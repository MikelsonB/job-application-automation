# Job Application Automation

A job and internship application system driven by [Claude Code](https://claude.com/claude-code).
Give it a job posting and it:

- **analyzes the posting** against your resume: fit score, eligibility check (graduation window, work
  authorization, language), the keywords to mirror, and the skills you may be missing,
- **tailors a one-page resume** and a short cover letter, in English and/or French, built as DOCX + PDF
  through Microsoft Word, without inventing anything that isn't in your resume or confirmed by you,
- **tracks every application** (status, history, follow-ups, response rate) in `tracker.csv`,
- **keeps a ranked Excel list** of open development internships and co-ops from Runway and community
  GitHub internship lists, with fit %, deadlines and open / closed status,
- **fills application forms** in Chrome from a reusable answer bank (you always submit yourself),
- **reads job emails** in Gmail and suggests status updates and calendar events.

The workflow lives in [`CLAUDE.md`](CLAUDE.md) and the slash commands in [`.claude/commands/`](.claude/commands).
The scripts in `tools/` and `templates/` are plain Node.js (`docx` + `yaml`).

## Contents

- [How it works](#how-it-works)
- [Setup](#setup)
- [Claude commands](#claude-commands)
- [Script commands](#script-commands)
- [Project structure](#project-structure)
- [Privacy](#privacy)

## How it works

```
job posting ──► /apply ──► analysis.md (fit, requirements, keywords, questions)
                   │
                   ├─► you answer the questions ──► tailored resume.yaml ──► DOCX + PDF (1 page)
                   ├─► cover letter draft ──► you approve ──► DOCX + PDF
                   └─► status "ready" ──► /fill (Chrome) ──► you submit ──► "applied" ──► tracker.csv
```

- Your resume is a YAML file (`master/resume.yaml`). Each application gets its own copy, which Claude edits.
- `master/experience_bank.md` stores skills you confirmed that don't fit on one page, so Claude never asks twice.
- `master/answer_bank.md` stores your answers to form questions, matched by meaning across wordings and languages.
- The resume layout is fixed in `templates/style.mjs`; every build is checked to be exactly one page.

## Setup

**Requirements:** Windows, Node.js 18+, Microsoft Word (for PDF export), Claude Code.
Optional: the Claude in Chrome extension (`/fill`), the Gmail and Google Calendar connectors (`/inbox`).

```
git clone https://github.com/MikelsonB/job-application-automation.git
cd job-application-automation
npm install
```

1. Install the Raleway fonts: right-click each file in `assets/fonts/` → **Install**.
2. Copy `examples/master/` to `master/` and replace the example content with your own resume, answers and
   extra experience.
3. Build your reference resume: `node tools/build.mjs master`
4. Start Claude Code in the folder (`claude`) and type `/apply <job link>`.

## Claude commands

| Command | What happens |
|---|---|
| `/apply <job link>` | Full workflow with 3 checkpoints: confirm the role, eligibility and language → answer questions about missing skills (then the tailored resume is built) → approve the cover letter. |
| `/apply` (file in `inbox/`) | Same, from a PDF / Word / text file / screenshot dropped in `inbox/`. |
| `/quick <job link>` | No questions: best tailored resume from what's already confirmed, with one report at the end. |
| `/fill` | Fills the application form open in Chrome, page by page, and stops before the final Submit. |
| `/fill all tabs` | Same, for every open application tab, one after the other. |
| `/status` | Dashboard: counts, response rate, deadlines, follow-ups due (checks Gmail first if connected). |
| `/status acme interview Nov 3` | Updates one application's status. |
| `/inbox` | Reads job emails in Gmail and suggests status updates / calendar events. Never sends or deletes email. |
| `/followups` | Applications needing a follow-up + email drafts. |
| `/updatejoblisting` | Refreshes `job_listing.xlsx`: new postings, fit %, open / closed, deadlines, applied status. |

In plain words: "applied" marks the current application as applied (follow-up in 10 days), "prep me for Acme"
prepares an interview, "apply to J0123" starts `/apply` on a job from the Excel list.

**Statuses:** `drafting` → `ready` → `applied` → `assessment` → `interview` → `offer` → `accepted` / `declined`,
plus `rejected`, `ghosted`, `withdrawn`.

## Script commands

Claude runs these itself; you can also run them from the project root.

| Command | What it does |
|---|---|
| `node tools/new.mjs --company "X" --position "Y" --lang en\|fr\|both [--url URL]` | Creates `applications/<Company>/<date>_<Position>/` with a copy of the master resume(s). |
| `node tools/build.mjs <folder> [--only resume\|cover] [--lang en\|fr]` | Builds DOCX + PDF through Word. Exit code 2 = more than 1 page. |
| `node tools/build.mjs master` | Builds the reference resume in `master/`. |
| `node tools/status.mjs <folder or search words> <status> [--date YYYY-MM-DD] [--note "..."]` | Updates a status (sets the follow-up date when `applied`) and refreshes `tracker.csv`. |
| `node tools/tracker.mjs` | Rebuilds `tracker.csv` and prints the dashboard. |
| `node tools/snapshot.mjs <url> <folder>` | Saves a public posting page as PDF (postings disappear once closed). |
| `node tools/joblist.mjs fetch [--days N] [--only runway\|github] [--dry]` | Gets new development postings. |
| `node tools/joblist.mjs pending` | New jobs still without a fit score. |
| `node tools/joblist.mjs fit J0012=75 J0013=60` | Sets fit scores. |
| `node tools/joblist.mjs note J0012 "text"` | Adds a note. |
| `node tools/joblist.mjs hide J0012 --reason "..."` / `unhide J0012` | Removes / restores a job. |
| `node tools/joblist.mjs mark J0012 applied\|skipped\|clear` | Application status for a job without a folder. |
| `node tools/joblist.mjs check [--all]` | Still open? Deadline? Graduate-only requirement? |
| `node tools/joblist.mjs build` | Writes `job_listing.xlsx` (exit code 3 = the file is open in Excel). |
| `node tools/joblist.mjs show J0012` | Everything stored for one job. |

Job list sources are set in `joblist/config.json` (Runway's public API and the JSON behind a GitHub internship
list; no login, no personal data sent).

## Project structure

```
CLAUDE.md              workflow and rules Claude follows
.claude/commands/      slash commands (/apply, /quick, /fill, /status, /inbox, /followups, /updatejoblisting)
tools/                 Node.js scripts (new, build, status, tracker, snapshot, joblist) + PowerShell helpers for Word / Excel
templates/             resume and cover letter templates (docx), style.mjs = the locked layout
examples/master/       example resume, answer bank and experience bank (fictional person)
assets/fonts/          Raleway (SIL Open Font License)

created when you use it (git-ignored):
master/                your resume (EN/FR), experience bank, answer bank
applications/          one folder per application: posting, analysis, resume, cover letter, status history
tracker.csv            all applications
job_listing.xlsx       ranked list of open roles
```

Resume YAML in short (see `examples/master/resume.yaml`):

```yaml
lang: en                       # en | fr (section titles, file names, French typography)
name: Alex Example
contact: [alex.example@example.com, "(555) 010-0000", "Montreal, QC"]
sections: [experience, education, projects, skills]
experience:
  - company: Example Corp
    title: Software Developer Intern
    dates: May 2026 - Aug 2026
    bullets:
      - "Built a **CI pipeline** in **Python** ..."   # **text** = bold
```

## Privacy

Everything personal (`master/`, `applications/`, `tracker.csv`, `job_listing.xlsx`, `joblist/joblist.json`,
`inbox_state.json`, PDFs and DOCX files) is listed in `.gitignore`. Keep it that way if your copy is public,
or use a private repository for your own data.

Fonts: Raleway, licensed under the [SIL Open Font License](assets/fonts/OFL.txt).
