# Job Application Workspace

Open this folder with Claude Code and give it a job posting. Claude already knows the whole process (see `CLAUDE.md`).

## Daily use

| You type | What happens |
|---|---|
| `/apply <job link>` (or just paste a link/text) | Claude reads the posting → **you confirm the role and choose the language (English, French or both)** → analysis + questions about missing skills → **you approve resume changes** → resume DOCX + PDF → **you approve the cover letter** → cover letter DOCX + PDF + ready-to-paste form answers |
| `/apply` with a file in `inbox/` | Same, from a PDF / Word / text file / screenshot you dropped in `inbox/` |
| "applied" | Marks it applied today, sets a follow-up reminder 10 days later |
| `/status` | Dashboard: all applications, response rate, deadlines, follow-ups due |
| `/status ukg interview Nov 3` | Updates a status and saves the details |
| `/followups` | What needs a follow-up + email drafts |

You always apply yourself, with the PDF (or DOCX) files from the application's folder.

## Where things are

- `applications/<Company>/<date>_<Position>/` — everything for one application (job description, analysis, resume, cover letter, form answers, status history)
- `tracker.csv` — all applications in one table (open with Excel; it is regenerated automatically)
- `master/resume.yaml` (English) and `master/resume.fr.yaml` (French) — your master resumes. To change your resume for **all future** applications (new job, new project), ask Claude to update the master.
- File names: English → `Mikelson_Benjamin_Resume` + `Mikelson_Benjamin_Cover_Letter`; French → `Mikelson_Benjamin_CV` + `Mikelson_Benjamin_Lettre_de_presentation` (each as .pdf and .docx)
- `master/experience_bank.md` — skills/experience you confirmed that don't fit on one page (Claude fills it as you answer its questions)
- `master/profile.md` — facts for forms and letters (work permit, availability, languages, links…)

## Good to know

- The resume style is locked to your originals (`master/Mikelson_Resume_Original.pdf` and `master/Mikelson_CV_2026_FR.pdf`); the content always has to fit on **one page**.
- PDFs are made with Microsoft Word. If a build fails, close the `.docx` in Word and try again.
- The Raleway font is installed for your Windows user (files in `assets/fonts/`). On a new computer, right-click each `.ttf` → **Install**.
- Everything is saved in a local git history (no GitHub needed). Your files contain your phone number and email: if you ever push to GitHub, use a **private** repository.
- Manual build: `node tools/build.mjs applications/<Company>/<folder>` · dashboard: `node tools/tracker.mjs`
