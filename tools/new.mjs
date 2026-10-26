// Creates applications/<Company>/<YYYY-MM-DD>_<Position>/ with application.md,
// a copy of the master resume in the chosen language(s), and company.md (once per company).
//
//   node tools/new.mjs --company "UKG" --position "Software Developer Intern" --lang en|fr|both [--url URL] [--date YYYY-MM-DD]
import fs from 'fs';
import path from 'path';
import { ROOT, APPS, slug, today } from './lib.mjs';

const args = process.argv.slice(2);
const opt = (name) => (args.includes(name) ? args[args.indexOf(name) + 1] : undefined);
const company = opt('--company');
const position = opt('--position');
const date = opt('--date') ?? today();
const lang = opt('--lang');
if (!company || !position || !['en', 'fr', 'both'].includes(lang)) {
  console.error('Usage: node tools/new.mjs --company "Company" --position "Position" --lang en|fr|both [--url URL] [--date YYYY-MM-DD]');
  process.exit(1);
}

// Reuse an existing company folder whatever its capitalisation.
fs.mkdirSync(APPS, { recursive: true });
const companySlug = slug(company);
const existing = fs.readdirSync(APPS).find((d) => d.toLowerCase() === companySlug.toLowerCase());
const companyDir = path.join(APPS, existing ?? companySlug);
fs.mkdirSync(companyDir, { recursive: true });

let appDir = path.join(companyDir, `${date}_${slug(position)}`);
for (let n = 2; fs.existsSync(appDir); n++) appDir = path.join(companyDir, `${date}_${slug(position)}-${n}`);
fs.mkdirSync(appDir);

const q = (s) => JSON.stringify(s ?? '');
fs.writeFileSync(path.join(appDir, 'application.md'), `---
company: ${q(company)}
position: ${q(position)}
job_id: ""
url: ${q(opt('--url'))}
source: ""         # LinkedIn | Indeed | Company site | Concordia CareerConnect | Referral | ...
type: ""           # Internship | Co-op | New grad | Full-time | Part-time | Contract
term: ""           # e.g. Summer 2027
duration: ""
location: ""
work_mode: ""      # On-site | Hybrid | Remote
salary: ""
language: ""       # language required by the job: English | French | Bilingual
apply_language: ${q(lang)}   # documents generated in: en | fr | both
posted: ""
deadline: ""
fit_score: ""      # 0-100, from analysis.md
priority: ""       # high | medium | low
status: drafting
date_created: ${q(date)}
date_applied: ""
follow_up: ""
contacts: []       # - { name: "", role: "", email: "", linkedin: "" }
referral: ""
resume: ""
cover_letter: ""
history:
  - date: ${q(date)}
    status: drafting
    note: Folder created
---

## Notes
`);

for (const file of { en: ['resume.yaml'], fr: ['resume.fr.yaml'], both: ['resume.yaml', 'resume.fr.yaml'] }[lang]) {
  fs.copyFileSync(path.join(ROOT, 'master', file), path.join(appDir, file));
}

const companyFile = path.join(companyDir, 'company.md');
if (!fs.existsSync(companyFile)) {
  fs.writeFileSync(companyFile, `# ${company}

- **Website:**
- **Careers page:**
- **Industry / what they do:**
- **Size / locations:**
- **Tech stack (known):**
- **Culture / values:**
- **Recent news:**

## Contacts

## Notes
`);
}

console.log(path.relative(ROOT, appDir).replace(/\\/g, '/'));
