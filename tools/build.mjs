// Builds the resumes and cover letters of a folder as .docx + .pdf.
//
//   node tools/build.mjs <folder> [--only resume|cover] [--lang en|fr] [--no-pdf]
//
// <folder> is "master" or an application folder. Every file is built in its `lang` (en by default):
//   resume.yaml, resume.fr.yaml              → <Name>_Resume (en) / <Name>_CV (fr)  .docx + .pdf
//   cover_letter.yaml, cover_letter.fr.yaml  → <Name>_Cover_Letter (en) / <Name>_Lettre_de_presentation (fr)
// A cover letter's header (name, contact) comes from the resume in the same language
// (same folder first, then master/).
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { spawnSync } from 'child_process';
import YAML from 'yaml';
import { Packer } from 'docx';
import { buildResume } from '../templates/resume.mjs';
import { buildCoverLetter } from '../templates/cover_letter.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUTPUT = { resume: { en: 'Resume', fr: 'CV' }, cover: { en: 'Cover_Letter', fr: 'Lettre_de_presentation' } };

const args = process.argv.slice(2);
const option = (name) => (args.includes(name) ? args[args.indexOf(name) + 1] : undefined);
const folderArg = args.find((a, i) => !a.startsWith('--') && !['--only', '--lang'].includes(args[i - 1]));
if (!folderArg) {
  console.error('Usage: node tools/build.mjs <folder> [--only resume|cover] [--lang en|fr] [--no-pdf]');
  process.exit(1);
}
const folder = path.resolve(ROOT, folderArg);
const only = option('--only');
const langFilter = option('--lang');

// French typography: non-breaking space before ":" (e.g. "clients : l’un d’eux").
const frenchTypography = (v) =>
  typeof v === 'string' ? v.replace(/ :/g, ' :')
    : Array.isArray(v) ? v.map(frenchTypography)
      : v && typeof v === 'object' ? Object.fromEntries(Object.entries(v).map(([k, x]) => [k, frenchTypography(x)]))
        : v;

function load(file) {
  const data = YAML.parse(fs.readFileSync(file, 'utf8')) ?? {};
  data.lang ??= 'en';
  if (!OUTPUT.resume[data.lang]) throw new Error(`${path.basename(file)}: lang must be "en" or "fr"`);
  return data.lang === 'fr' ? frenchTypography(data) : data;
}

const files = fs.readdirSync(folder).filter((f) => /^(resume|cover_letter)(\.[a-z]{2})?\.ya?ml$/.test(f));
const resumes = files.filter((f) => f.startsWith('resume')).map((f) => ({ f, data: load(path.join(folder, f)) }));
const covers = files.filter((f) => f.startsWith('cover_letter')).map((f) => ({ f, data: load(path.join(folder, f)) }));

function personFor(lang) {
  const own = resumes.find((r) => r.data.lang === lang) ?? resumes[0];
  if (own) return own.data;
  const master = path.join(ROOT, 'master', lang === 'fr' ? 'resume.fr.yaml' : 'resume.yaml');
  return load(fs.existsSync(master) ? master : path.join(ROOT, 'master', 'resume.yaml'));
}

const wanted = (kind, lang) => (!only || only === kind) && (!langFilter || langFilter === lang);
const jobs = [];
for (const { data } of resumes) {
  if (!wanted('resume', data.lang)) continue;
  const base = data.name.trim().replace(/\s+/g, '_');
  jobs.push({ kind: 'resume', file: `${base}_${OUTPUT.resume[data.lang]}.docx`, ...buildResume(data) });
}
for (const { data } of covers) {
  if (!wanted('cover', data.lang)) continue;
  const person = personFor(data.lang);
  const base = person.name.trim().replace(/\s+/g, '_');
  jobs.push({ kind: 'cover letter', file: `${base}_${OUTPUT.cover[data.lang]}.docx`, ...buildCoverLetter(data, person) });
}
if (!jobs.length) {
  console.error(`Nothing to build in ${folder} (no matching resume*.yaml / cover_letter*.yaml).`);
  process.exit(1);
}
const names = jobs.map((j) => j.file);
const dup = names.find((n, i) => names.indexOf(n) !== i);
if (dup) {
  console.error(`Two YAML files would both produce ${dup} — check their "lang" fields.`);
  process.exit(1);
}

const outputs = [];
for (const job of jobs) {
  const out = path.join(folder, job.file);
  try {
    fs.writeFileSync(out, await Packer.toBuffer(job.doc));
  } catch (e) {
    if (e.code === 'EBUSY' || e.code === 'EPERM') {
      console.error(`Cannot write ${job.file}: it is open in another program (close it in Word and run again).`);
      process.exit(1);
    }
    throw e;
  }
  for (const w of job.warnings) console.warn(`  spacing warning (${job.kind}): ${w}`);
  outputs.push({ ...job, out });
  console.log(`DOCX  ${path.relative(ROOT, out)}`);
}

if (args.includes('--no-pdf')) process.exit(0);

const ps = spawnSync('powershell', ['-NoProfile', '-File', path.join(ROOT, 'tools', 'docx2pdf.ps1'), ...outputs.map((o) => o.out)], { encoding: 'utf8' });
if (ps.status !== 0) {
  console.error('PDF conversion failed:\n' + (ps.stderr || ps.stdout));
  process.exit(1);
}
let tooLong = false;
for (const line of ps.stdout.split(/\r?\n/).filter((l) => l.startsWith('OK'))) {
  const [, pagesPart, pdf] = line.split(' | ');
  const pages = Number(pagesPart.replace('pages=', ''));
  const job = outputs.find((o) => path.resolve(o.out.replace(/\.docx$/, '.pdf')) === path.resolve(pdf.trim()));
  console.log(`PDF   ${path.relative(ROOT, pdf.trim())}  (${pages} page${pages > 1 ? 's' : ''})`);
  if (pages > 1) {
    tooLong = true;
    console.warn(`  WARNING: the ${job?.kind ?? 'document'} is ${pages} pages — it must fit on 1 page.`);
  }
}
process.exit(tooLong ? 2 : 0);
