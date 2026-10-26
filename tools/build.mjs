// Builds the resume and/or cover letter of a folder as .docx + .pdf.
//
//   node tools/build.mjs <folder> [--only resume|cover] [--no-pdf]
//
// <folder> is "master" or an application folder. It reads:
//   resume.yaml        → <Name>_Resume.docx/.pdf
//   cover_letter.yaml  → <Name>_Cover_Letter.docx/.pdf  (or _Lettre_de_presentation when lang: fr)
// The cover letter header uses resume.yaml from the same folder (or master/).
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { spawnSync } from 'child_process';
import YAML from 'yaml';
import { Packer } from 'docx';
import { buildResume } from '../templates/resume.mjs';
import { buildCoverLetter } from '../templates/cover_letter.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const flag = (name) => args.includes(name);
const option = (name) => (args.includes(name) ? args[args.indexOf(name) + 1] : undefined);
const folderArg = args.find((a, i) => !a.startsWith('--') && args[i - 1] !== '--only');
if (!folderArg) {
  console.error('Usage: node tools/build.mjs <folder> [--only resume|cover] [--no-pdf]');
  process.exit(1);
}
const folder = path.resolve(ROOT, folderArg);
const only = option('--only');
const readYaml = (file) => YAML.parse(fs.readFileSync(file, 'utf8'));

const resumeFile = path.join(folder, 'resume.yaml');
const coverFile = path.join(folder, 'cover_letter.yaml');
const person = readYaml(fs.existsSync(resumeFile) ? resumeFile : path.join(ROOT, 'master', 'resume.yaml'));
const baseName = person.name.trim().replace(/\s+/g, '_');

const jobs = [];
if (only !== 'cover' && fs.existsSync(resumeFile)) {
  jobs.push({ kind: 'resume', file: `${baseName}_Resume.docx`, ...buildResume(person) });
}
if (only !== 'resume' && fs.existsSync(coverFile)) {
  const letter = readYaml(coverFile);
  const file = letter.lang === 'fr' ? `${baseName}_Lettre_de_presentation.docx` : `${baseName}_Cover_Letter.docx`;
  jobs.push({ kind: 'cover', file, ...buildCoverLetter(letter, person) });
}
if (!jobs.length) {
  console.error(`Nothing to build in ${folder} (no resume.yaml / cover_letter.yaml).`);
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

if (flag('--no-pdf')) process.exit(0);

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
