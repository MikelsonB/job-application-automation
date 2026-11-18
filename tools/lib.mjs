// Shared helpers for the application tools.
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import YAML from 'yaml';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const APPS = path.join(ROOT, 'applications');

export const STATUSES = ['drafting', 'ready', 'applied', 'assessment', 'interview', 'offer', 'accepted', 'declined', 'rejected', 'ghosted', 'withdrawn'];
export const CLOSED = ['accepted', 'declined', 'rejected', 'ghosted', 'withdrawn'];
export const FOLLOW_UP_DAYS = 10;

export const today = () => new Date().toLocaleDateString('en-CA'); // YYYY-MM-DD, local time
export function addDays(date, days) {
  const d = new Date(`${date}T12:00:00`);
  d.setDate(d.getDate() + days);
  return d.toLocaleDateString('en-CA');
}

// File/folder-safe name: "Acme Software Inc" → "Acme-Software-Inc"
export const slug = (s) =>
  String(s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[<>:"/\\|?*()[\],.'’]+/g, ' ').trim().replace(/\s+/g, '-').replace(/-+/g, '-');

// application.md = YAML front matter + markdown notes
export function readApplication(file) {
  const text = fs.readFileSync(file, 'utf8');
  const m = text.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n?([\s\S]*)$/);
  if (!m) throw new Error(`${file}: missing YAML front matter`);
  return { doc: YAML.parseDocument(m[1]), body: m[2] };
}
export function writeApplication(file, { doc, body }) {
  fs.writeFileSync(file, `---\n${doc.toString({ lineWidth: 0 })}---\n${body}`);
}

// Every applications/<Company>/<date_role>/application.md
export function listApplications() {
  if (!fs.existsSync(APPS)) return [];
  const out = [];
  for (const company of fs.readdirSync(APPS, { withFileTypes: true })) {
    if (!company.isDirectory()) continue;
    for (const app of fs.readdirSync(path.join(APPS, company.name), { withFileTypes: true })) {
      const file = path.join(APPS, company.name, app.name, 'application.md');
      if (app.isDirectory() && fs.existsSync(file)) out.push({ dir: path.dirname(file), file, rel: `${company.name}/${app.name}` });
    }
  }
  return out;
}

// Accepts a folder path, or a search like "acme" / "acme developer" matched against "<Company>/<folder>".
export function findApplication(query) {
  const asPath = path.resolve(ROOT, query);
  if (fs.existsSync(path.join(asPath, 'application.md'))) return asPath;
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  const hits = listApplications().filter((a) => words.every((w) => a.rel.toLowerCase().includes(w)));
  if (hits.length === 1) return hits[0].dir;
  if (!hits.length) throw new Error(`No application matches "${query}".`);
  throw new Error(`"${query}" matches several applications:\n  ${hits.map((h) => h.rel).join('\n  ')}`);
}
