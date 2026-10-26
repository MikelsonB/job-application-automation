// Rebuilds tracker.csv (opens in Excel) from every application.md and prints a dashboard.
//
//   node tools/tracker.mjs [--quiet]
import fs from 'fs';
import path from 'path';
import { ROOT, CLOSED, STATUSES, today, addDays, listApplications, readApplication } from './lib.mjs';

const COLUMNS = ['date_created', 'date_applied', 'company', 'position', 'type', 'term', 'location', 'work_mode', 'status',
  'apply_language', 'expected_graduation', 'fit_score', 'priority', 'deadline', 'follow_up', 'salary', 'language', 'source', 'job_id', 'url', 'folder'];

const rows = listApplications().map(({ file, rel }) => {
  const data = readApplication(file).doc.toJS() ?? {};
  return { ...data, folder: `applications/${rel}` };
}).sort((a, b) => String(b.date_applied || b.date_created).localeCompare(String(a.date_applied || a.date_created)));

const cell = (v) => {
  const s = v === undefined || v === null ? '' : String(v);
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};
const csv = [COLUMNS.join(','), ...rows.map((r) => COLUMNS.map((c) => cell(r[c])).join(','))].join('\r\n');
try {
  fs.writeFileSync(path.join(ROOT, 'tracker.csv'), '﻿' + csv + '\r\n'); // BOM so Excel reads accents
} catch (e) {
  if (e.code !== 'EBUSY' && e.code !== 'EPERM') throw e;
  console.warn('tracker.csv is open in Excel — close it and run `node tools/tracker.mjs` again.');
}
if (process.argv.includes('--quiet')) process.exit(0);
if (!rows.length) {
  console.log('No applications yet.');
  process.exit(0);
}

// Dashboard
const now = today();
const soon = addDays(now, 7);
const counts = Object.fromEntries(STATUSES.map((s) => [s, 0]));
for (const r of rows) counts[r.status] = (counts[r.status] ?? 0) + 1;
console.log(`Applications: ${rows.length}  |  ` + Object.entries(counts).filter(([, n]) => n).map(([s, n]) => `${s}: ${n}`).join('  '));
const sent = rows.filter((r) => r.date_applied).length;
const responses = rows.filter((r) => ['assessment', 'interview', 'offer', 'accepted', 'declined'].includes(r.status) ||
  (r.status === 'rejected' && (r.history ?? []).some((h) => ['assessment', 'interview'].includes(h.status)))).length;
if (sent) console.log(`Sent: ${sent}  |  positive responses: ${responses} (${Math.round((100 * responses) / sent)}%)`);

const actions = [];
for (const r of rows) {
  const name = `${r.company} - ${r.position}`;
  if (['drafting', 'ready'].includes(r.status) && r.deadline && r.deadline <= soon) actions.push(`DEADLINE ${r.deadline}  ${name} (${r.status})`);
  if (!CLOSED.includes(r.status) && r.follow_up && r.follow_up <= now) actions.push(`FOLLOW UP (due ${r.follow_up})  ${name} (${r.status})`);
}
if (actions.length) console.log('\nAction needed:\n  ' + actions.join('\n  '));

if (rows.length) {
  console.log('\n' + ['Applied', 'Status', 'Company', 'Position'].join(' | '));
  for (const r of rows) console.log([r.date_applied || '—'.padEnd(10), r.status.padEnd(10), r.company, r.position].join(' | '));
}
