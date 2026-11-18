// Updates the status of an application, logs it in its history and refreshes tracker.csv.
//
//   node tools/status.mjs <folder | search words> <status> [--date YYYY-MM-DD] [--note "text"]
//   e.g. node tools/status.mjs acme applied
//        node tools/status.mjs "acme developer" interview --date 2026-11-03 --note "Tech interview with Jane"
import { spawnSync } from 'child_process';
import path from 'path';
import { ROOT, STATUSES, CLOSED, FOLLOW_UP_DAYS, today, addDays, findApplication, readApplication, writeApplication } from './lib.mjs';

const args = process.argv.slice(2);
const opt = (name) => (args.includes(name) ? args[args.indexOf(name) + 1] : undefined);
const positional = args.filter((a, i) => !a.startsWith('--') && !['--date', '--note'].includes(args[i - 1]));
const status = positional.pop()?.toLowerCase();
const query = positional.join(' ');
if (!query || !STATUSES.includes(status)) {
  console.error(`Usage: node tools/status.mjs <folder | search> <status> [--date YYYY-MM-DD] [--note "text"]\nStatuses: ${STATUSES.join(', ')}`);
  process.exit(1);
}

let dir;
try {
  dir = findApplication(query);
} catch (e) {
  console.error(e.message);
  process.exit(1);
}
const file = path.join(dir, 'application.md');
const app = readApplication(file);
const { doc } = app;
const date = opt('--date') ?? today();

doc.set('status', status);
if (status === 'applied' && !doc.get('date_applied')) doc.set('date_applied', date);
if (status === 'applied' && !doc.get('follow_up')) doc.set('follow_up', addDays(date, FOLLOW_UP_DAYS));
if (CLOSED.includes(status)) doc.set('follow_up', '');
const entry = { date, status };
if (opt('--note')) entry.note = opt('--note');
if (!doc.get('history')) doc.set('history', []);
doc.get('history').add(doc.createNode(entry));
writeApplication(file, app);

console.log(`${path.relative(ROOT, dir).replace(/\\/g, '/')}: ${status} (${date})`);
spawnSync(process.execPath, [path.join(ROOT, 'tools', 'tracker.mjs'), '--quiet'], { stdio: 'inherit' });
