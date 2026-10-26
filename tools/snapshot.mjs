// Saves a job posting page as job_posting_snapshot.pdf (postings disappear once closed).
// Uses Edge/Chrome in headless mode — works for public pages, not for pages behind a login.
//
//   node tools/snapshot.mjs <url> <application folder>
import fs from 'fs';
import os from 'os';
import path from 'path';
import { spawnSync } from 'child_process';
import { ROOT } from './lib.mjs';

const [url, folderArg] = process.argv.slice(2);
if (!url || !folderArg) {
  console.error('Usage: node tools/snapshot.mjs <url> <application folder>');
  process.exit(1);
}
const browser = [
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  'C:/Program Files/Microsoft/Edge/Application/msedge.exe',
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
].find((p) => fs.existsSync(p));
if (!browser) {
  console.error('Edge or Chrome not found.');
  process.exit(1);
}
const out = path.join(path.resolve(ROOT, folderArg), 'job_posting_snapshot.pdf');
const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'jobapp-browser-')); // separate profile: works while your browser is open
const res = spawnSync(browser, [
  '--headless', '--disable-gpu', '--no-first-run', '--no-pdf-header-footer',
  `--user-data-dir=${profile}`, '--virtual-time-budget=15000', `--print-to-pdf=${out}`, url,
], { encoding: 'utf8', timeout: 90000 });
fs.rmSync(profile, { recursive: true, force: true });

if (!fs.existsSync(out) || fs.statSync(out).size < 2000) {
  console.error(`Snapshot failed (${res.error?.message ?? 'empty page'}). Save the page manually: Ctrl+P → Save as PDF.`);
  process.exit(1);
}
console.log(`Saved ${path.relative(ROOT, out)} (${Math.round(fs.statSync(out).size / 1024)} KB)`);
