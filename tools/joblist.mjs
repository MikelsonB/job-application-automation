// Job listing: development roles from Runway (app.joinrunway.io/explore) + community GitHub internship lists → joblist/joblist.json → job_listing.xlsx
// Runway is read through its public API (no browser, no login, no personal data sent).
//
//   node tools/joblist.mjs fetch [--days N] [--only runway|github] [--dry]   new postings from Runway + GitHub lists → adds the development roles
//   node tools/joblist.mjs toptier                     collect the companies Runway tags top-tier (used for US jobs from GitHub lists)
//   node tools/joblist.mjs pending                     new jobs that still need a fit score (or to be hidden)
//   node tools/joblist.mjs fit J0012=75 J0013=60       set fit scores (%)
//   node tools/joblist.mjs note J0012 "text"           note shown in the Excel "notes" column ("" clears it)
//   node tools/joblist.mjs hide J0012 J0013 --reason "not a dev role"     remove from the Excel (kept so it never comes back)
//   node tools/joblist.mjs unhide J0012
//   node tools/joblist.mjs keep J0012                  dismiss a wrong "graduate requirement" flag
//   node tools/joblist.mjs mark J0012 applied|skipped|clear [--date YYYY-MM-DD]   for jobs applied outside /apply
//   node tools/joblist.mjs check [--all] [J0012 ...]   still open? application deadline? (careers site of each job)
//   node tools/joblist.mjs build                       writes job_listing.xlsx (+ joblist/job_listing.csv); Excel file must be closed
//   node tools/joblist.mjs show J0012                  everything stored for one job
//   node tools/joblist.mjs import-csv <file>           one-time import of an older list
import fs from 'fs';
import path from 'path';
import { execFileSync } from 'child_process';
import { ROOT, today, addDays, listApplications, readApplication } from './lib.mjs';

const DIR = path.join(ROOT, 'joblist');
const STORE = path.join(DIR, 'joblist.json');
const CONFIG = path.join(DIR, 'config.json');
const CSV = path.join(DIR, 'job_listing.csv');
const XLSX = path.join(ROOT, 'job_listing.xlsx');
const PS1 = path.join(ROOT, 'tools', 'joblist_xlsx.ps1');

const DEFAULT_CONFIG = {
  sources: [
    { name: 'internships, top-tier companies', employmentType: 'INTERNSHIP', topTierOnly: true, keep: 'canada+usa' },
    { name: 'internships, all companies (Canada only)', employmentType: 'INTERNSHIP', topTierOnly: false, keep: 'canada' },
    { name: 'co-ops, all companies', employmentType: 'CO_OP', topTierOnly: false, keep: 'canada+usa' },
  ],
  // community GitHub lists (JSON behind the README). Only open postings (no 🔒), dev titles, Canada,
  // or US when Runway tags the company top-tier, unless the posting says no sponsorship / US citizenship required.
  github_sources: [
    { name: 'GitHub (Summer2027-Internships)', url: 'https://raw.githubusercontent.com/vanshb03/Summer2027-Internships/dev/.github/scripts/listings.json',
      seasons: ['Summer', 'Fall', 'Winter', 'Spring', 'Spring/Summer'], max_age_days: 180, us_needs_top_tier: true, us_skip_sponsorship: ['Does Not Offer Sponsorship', 'U.S. Citizenship is Required'] },
  ],
  first_fetch_days: 21,
  max_pages_per_source: 400,
  canada_bonus: 8,
  deadline_soon_days: 7,
};

// ---------- store ----------
fs.mkdirSync(DIR, { recursive: true });
if (!fs.existsSync(CONFIG)) fs.writeFileSync(CONFIG, JSON.stringify(DEFAULT_CONFIG, null, 2) + '\n');
const cfg = { ...DEFAULT_CONFIG, ...JSON.parse(fs.readFileSync(CONFIG, 'utf8')) };
const store = fs.existsSync(STORE) ? JSON.parse(fs.readFileSync(STORE, 'utf8')) : { last_fetch: {}, jobs: [] };
const save = () => fs.writeFileSync(STORE, JSON.stringify(store, null, 1) + '\n');
const byKey = (k) => store.jobs.find((j) => j.key.toLowerCase() === String(k).toLowerCase());
const nextKey = () => 'J' + String(store.jobs.reduce((m, j) => Math.max(m, +j.key.slice(1)), 0) + 1).padStart(4, '0');

// ---------- classification ----------
const DEV = /software|develop|d[ée]velopp|logiciel|full[- ]?stack|back[- ]?end|front[- ]?end|\bweb\b|mobile|\bios\b|android|devops|\bsre\b|site reliability|cloud|platform|infrastructure|machine learning|\bml\b|\bai\b|\bia\b|artificial intelligence|intelligence artificielle|data engineer|donn[ée]es|firmware|embedded|embarqu|programm|\bswe\b|\bsde\b|coding|automation|automatisation|\bqa\b|quality assurance|test engineer|software test|test automation|sdet|validation engineer|application(s)? (engineer|developer)|computer (engineer|science)|informatique|security engineer|systems software|game (engineer|programmer)|gameplay|engine programm|game engine|tools engineer|technology intern|tech intern|\bit\b|llm|genai|generative|compiler|kernel|graphics|database/i;
const NOT_DEV = /ph\.?d|doctora|postdoc|master'?s|\bmba\b|marketing|sales|account(ing|ant)|finance|financial|\bart\b|artist|recruit|\bhr\b|human resources|people partner|legal|mechanical|civil eng|chemical|electrical eng|hardware eng|manufacturing|supply chain|business development|communications|content|copywrit|underwrit|audit|research scientist|construction|environmental|process eng|product design|industrial design|editorial|sports programming|television|film|wealth|tax\b|actuar|nurs|pharma|clinical|biolog|chemist|field (service )?eng|traffic|water|rail\b|structural|optical|photonic|\brf\b|analog|thermal|materials|industrial eng|procurement|asset management|debt|transmission|distribution eng|substation|geotech|survey|hvac|facilities/i;
const CA = /canada|ontario|qu[ée]bec|british columbia|alberta|manitoba|nova scotia|saskatchewan|new brunswick|newfoundland|prince edward/i;
const US = /united states|\busa\b|remote|, (alabama|alaska|arizona|arkansas|california|colorado|connecticut|delaware|florida|georgia|hawaii|idaho|illinois|indiana|iowa|kansas|kentucky|louisiana|maine|maryland|massachusetts|michigan|minnesota|mississippi|missouri|montana|nebraska|nevada|new hampshire|new jersey|new mexico|new york|north carolina|north dakota|ohio|oklahoma|oregon|pennsylvania|rhode island|south carolina|south dakota|tennessee|texas|utah|vermont|virginia|washington|west virginia|wisconsin|wyoming|district of columbia)\b/i;
const DEFENSE = /northrop|l3harris|leidos|\bcaci\b|\bkbr\b|lockheed|raytheon|\brtx\b|general dynamics|bae systems|booz allen|\bsaic\b|federal reserve|\bmitre\b|jet propulsion|\bnasa\b|anduril|huntington ingalls|peraton|sandia|johns hopkins (apl|applied physics)|livermore|los alamos|amentum/i;

export function isDev(title) {
  const t = String(title).replace(/bachelor'?s\s*\/\s*master'?s/i, 'bachelors');
  return DEV.test(t) && !NOT_DEV.test(t);
}
const countryOf = (loc) => (CA.test(loc || '') ? 'Canada' : US.test(loc || '') ? 'USA' : /no location|^$/i.test(loc || '') ? 'Unknown' : 'Other');
function typeOf(title, employmentType) {
  if (employmentType === 'CO_OP' || /\bco-?op\b|\bcoop\b/i.test(title)) return 'Co-op (internship)';
  if (/intern|stagiaire|\bstage\b|summer analyst|placement|student/i.test(title)) return 'Internship';
  if (employmentType === 'FULL_TIME' || /new grad|early career|graduate|development program|entry level|university hire/i.test(title)) return 'Full-time (new grad)';
  return 'Internship';
}
function categoryOf(title, company) {
  if (DEFENSE.test(company)) return 'defense (not eligible)';
  if (/machine learning|\bml\b|\bai\b|\bia\b|artificial|intelligence|data|donn[ée]es|llm|genai|generative|deep learning|aiops/i.test(title)) return 'ai / ml / data';
  if (/embedded|embarqu|firmware|systems software|compiler|kernel|\btest|sdet|\bqa\b|quality|validation|performance|asic|\beda\b|hardware|systems? engineer/i.test(title)) return 'embedded / systems / test';
  return 'software / web / mobile';
}

// ---------- URLs ----------
export function cleanUrl(u) {
  try {
    const x = new URL(u);
    const keep = new URLSearchParams();
    for (const [k, v] of x.searchParams) if (/^(gh_jid|job|jobid|jid|id|lang)$/i.test(k) && v) keep.set(k, v);
    x.search = keep.toString();
    x.hash = '';
    return x.toString();
  } catch { return u; }
}
const normUrl = (u) => {
  try {
    const x = new URL(cleanUrl(u));
    return (x.host.replace(/^www\./, '') + x.pathname.replace(/\/+$/, '') + (x.search || '')).toLowerCase();
  } catch { return String(u || '').toLowerCase(); }
};
const longestId = (s) => (String(s || '').match(/\d{5,}/g) || []).sort((a, b) => b.length - a.length)[0] || '';
const normCo = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
const sameCompany = (a, b) => { const c1 = normCo(a), c2 = normCo(b); return !!(c1 && c2 && (c1.includes(c2) || c2.includes(c1) || c1.slice(0, 6) === c2.slice(0, 6))); };
// same posting: same URL, or same company + same posting number (one job listed on two sites, e.g. jobs.sap.com + SmartRecruiters)
const sameJob = (a, b) => normUrl(a.apply_url) === normUrl(b.apply_url) || (sameCompany(a.company, b.company) && longestId(a.apply_url) && longestId(a.apply_url) === longestId(b.apply_url));

// ---------- args ----------
const argv = process.argv.slice(2);
const cmd = argv[0];
const opt = (name) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : undefined; };
const flag = (name) => argv.includes(name);
const positional = argv.slice(1).filter((a, i, all) => !a.startsWith('--') && !(i > 0 && all[i - 1].startsWith('--') && !['--all'].includes(all[i - 1])));

// ---------- Runway ----------
const RUNWAY = 'https://app.joinrunway.io/api/trpc/jobPosting.getJobs?batch=1';
async function runwayPage(src, cursor) {
  const json = {
    filters: { location: [], employmentType: src.employmentType, minMatchScore: 60, maxMatchScore: 100, topTierOnly: !!src.topTierOnly },
    hasLogo: true, limit: 50, sortBy: 'matchScore', direction: 'forward',
  };
  if (cursor) json.cursor = cursor;
  const body = { 0: cursor ? { json, meta: { values: { 'cursor.lastPostedAt': ['Date'] } } } : { json } };
  for (let attempt = 1; ; attempt++) {
    try {
      const r = await fetch(RUNWAY, { method: 'POST', headers: { 'content-type': 'application/json', 'x-trpc-source': 'nextjs-react' }, body: JSON.stringify(body), signal: AbortSignal.timeout(30000) });
      const t = await r.text();
      if (!r.ok) throw new Error(`Runway API ${r.status}: ${t.slice(0, 300)}`);
      return JSON.parse(t)[0].result.data.json;
    } catch (e) {
      if (attempt >= 3) throw e;
      await new Promise((res) => setTimeout(res, 2000 * attempt));
    }
  }
}

async function cmdFetch() {
  const days = opt('--days');
  const only = opt('--only'); // runway | github
  const dry = flag('--dry');  // show what would be added, change nothing
  const added = [];
  let scanned = 0, dev = 0, updated = 0;
  for (const src of only === 'github' ? [] : cfg.sources) {
    const last = store.last_fetch?.[src.name];
    const since = days ? addDays(today(), -Number(days)) : last ? addDays(last, -2) : addDays(today(), -cfg.first_fetch_days);
    let cursor = null, pages = 0, newest = last || '', n = 0;
    while (pages < cfg.max_pages_per_source) {
      const d = await runwayPage(src, cursor);
      pages++;
      for (const r of d.jobs) {
        n++; scanned++;
        if (src.topTierOnly) addTopTier(r.company);
        const posted = new Date(r.postedAt).toLocaleDateString('en-CA');
        if (posted > newest) newest = posted;
        const country = countryOf(r.location);
        if (country === 'Other') continue;
        if (src.keep === 'canada' && country !== 'Canada') continue;
        if (!isDev(r.title)) continue;
        dev++;
        const url = cleanUrl(r.applyUrl);
        let job = store.jobs.find((j) => j.runway_id === r.id) || store.jobs.find((j) => sameJob(j, { company: r.company, apply_url: url }));
        if (job) {
          if (!job.runway_id) job.runway_id = r.id;
          job.last_seen_runway = today();
          updated++;
          continue;
        }
        job = {
          key: nextKey(), runway_id: r.id, company: r.company, title: r.title, location: r.location || '', country,
          work_mode: r.workArrangement || '', type: typeOf(r.title, src.employmentType), category: categoryOf(r.title, r.company),
          date_posted: posted, apply_url: url, runway_url: r.shortenedUrl || '', source: src.name,
          fit: null, notes: '', hidden: false, first_seen: today(), last_seen_runway: today(),
        };
        store.jobs.push(job);
        added.push(job);
      }
      const oldest = d.jobs.length ? new Date(d.jobs.at(-1).postedAt).toLocaleDateString('en-CA') : '';
      if (!d.nextCursor || !d.jobs.length || oldest < since) break;
      cursor = d.nextCursor;
    }
    store.last_fetch[src.name] = newest || today();
    console.log(`${src.name}: ${n} postings scanned (${pages} pages, since ${since})`);
  }
  for (const src of only === 'runway' ? [] : cfg.github_sources || []) {
    const r = await fetchGithub(src, added, dry);
    scanned += r.scanned; dev += r.dev; updated += r.updated;
  }
  if (!dry) save();
  console.log(`\n${dry ? '[dry run, nothing saved] ' : ''}Scanned ${scanned} postings · ${dev} development roles kept · ${added.length} NEW · ${updated} already in the list`);
  if (added.length) { console.log('\nNew jobs (need a fit score or hide):'); printJobs(added); }
}

// ---------- Runway top-tier companies ----------
// Runway has no "top-tier" field per job: a company is top-tier when it shows up with the topTierOnly filter on.
// stricter than sameCompany: same name once legal suffixes are removed, or one name starting with the other (Rivian / Rivian and VW)
const coKey = (s) => String(s || '').toLowerCase().replace(/\b(inc|corp|corporation|co|company|llc|ltd|limited|plc|group|holdings|technologies|technology|the)\b\.?/g, '').replace(/[^a-z0-9]/g, '');
const topTierNames = () => [...new Set([...(store.top_tier_companies || []), ...store.jobs.filter((j) => /top-tier/.test(j.source || '')).map((j) => j.company)])];
function addTopTier(company) {
  if (!company) return false;
  store.top_tier_companies ||= [];
  if (store.top_tier_companies.includes(company)) return false;
  store.top_tier_companies.push(company);
  return true;
}
// one-time / occasional: scan every top-tier page on Runway (internships, co-ops, full-time) to collect the company names
async function cmdTopTier() {
  const before = topTierNames().length;
  for (const employmentType of ['INTERNSHIP', 'CO_OP', 'FULL_TIME']) {
    let cursor = null, pages = 0, n = 0;
    while (pages < cfg.max_pages_per_source) {
      const d = await runwayPage({ employmentType, topTierOnly: true }, cursor);
      pages++;
      for (const r of d.jobs) { n++; addTopTier(r.company); }
      if (!d.nextCursor || !d.jobs.length) break;
      cursor = d.nextCursor;
    }
    console.log(`Runway top-tier ${employmentType}: ${n} postings scanned (${pages} pages)`);
  }
  save();
  const all = topTierNames();
  console.log(`Top-tier companies known: ${all.length} (+${all.length - before})`);
}

// ---------- GitHub internship lists ----------
// one location string → Canada / USA / Other (province or "Canada" first, then a US state code, then Canadian city names)
const ghLocCountry = (l) => (CA.test(l) || /\bCAN\b|, (ON|QC|BC|AB|MB|NS|SK|NB|NL|PE)\b/.test(l) ? 'Canada'
  : US.test(l) || /, [A-Z]{2}\b/.test(l) ? 'USA'
  : /toronto|montr[eé]al|vancouver|ottawa|waterloo|calgary|edmonton|kanata|mississauga|markham|kitchener|halifax|winnipeg|burnaby/i.test(l) ? 'Canada'
  : /united states|\bUSA?\b|remote/i.test(l) ? 'USA' : 'Other');
const ghCountry = (locs) => { const c = locs.map(ghLocCountry); return c.includes('Canada') ? 'Canada' : c.includes('USA') ? 'USA' : 'Other'; };
const normTitle = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
// the list links through a short-link service: follow one redirect to get the company's own careers page
async function resolveShort(u) {
  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const r = await fetch(u, { redirect: 'manual', headers: { 'user-agent': UA }, signal: AbortSignal.timeout(15000) });
      const loc = r.headers.get('location');
      if (loc) return cleanUrl(new URL(loc, u).toString());
    } catch { /* retry once */ }
  }
  return u;
}
async function fetchGithub(src, added, dry) {
  const res = { scanned: 0, dev: 0, updated: 0 };
  const r = await get(src.url);
  const list = r.status === 200 ? json(r.text) : null;
  if (!Array.isArray(list)) { console.log(`${src.name}: could not read the list (${r.status || r.error})`); return res; }
  // US jobs: only companies Runway tags top-tier (store.top_tier_companies, filled by `fetch` and `toptier`)
  const tierKeys = [...new Set(topTierNames().map(coKey).filter((k) => k.length >= 3))];
  if (src.us_needs_top_tier && !tierKeys.length) console.log('  No top-tier companies known yet: run `node tools/joblist.mjs toptier` first (US jobs are all skipped).');
  const isTopTier = (co) => { const a = coKey(co); return a.length >= 3 && tierKeys.some((k) => a === k || (k.length >= 4 && a.startsWith(k))); };
  const minDate = addDays(today(), -(src.max_age_days || 180));
  const skip = { locked: 0, season: 0, old: 0, country: 0, not_dev: 0, us_no_sponsorship: 0, us_defense: 0, us_not_top_tier: 0 };
  const closedNow = [], candidates = [];
  for (const x of list) {
    if (x.is_visible === false) continue;
    res.scanned++;
    const existing = store.jobs.find((j) => j.gh_id === x.id);
    if (existing) {
      res.updated++;
      existing.last_seen_github = today();
      // 🔒 on the list = applications closed
      if (!x.active && existing.posting_status !== 'closed') {
        Object.assign(existing, { posting_status: 'closed', check_detail: 'locked (closed) on the GitHub list', last_checked: today(), deadline: undefined });
        closedNow.push(existing);
      }
      continue;
    }
    if (!x.active) { skip.locked++; continue; }
    if (src.seasons && !src.seasons.includes(x.season)) { skip.season++; continue; }
    const posted = new Date(x.date_posted * 1000).toLocaleDateString('en-CA');
    if (posted < minDate) { skip.old++; continue; }
    const locs = (x.locations || []).map(String);
    const country = ghCountry(locs);
    if (country === 'Other') { skip.country++; continue; }
    const title = String(x.title || '').replace(/[🛂🔒]|🇺🇸/gu, '').trim();
    if (!isDev(title)) { skip.not_dev++; continue; }
    if (country === 'USA') {
      if ((src.us_skip_sponsorship || []).includes(x.sponsorship)) { skip.us_no_sponsorship++; continue; }
      if (DEFENSE.test(x.company_name)) { skip.us_defense++; continue; } // US defense = citizenship / clearance
      if (src.us_needs_top_tier && !isTopTier(x.company_name)) { skip.us_not_top_tier++; continue; }
    }
    res.dev++;
    candidates.push({ x, title, posted, country, locs });
  }
  await pool(candidates, 8, async (c) => { c.url = await resolveShort(c.x.url); });
  for (const { x, title, posted, country, locs, url } of candidates) {
    const job = store.jobs.find((j) => sameJob(j, { company: x.company_name, apply_url: url }))
      || store.jobs.find((j) => sameCompany(j.company, x.company_name) && normTitle(j.title) === normTitle(title));
    if (job) { // same job already found on Runway (or earlier): just link it
      if (!dry) Object.assign(job, { gh_id: x.id, gh_url: x.url, last_seen_github: today() });
      res.updated++;
      continue;
    }
    const nj = {
      key: nextKey(), gh_id: x.id, company: x.company_name, title, location: locs.join('; '), country,
      work_mode: /remote/i.test(locs.join(' ')) ? 'Remote' : '', type: typeOf(title), category: categoryOf(title, x.company_name),
      date_posted: posted, apply_url: url, gh_url: x.url, term: x.season || '', source: src.name,
      fit: null, notes: [x.season && x.season !== 'Summer' ? `${x.season} term (check the year)` : '', country === 'USA' ? (x.sponsorship === 'Offers Sponsorship' ? 'offers sponsorship (GitHub list)' : 'sponsorship not mentioned (check)') : ''].filter(Boolean).join('; '),
      hidden: false, first_seen: today(), last_seen_github: today(),
    };
    if (!dry) store.jobs.push(nj); else nj.key = '(new)';
    added.push(nj);
  }
  if (!dry) store.last_fetch[src.name] = today();
  console.log(`${src.name}: ${res.scanned} postings · kept ${res.dev} · skipped ${Object.entries(skip).map(([k, v]) => `${k.replace(/_/g, ' ')} ${v}`).join(', ')}`);
  if (closedNow.length) { console.log(`Now locked (closed) on ${src.name}: ${closedNow.length}`); printJobs(closedNow); }
  return res;
}

const printJobs = (jobs) => jobs.forEach((j) => console.log(`${j.key} | ${j.type} | ${j.country} | ${j.date_posted} | ${j.company} | ${j.title} | ${j.location}`));

// ---------- posting check (open / closed / deadline) ----------
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36';
async function get(url, accept = 'application/json') {
  for (let attempt = 1; ; attempt++) {
    try {
      const r = await fetch(url, { headers: { 'user-agent': UA, accept }, redirect: 'follow', signal: AbortSignal.timeout(25000) });
      // "too many requests": wait and retry (Microsoft's careers site does this)
      if ((r.status === 429 || r.status === 503) && attempt < 4) { await r.text(); await new Promise((res) => setTimeout(res, 5000 * attempt)); continue; }
      return { status: r.status, url: r.url, text: await r.text() };
    } catch (e) {
      if (attempt < 2) continue;
      return { status: 0, url, text: '', error: e.message };
    }
  }
}
const json = (t) => { try { return JSON.parse(t); } catch { return null; } };
const htmlText = (h) => String(h || '')
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
  .replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

const MON = { jan: 1, january: 1, janvier: 1, feb: 2, february: 2, fevrier: 2, mar: 3, march: 3, mars: 3, apr: 4, april: 4, avril: 4, may: 5, mai: 5,
  jun: 6, june: 6, juin: 6, jul: 7, july: 7, juillet: 7, aug: 8, august: 8, aout: 8, sep: 9, sept: 9, september: 9, septembre: 9,
  oct: 10, october: 10, octobre: 10, nov: 11, november: 11, novembre: 11, dec: 12, december: 12, decembre: 12 };
const MONRE = Object.keys(MON).sort((a, b) => b.length - a.length).join('|');
const iso = (y, m, d) => (m >= 1 && m <= 12 && d >= 1 && d <= 31 ? `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}` : null);
function firstDate(seg) {
  const now = today(), y0 = +now.slice(0, 4);
  const withYear = (y, m, d) => {
    let out = iso(y ? +y : y0, m, d);
    if (out && !y && out < addDays(now, -30)) out = iso(y0 + 1, m, d);
    return out;
  };
  const found = [];
  let m;
  if ((m = seg.match(/\b(20\d\d)-(\d\d)-(\d\d)\b/))) found.push([m.index, iso(+m[1], +m[2], +m[3])]);
  if ((m = seg.match(new RegExp(`\\b(${MONRE})\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?\\b(?:,?\\s+(20\\d\\d))?`, 'i')))) found.push([m.index, withYear(m[3], MON[m[1].toLowerCase()], +m[2])]);
  if ((m = seg.match(new RegExp(`\\b(\\d{1,2})(?:er)?\\s+(${MONRE})\\.?(?:,?\\s+(20\\d\\d))?`, 'i')))) found.push([m.index, withYear(m[3], MON[m[2].toLowerCase()], +m[1])]);
  if ((m = seg.match(/\b(\d{1,2})\/(\d{1,2})\/(20\d\d)\b/))) found.push([m.index, iso(+m[3], +m[1], +m[2])]);
  const ok = found.filter(([, d]) => d && d >= '2026-01-01' && d <= addDays(now, 400)).sort((a, b) => a[0] - b[0]);
  return ok[0]?.[1] || null;
}
export function findDeadline(text) {
  const t = String(text || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ');
  const KEY = /(application deadline|deadline to apply|deadline for (?:applications|submission)|apply by|apply before|applications? (?:will )?(?:be )?(?:accepted )?(?:close|closes|closing|due)|closing date|posting (?:end|close|closing) date|job posting end date|expected (?:posting )?(?:close|closing) date|anticipated (?:posting )?(?:close|closing) date|date limite|date de cloture|postuler avant|deadline:)[^.;]{0,80}/gi;
  const dates = [];
  for (const m of t.matchAll(KEY)) { const d = firstDate(m[0]); if (d) dates.push(d); }
  return dates.sort()[0] || null;
}
// Graduate requirement in the posting: graduate students only (Master's / PhD) or a completed degree required.
// Returns { kind, snippet } for Claude to review (it never hides by itself), or null.
const UNDERGRAD = /bachelor|undergrad|baccalaur|\bb\.?\s?s\.?c?\b|\bbeng\b|\bb\.?a\.?sc|\bb\.?eng|\bbs\s?\/|\bba\b|1st|2nd|3rd|first|second|third|sophomore|junior|senior year|any level/i;
export function findGradRequirement(text) {
  const t = String(text || '').replace(/\s+/g, ' ');
  const around = (i, len) => t.slice(Math.max(0, i - 120), i + len + 120);
  const rules = [
    ['graduate students only', /\b(?:ph\.?\s?d|master'?s|graduate) (?:students?|candidates?) only\b|open (?:only )?to (?:master'?s|ph\.?\s?d|graduate) students|(?:must|need to) be (?:currently )?enrolled in a (?:master'?s|ph\.?\s?d|doctoral|graduate)/i, false],
    ['graduate students only', /(?:currently |actively )?(?:enrolled|pursuing|working towards?|candidate for)[^.;]{0,60}\b(?:master'?s|m\.?\s?sc|ph\.?\s?d|doctoral|doctorate|graduate (?:degree|program|studies|school))\b/i, true],
    ['degree completed', /\b(?:must|should|required to) (?:have|hold) (?:completed|obtained|earned|graduated)[^.;]{0,60}(?:degree|bachelor|diploma)|\bdegree (?:must be )?(?:completed|conferred) (?:by|before|prior to)|\bgraduated (?:within|in) the (?:last|past)|\bnot (?:be )?currently enrolled|\bmust have graduated\b/i, false],
  ];
  for (const [kind, re, needsNoUndergrad] of rules) {
    const m = t.match(re);
    if (!m) continue;
    const win = around(m.index, m[0].length);
    if (needsNoUndergrad && UNDERGRAD.test(win)) continue;
    // a wish, not a requirement: "(preferred)", "Desired qualifications: pursuing MS or PhD"
    const before = t.slice(Math.max(0, m.index - 45), m.index), after = t.slice(m.index + m[0].length, m.index + m[0].length + 100);
    if (/(preferred|desired|nice[- ]to[- ]have|bonus)(?: qualifications?| skills| requirements)?\s*:?\s*$/i.test(before) || /^[^.;]{0,100}?\b(preferred|is a plus|an asset|nice to have|bonus)\b/i.test(after)) continue;
    // "completed degree" rules: skip when students are also accepted, or when it's a form question ("Yes / No, I am…")
    if (kind === 'degree completed' && (/currently enrolled in a (bachelor|undergrad)|students?\b|\byes\b.{0,40}\bno\b/i.test(win))) continue;
    return { kind, snippet: win.trim() };
  }
  return null;
}

const CLOSED_TEXT =/no longer (accepting applications|available|open|active|accepting)|this (job|position|posting|requisition|opportunity) (is )?(no longer|has (been )?(closed|expired|filled))|position has been filled|job (posting )?has expired|posting has expired|n'est plus (disponible|accessible)|offre (d'emploi )?(n'est plus|a expire)/i;

// cached as promises so parallel checks share one lookup
const ghBoards = new Map();
function ghBoardExists(b) {
  if (!ghBoards.has(b)) ghBoards.set(b, get(`https://boards-api.greenhouse.io/v1/boards/${b}`).then((r) => r.status === 200));
  return ghBoards.get(b);
}
const pcsDomains = new Map();
function pcsDomain(host) {
  if (!pcsDomains.has(host)) {
    const dom = host.split('.').slice(-2).join('.');
    // → the API domain, null (not an Eightfold site) or 'busy' (site answered "too many requests")
    pcsDomains.set(host, (async () => {
      const r = await get(`https://${host}/api/pcsx/search?domain=${dom}&query=intern&start=0&num=1&hl=en`);
      if (r.status === 200 && json(r.text)?.data?.positions) return dom;
      if (r.status === 429 || r.status === 503) return 'busy';
      return null;
    })());
  }
  return pcsDomains.get(host);
}

export async function checkPosting(job) {
  const url = job.apply_url;
  let u;
  try { u = new URL(url); } catch { return { status: 'unknown', detail: 'bad url' }; }
  const host = u.host.toLowerCase();
  const done = (status, detail, text, deadline, deadline_source) => {
    const d = deadline || findDeadline(text);
    return { status, detail, deadline: d || null, deadline_source: d ? deadline_source || 'posting text' : null, grad: findGradRequirement(text) };
  };

  // Workday
  let m = url.match(/^https:\/\/([^/]+\.myworkdayjobs\.com)\/(?:[a-z]{2}-[A-Z]{2}\/)?([^/]+)\/job\/(.+?)\/?$/);
  let cxs = m ? `https://${m[1]}/wday/cxs/${m[1].split('.')[0]}/${m[2]}/job/${m[3]}` : null;
  if (!cxs && (m = url.match(/^https:\/\/([^/]+\.myworkdaysite\.com)\/(?:[a-z]{2}-[A-Z]{2}\/)?recruiting\/([^/]+)\/([^/]+)\/job\/(.+?)\/?$/))) cxs = `https://${m[1]}/wday/cxs/${m[2]}/${m[3]}/job/${m[4]}`;
  if (cxs) {
    const r = await get(cxs);
    const info = json(r.text)?.jobPostingInfo;
    if (r.status === 200 && info) return done(info.posted === false || info.canApply === false ? 'closed' : 'open', 'workday', htmlText(info.jobDescription));
    if (r.status === 404) return done('closed', 'workday: not found');
    return done('unknown', `workday ${r.status}`);
  }
  // Oracle Cloud (Candidate Experience)
  if ((m = url.match(/^https:\/\/([^/]+)\/hcmUI\/CandidateExperience\/[a-z]{2}\/sites\/([^/]+)\/job\/(\d+)/))) {
    const r = await get(`https://${m[1]}/hcmRestApi/resources/latest/recruitingCEJobRequisitionDetails?expand=all&onlyData=true&finder=ById;Id=%22${m[3]}%22,siteNumber=${m[2]}`);
    const d = json(r.text);
    if (r.status === 200 && d) {
      const it = d.items?.[0];
      if (!it) return done('closed', 'oracle: not found');
      const end = it.ExternalPostedEndDate ? new Date(it.ExternalPostedEndDate) : null;
      const endDay = end && end.getFullYear() < 2100 ? end.toLocaleDateString('en-CA') : null;
      const text = Object.values(it).filter((v) => typeof v === 'string' && v.length > 40).map(htmlText).join(' ');
      return done(endDay && endDay < today() ? 'closed' : 'open', 'oracle', text, endDay, endDay ? 'site end date' : null);
    }
    return done('unknown', `oracle ${r.status}`);
  }
  // Greenhouse (board URL, or a company page with ?gh_jid=)
  m = url.match(/(?:boards|job-boards)(?:\.eu)?\.greenhouse\.io\/([^/]+)\/jobs\/(\d+)/);
  let gh = m ? { board: m[1], id: m[2] } : null;
  if (!gh && u.searchParams.get('gh_jid')) {
    const id = u.searchParams.get('gh_jid');
    const sld = host.replace(/^www\./, '').split('.').slice(-2, -1)[0];
    for (const b of [...new Set([sld, sld.replace(/hq$|careers$|jobs$/, '')])]) if (await ghBoardExists(b)) { gh = { board: b, id }; break; }
  }
  if (gh) {
    const r = await get(`https://boards-api.greenhouse.io/v1/boards/${gh.board}/jobs/${gh.id}`);
    if (r.status === 200) return done('open', 'greenhouse', htmlText(json(r.text)?.content));
    if (r.status === 404 && (await ghBoardExists(gh.board))) return done('closed', 'greenhouse: not found');
    return done('unknown', `greenhouse ${r.status}`);
  }
  // Lever
  if ((m = url.match(/jobs\.lever\.co\/([^/]+)\/([0-9a-f-]{36})/))) {
    const r = await get(`https://api.lever.co/v0/postings/${m[1]}/${m[2]}`);
    const d = json(r.text);
    if (r.status === 200 && d) return done('open', 'lever', [d.descriptionPlain, d.additionalPlain, ...(d.lists || []).map((l) => htmlText(l.content))].join(' '));
    if (r.status === 404) return done('closed', 'lever: not found');
    return done('unknown', `lever ${r.status}`);
  }
  // Ashby
  if ((m = url.match(/jobs\.ashbyhq\.com\/([^/]+)\/([0-9a-f-]{36})/))) {
    const r = await get(`https://api.ashbyhq.com/posting-api/job-board/${m[1]}`);
    const d = json(r.text);
    if (r.status === 200 && d?.jobs) {
      const jb = d.jobs.find((x) => x.id === m[2]);
      return jb ? done('open', 'ashby', jb.descriptionPlain || htmlText(jb.descriptionHtml)) : done('closed', 'ashby: not on the board');
    }
    return done('unknown', `ashby ${r.status}`);
  }
  // SmartRecruiters
  // jobs.sap.com uses the same posting IDs as SAP's SmartRecruiters account (jobs.sap.com blocks scripts)
  if ((m = url.match(/jobs\.sap\.com\/(?:[a-z]{2}\/)?jobs\/(\d+)/))) m = [null, 'SAPITBusinessSysteme', m[1]];
  else m = url.match(/jobs\.smartrecruiters\.com\/([^/]+)\/(\d+)/);
  if (m) {
    const r = await get(`https://api.smartrecruiters.com/v1/companies/${m[1]}/postings/${m[2]}`);
    const d = json(r.text);
    if (r.status === 200 && d) return done(d.active === false ? 'closed' : 'open', 'smartrecruiters', Object.values(d.jobAd?.sections || {}).map((s) => htmlText(s?.text)).join(' '));
    if (r.status === 404 || r.status === 410) return done('closed', 'smartrecruiters: not found');
    return done('unknown', `smartrecruiters ${r.status}`);
  }
  // Eightfold "PCS" sites (Microsoft, Qualcomm, Ford, Applied Materials…): /careers/job/<id>
  if ((m = url.match(/^https:\/\/([^/]+)\/careers\/job\/(\d{9,})/)) && !/eightfold\.ai$/.test(host)) {
    const dom = await pcsDomain(host);
    if (dom === 'busy') return done('unknown', 'site busy (too many requests), try later');
    if (dom) {
      const r = await get(`https://${host}/api/pcsx/position_details?position_id=${m[2]}&domain=${dom}&hl=en`);
      const d = json(r.text)?.data;
      if (r.status === 200 && d?.id) return done('open', 'eightfold', htmlText(d.jobDescription));
      if (r.status === 404) return done('closed', 'eightfold: position not found');
      return done('unknown', `eightfold ${r.status}`);
    }
  }
  // Eightfold classic (<company>.eightfold.ai/careers/job/<id>)
  if ((m = url.match(/^https:\/\/([^.]+)\.eightfold\.ai\/careers\/job\/(\d+)/))) {
    const r = await get(`https://${m[1]}.eightfold.ai/api/apply/v2/jobs/${m[2]}?domain=${m[1]}.com`);
    const d = json(r.text);
    if (r.status === 200 && d?.id) return done('open', 'eightfold', htmlText(d.job_description));
  }
  // Amazon
  if ((m = url.match(/amazon\.jobs\/[a-z]{2}\/jobs\/(\d+)/))) {
    const r = await get(url, 'text/html');
    if (r.status === 404 || r.status === 410) return done('closed', 'amazon: not found');
    if (r.status === 200 || r.status === 406) return done('open', 'amazon', htmlText(r.text));
    return done('unknown', `amazon ${r.status}`);
  }
  // Any other site: read the page
  const r = await get(url, 'text/html,application/xhtml+xml');
  if (r.status === 404 || r.status === 410) return done('closed', `page ${r.status}`);
  if (r.status !== 200) return done('unknown', r.error ? 'page error' : `page ${r.status}`);
  const text = htmlText(r.text);
  if (CLOSED_TEXT.test(text)) return done('closed', 'page says closed', text);
  const words = String(job.title).toLowerCase().split(/[^a-z0-9]+/).filter((w) => w.length >= 4 && !/intern|summer|2027|2026/.test(w)).slice(0, 4);
  const hits = words.filter((w) => text.toLowerCase().includes(w)).length;
  return done(words.length && hits >= Math.min(2, words.length) ? 'open' : 'unknown', hits ? 'page loads' : 'page needs a browser', text);
}

async function pool(items, n, fn) {
  let i = 0;
  await Promise.all(Array.from({ length: n }, async () => { while (i < items.length) await fn(items[i++]); }));
}

async function cmdCheck() {
  const keys = positional;
  let jobs = keys.length ? keys.map((k) => byKey(k) || die(`No job ${k}`)) : store.jobs.filter((j) => !j.hidden && (flag('--all') || (j.posting_status !== 'closed' && j.last_checked !== today())));
  console.log(`Checking ${jobs.length} postings…`);
  const before = Object.fromEntries(jobs.map((j) => [j.key, j.posting_status]));
  let done = 0;
  await pool(jobs, 8, async (j) => {
    const r = await checkPosting(j);
    j.posting_status = r.status;
    j.check_detail = r.detail;
    if (r.status !== 'unknown') { j.deadline = r.deadline || undefined; j.deadline_source = r.deadline_source || undefined; } // re-read each time (the posting can change)
    if (r.status !== 'unknown') {
      const g = r.grad
        || (/\bgraduate (intern|student|co-?op)\b|\bgrad intern\b/i.test(j.title) ? { kind: 'graduate students only', snippet: 'title: ' + j.title } : null)
        || (j.type === 'Full-time (new grad)' ? { kind: 'degree completed', snippet: 'full-time new-grad role (starts after graduation)' } : null);
      j.grad_flag = g?.kind;
      j.grad_snippet = g?.snippet;
    }
    j.last_checked = today();
    if (++done % 25 === 0) console.log(`  ${done}/${jobs.length}`);
  });
  save();
  const count = (s) => jobs.filter((j) => j.posting_status === s).length;
  console.log(`open: ${count('open')} · closed: ${count('closed')} · unknown: ${count('unknown')} · with a deadline: ${jobs.filter((j) => j.deadline).length}`);
  const newlyClosed = jobs.filter((j) => j.posting_status === 'closed' && before[j.key] !== 'closed');
  if (newlyClosed.length) { console.log('\nNow closed:'); printJobs(newlyClosed); }
  const flagged = store.jobs.filter((j) => !j.hidden && j.grad_flag && !j.grad_ok);
  if (flagged.length) {
    console.log(`\nGraduate requirement found (review, then hide or keep): ${flagged.length}`);
    flagged.forEach((j) => console.log(`${j.key} | ${j.grad_flag} | ${j.company} | ${j.title}\n      "${String(j.grad_snippet).slice(0, 260)}"`));
  }
  const soon = jobs.filter((j) => j.deadline && j.deadline >= today() && j.deadline <= addDays(today(), cfg.deadline_soon_days) && j.posting_status !== 'closed');
  if (soon.length) { console.log(`\nDeadline within ${cfg.deadline_soon_days} days:`); soon.forEach((j) => console.log(`${j.deadline} | ${j.key} | ${j.company} | ${j.title}`)); }
}

// ---------- application status (from applications/ + manual marks) ----------
function loadApplications() {
  return listApplications().map(({ file, rel }) => {
    const d = readApplication(file).doc.toJS() ?? {};
    return { rel, company: d.company, position: d.position, status: d.status, url: d.url, id: longestId(d.job_id) || longestId(d.url) };
  });
}
function applicationOf(job, apps) {
  const nu = normUrl(job.apply_url);
  const hit = apps.find((a) => a.url && normUrl(a.url) === nu) || apps.find((a) => {
    const host = (x) => { try { return new URL(x).host.replace(/^www\./, ''); } catch { return ''; } };
    // same company + same posting number, or same careers site + same posting number (company names can differ: Beenox / Activision)
    return a.id && a.id === longestId(job.apply_url) && (sameCompany(a.company, job.company) || (a.url && host(a.url) === host(job.apply_url)));
  });
  if (hit) return { status: hit.status, folder: `applications/${hit.rel}` };
  if (job.manual_status === 'applied') return { status: 'applied (outside /apply)' };
  if (job.manual_status === 'skipped') return { status: 'skipped' };
  return { status: 'not applied' };
}

// ---------- build Excel ----------
function postingLabel(j) {
  if (j.posting_status === 'closed') return 'closed';
  if (j.deadline && j.deadline < today()) return 'deadline passed';
  return j.posting_status || 'not checked';
}
const sourceLabel = (j) => { const gh = !!j.gh_id, rw = !!j.runway_id || !String(j.source || '').startsWith('GitHub'); return gh && rw ? 'Runway + GitHub' : gh ? 'GitHub' : 'Runway'; };
// Excel is editable for 3 columns: my_notes (free text), fit (0-100) and application
// (applied / skipped / not applied / hide). A cell counts as edited when it differs from what the last build wrote.
const APP_WORDS = {
  applied: ['applied', 'apply', 'yes', 'done', 'sent', 'postulé', 'postule', 'oui'],
  skipped: ['skipped', 'skip', 'no', 'not interested', 'pass', 'non'],
  clear: ['not applied', 'clear', ''],
  hide: ['hide', 'hidden', 'remove', 'delete'],
};
function importExcelEdits() {
  let cells;
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', PS1, '-Mode', 'read', '-Xlsx', XLSX], { encoding: 'utf8' });
    cells = JSON.parse(out.trim() || '{}');
  } catch (e) { console.warn('Could not read your edits from the Excel file: ' + e.message.split('\n')[0]); return; }
  const apps = loadApplications();
  const changes = [], warnings = [];
  for (const j of store.jobs) {
    const c = cells[j.key];
    if (!c) continue;
    j.my_notes = c.n || '';
    // fit
    const before = j.excel ? j.excel.fit : j.fit ?? null;
    const typed = c.f === '' || c.f == null ? null : Number(c.f);
    if (typed !== before && !(typed == null && before == null)) {
      if (typed == null) warnings.push(`${j.key}: fit cell emptied, kept ${j.fit ?? 'empty'}`);
      else if (!(typed >= 0 && typed <= 100)) warnings.push(`${j.key}: fit "${c.f}" ignored (must be 0-100)`);
      else { changes.push(`${j.key} ${j.company}: fit ${j.fit ?? '-'} → ${Math.round(typed)}`); j.fit = Math.round(typed); j.fit_source = 'excel'; }
    }
    // application
    const wrote = String(j.excel ? j.excel.application : applicationOf(j, apps).status).trim().toLowerCase();
    const now = String(c.a ?? '').trim().toLowerCase();
    if (now === wrote) continue;
    const word = Object.keys(APP_WORDS).find((k) => APP_WORDS[k].includes(now));
    const folder = applicationOf({ ...j, manual_status: undefined }, apps).folder;
    if (!word) { warnings.push(`${j.key}: application "${c.a}" not understood (use applied, skipped, not applied or hide)`); continue; }
    if (word === 'hide') { j.hidden = true; j.hidden_reason = 'hidden in Excel'; changes.push(`${j.key} ${j.company}: hidden`); continue; }
    if (folder) {
      // Excel is the source of truth for "applied": a folder still in drafting / ready is moved to applied
      const folderStatus = applicationOf({ ...j, manual_status: undefined }, apps).status;
      if (word === 'applied' && ['drafting', 'ready'].includes(folderStatus)) {
        execFileSync('node', [path.join(ROOT, 'tools', 'status.mjs'), folder, 'applied', '--note', 'Marked applied in job_listing.xlsx'], { cwd: ROOT, stdio: 'ignore' });
        changes.push(`${j.key} ${j.company}: applied (folder ${folder} updated from ${folderStatus})`);
      } else if (word === 'applied') {
        changes.push(`${j.key} ${j.company}: already "${folderStatus}" in its folder, nothing to change`);
      } else {
        warnings.push(`${j.key}: has an application folder (${folder}, status ${folderStatus}); "${c.a}" not applied there, tell Claude to update it with status.mjs`);
      }
      continue;
    }
    j.manual_status = word === 'clear' ? undefined : word;
    j.manual_status_date = word === 'clear' ? undefined : today();
    changes.push(`${j.key} ${j.company}: ${word === 'clear' ? 'not applied' : word}`);
  }
  save();
  if (changes.length) console.log('Your edits in Excel, saved:\n  ' + changes.join('\n  '));
  if (warnings.length) console.log('Excel edits NOT applied:\n  ' + warnings.join('\n  '));
}

function cmdBuild() {
  if (fs.existsSync(XLSX)) {
    try { fs.closeSync(fs.openSync(XLSX, 'r+')); } catch (e) {
      if (['EBUSY', 'EPERM', 'EACCES'].includes(e.code)) die('job_listing.xlsx is open in Excel. Close it, then run `node tools/joblist.mjs build` again.', 3);
      throw e;
    }
    importExcelEdits();
  }
  const apps = loadApplications();
  // closed postings / deadline passed leave the list automatically, unless Mikelson applied (then they stay, greyed)
  const removed = [];
  for (const j of store.jobs.filter((x) => !x.hidden)) {
    const posting = postingLabel(j);
    if (!['closed', 'deadline passed'].includes(posting)) continue;
    if (!['not applied', 'skipped'].includes(applicationOf(j, apps).status)) continue;
    j.hidden = true;
    j.hidden_reason = posting === 'closed' ? 'posting closed (auto)' : 'deadline passed (auto)';
    removed.push(`${j.key} | ${posting} | ${j.company} | ${j.title}`);
  }
  if (removed.length) { save(); console.log(`Removed (closed, not applied): ${removed.length}\n  ${removed.join('\n  ')}`); }
  const rows = store.jobs.filter((j) => !j.hidden).map((j) => {
    const app = applicationOf(j, apps);
    const posting = postingLabel(j);
    const defense = j.category === 'defense (not eligible)';
    const active = !['closed', 'deadline passed'].includes(posting);
    const rankable = !defense && active && app.status === 'not applied' && j.fit != null;
    const score = (j.fit ?? 0) + (j.country === 'Canada' ? cfg.canada_bonus : 0);
    const group = rankable ? 0 : app.status !== 'not applied' && app.status !== 'skipped' ? 1 : defense ? 4 : !active ? 3 : 2;
    return { j, app, posting, rankable, score, group };
  });
  rows.sort((a, b) => a.group - b.group || b.score - a.score || String(b.j.date_posted).localeCompare(String(a.j.date_posted)));
  let rank = 0;
  for (const r of rows) r.rank = r.rankable ? ++rank : '';

  const cols = ['rank', 'fit', 'application', 'posting', 'deadline', 'date_posted', 'type', 'country', 'company', 'role', 'location', 'category', 'source', 'notes', 'my_notes', 'key', 'last_checked', 'link'];
  const q = (v) => { const s = v == null ? '' : String(v); return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const lines = rows.map(({ j, app, posting, rank }) => [rank, j.fit ?? '', app.status, posting, j.deadline || '', j.date_posted || '', j.type, j.country, j.company, j.title, j.location,
    j.category, sourceLabel(j), [j.notes, j.fit == null && j.category !== 'defense (not eligible)' ? 'fit not scored yet' : ''].filter(Boolean).join('; '), j.my_notes || '', j.key, j.last_checked || '', j.apply_url].map(q).join(','));
  fs.writeFileSync(CSV, '﻿' + [cols.join(','), ...lines].join('\r\n') + '\r\n');
  // remember what this build wrote, so the next build can tell what Mikelson changed in Excel
  for (const { j, app } of rows) j.excel = { fit: j.fit ?? null, application: app.status };
  save();
  const out = execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', PS1, '-Mode', 'write', '-Csv', CSV, '-Xlsx', XLSX, '-Today', today(), '-SoonDays', String(cfg.deadline_soon_days)], { encoding: 'utf8' });
  process.stdout.write(out);
  const c = (f) => rows.filter(f).length;
  console.log(`Ranked (open, not applied): ${rank} · applied/in progress: ${c((r) => r.group === 1)} · closed: ${c((r) => r.group === 3)} · defense: ${c((r) => r.group === 4)} · not scored: ${c((r) => r.j.fit == null)}`);
}

// ---------- small commands ----------
function die(msg, code = 1) { console.error(msg); process.exit(code); }
function cmdFit() {
  for (const a of positional) {
    const [k, v] = a.split('=');
    const j = byKey(k) || die(`No job ${k}`);
    j.fit = Number(v);
    if (!(j.fit >= 0 && j.fit <= 100)) die(`Bad fit for ${k}: ${v}`);
  }
  save();
  console.log(`Fit set for ${positional.length} job(s).`);
}
function cmdHide(hidden) {
  const reason = opt('--reason') || '';
  for (const k of positional) { const j = byKey(k) || die(`No job ${k}`); j.hidden = hidden; j.hidden_reason = hidden ? reason : ''; }
  save();
  console.log(`${hidden ? 'Hidden' : 'Unhidden'}: ${positional.length}`);
}
function cmdMark() {
  const [k, status] = positional;
  const j = byKey(k) || die(`No job ${k}`);
  if (!['applied', 'skipped', 'clear'].includes(status)) die('Status must be applied, skipped or clear');
  j.manual_status = status === 'clear' ? undefined : status;
  j.manual_status_date = status === 'clear' ? undefined : opt('--date') || today();
  save();
  console.log(`${j.key} ${j.company} – ${j.title}: ${status}`);
}
function cmdNote() {
  const [k, ...text] = positional;
  const j = byKey(k) || die(`No job ${k}`);
  j.notes = text.join(' ');
  save();
  console.log(`${j.key}: note ${j.notes ? 'saved' : 'cleared'}`);
}
function cmdImportCsv() {
  const file = positional[0] || die('Give the CSV file');
  const text = fs.readFileSync(file, 'utf8').replace(/^﻿/, '');
  const parse = (line) => { const out = []; let cur = '', inq = false; for (let i = 0; i < line.length; i++) { const ch = line[i]; if (inq) { if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') inq = false; else cur += ch; } else if (ch === '"') inq = true; else if (ch === ',') { out.push(cur); cur = ''; } else cur += ch; } out.push(cur); return out; };
  const [head, ...lines] = text.split(/\r?\n/).filter(Boolean);
  const H = parse(head);
  let n = 0;
  for (const line of lines) {
    const r = Object.fromEntries(parse(line).map((v, i) => [H[i], v]));
    const url = cleanUrl(r.link);
    if (store.jobs.some((j) => sameJob(j, { company: r.company, apply_url: url }))) continue;
    const country = r.country || countryOf(r.location);
    store.jobs.push({
      key: nextKey(), runway_id: null, company: r.company, title: r.role, location: r.location, country, work_mode: '',
      type: r.type || typeOf(r.role), category: /defense/.test(r.category) ? 'defense (not eligible)' : categoryOf(r.role, r.company),
      date_posted: r.date_posted, apply_url: url, runway_url: '', source: 'first list (2026-11-10)',
      fit: r.fit_percent === '' || r.fit_percent == null ? null : Number(r.fit_percent),
      notes: /US defense/.test(r.notes || '') ? '' : r.notes || '', hidden: false, first_seen: '2026-11-10',
    });
    n++;
  }
  save();
  console.log(`Imported ${n} jobs.`);
}

const commands = {
  fetch: cmdFetch, toptier: cmdTopTier, check: cmdCheck, build: cmdBuild, fit: cmdFit, note: cmdNote, mark: cmdMark, 'import-csv': cmdImportCsv,
  hide: () => cmdHide(true), unhide: () => cmdHide(false),
  keep: () => { for (const k of positional) { const j = byKey(k) || die(`No job ${k}`); j.grad_ok = true; } save(); console.log(`Kept (graduate flag dismissed): ${positional.length}`); },
  pending: () => { const p = store.jobs.filter((j) => !j.hidden && j.fit == null && j.category !== 'defense (not eligible)'); console.log(`${p.length} job(s) without a fit score:`); printJobs(p); },
  show: () => console.log(JSON.stringify(byKey(positional[0]) || die('No such job'), null, 2)),
};
if (!commands[cmd]) die(fs.readFileSync(new URL(import.meta.url), 'utf8').split('\n').filter((l) => l.startsWith('//')).join('\n'));
await commands[cmd]();
