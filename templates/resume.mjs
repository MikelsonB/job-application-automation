// Resume template — reproduces the layout of master/Mikelson_Resume_Original.pdf.
import { Document, Paragraph, ExternalHyperlink, Tab, TabStopType, LevelFormat, AlignmentType, UnderlineType } from 'docx';
import { COLOR, SIZE, PAGE, TEXT_WIDTH, BULLET_INDENT, SCALE } from './style.mjs';
import { run, richRuns, createFlow, header, sectionHeading, docDefaults } from './common.mjs';

const TITLES = { experience: 'Work Experience', education: 'Education', projects: 'Projects', skills: 'Skills' };

// "Company - Title ............ dates" row (also used for schools)
function entryRow(type, flow, strong, rest, dates) {
  const children = [run(strong, { bold: true, size: SIZE.entry })];
  if (rest) children.push(run(` - ${rest}`, { size: SIZE.entry }));
  if (dates) children.push(run('', { size: SIZE.entry, color: COLOR.date, children: [new Tab(), dates] }));
  return new Paragraph({
    keepNext: true,
    spacing: flow.spacing(type),
    tabStops: [{ type: TabStopType.RIGHT, position: TEXT_WIDTH }],
    children,
  });
}

const sections = {
  experience(data, flow) {
    const out = [];
    for (const job of data.experience ?? []) {
      out.push(entryRow('job', flow, job.company, job.title, job.dates));
      for (const b of job.bullets ?? []) {
        out.push(new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          spacing: flow.spacing('bullet'),
          children: richRuns(b, { color: COLOR.bulletText }),
        }));
      }
    }
    return out;
  },

  education(data, flow) {
    const out = [];
    for (const ed of data.education ?? []) {
      out.push(entryRow('school', flow, ed.school, ed.title, ed.dates));
      for (const d of ed.details ?? []) {
        out.push(new Paragraph({ spacing: flow.spacing('detail'), children: richRuns(d, { size: SIZE.detail }) }));
      }
    }
    return out;
  },

  projects(data, flow) {
    const out = [];
    for (const p of data.projects ?? []) {
      const nameOpts = { bold: true, italics: true, size: SIZE.entry };
      const name = p.url
        ? new ExternalHyperlink({ link: p.url, children: [run(p.name, { ...nameOpts, underline: { type: UnderlineType.SINGLE, color: COLOR.text } })] })
        : run(p.name, nameOpts);
      const children = [name];
      if (p.tech) children.push(run(' - ', { size: SIZE.entry }), run(p.tech, { bold: true }));
      out.push(new Paragraph({ keepNext: true, spacing: flow.spacing('project'), children }));
      if (p.description) out.push(new Paragraph({ spacing: flow.spacing('desc'), children: richRuns(p.description) }));
    }
    return out;
  },

  skills(data, flow) {
    return (data.skills ?? []).map((s) => new Paragraph({
      spacing: flow.spacing('skill'),
      children: [run(`${s.label}:`, { bold: true, size: SIZE.entry }), ...richRuns(` ${s.items}`, { size: SIZE.entry })],
    }));
  },
};

export function buildResume(data) {
  const flow = createFlow();
  const children = [...header(data, flow)];
  for (const key of data.sections ?? ['experience', 'education', 'projects', 'skills']) {
    if (!sections[key]) throw new Error(`Unknown section "${key}" in resume.yaml`);
    if (!data[key]?.length) continue;
    // Heading first: the flow computes spacing in document order.
    children.push(sectionHeading(data.titles?.[key] ?? TITLES[key], flow));
    children.push(...sections[key](data, flow));
  }

  const doc = new Document({
    creator: data.name,
    lastModifiedBy: data.name,
    title: `${data.name} - Resume`,
    styles: docDefaults(),
    numbering: {
      config: [{
        reference: 'bullets',
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: '•',
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: { indent: BULLET_INDENT },
            run: { font: 'Arial', size: SIZE.bullet, color: COLOR.bulletText, scale: SCALE },
          },
        }],
      }],
    },
    sections: [{
      properties: { page: { size: { width: PAGE.width, height: PAGE.height }, margin: PAGE.margin } },
      children,
    }],
  });
  return { doc, warnings: flow.warnings };
}
