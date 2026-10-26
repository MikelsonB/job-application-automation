// Cover letter template — same header, font and colors as the resume.
import { Document, Paragraph } from 'docx';
import { SIZE, PAGE } from './style.mjs';
import { richRuns, createFlow, header, docDefaults } from './common.mjs';

const MARGIN = { top: PAGE.margin.top, bottom: 1080, left: 1440, right: 1440 };

export function buildCoverLetter(letter, person) {
  const fr = letter.lang === 'fr';
  const flow = createFlow();
  const children = header(person, flow, { rule: true });
  const add = (type, text) =>
    children.push(new Paragraph({ spacing: flow.spacing(type), children: richRuns(text, { size: SIZE.entry }) }));

  if (letter.date) add('date', letter.date);
  for (const line of letter.recipient ?? []) add('recipient', line);
  if (letter.subject) add('subject', `**${letter.subject.replace(/\*\*/g, '')}**`);
  add('salutation', letter.salutation ?? (fr ? 'Madame, Monsieur,' : 'Dear Hiring Manager,'));
  for (const para of letter.paragraphs ?? []) add('para', para);
  add('closing', letter.closing ?? (fr ? 'Veuillez agréer mes salutations distinguées,' : 'Sincerely,'));
  add('signature', letter.signature ?? person.name);

  const doc = new Document({
    creator: person.name,
    lastModifiedBy: person.name,
    title: `${person.name} - ${fr ? 'Lettre de présentation' : 'Cover Letter'}`,
    styles: docDefaults(),
    sections: [{
      properties: { page: { size: { width: PAGE.width, height: PAGE.height }, margin: MARGIN } },
      children,
    }],
  });
  return { doc, warnings: flow.warnings };
}
