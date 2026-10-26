// Building blocks shared by the resume and cover letter templates.
import { Paragraph, TextRun, AlignmentType, BorderStyle, LineRuleType } from 'docx';
import { FONT, SCALE, COLOR, SIZE, LINE, PITCH, RULE } from './style.mjs';

const tw = (pt) => Math.round(pt * 20);

// A text run in the resume font. `opts` accepts any docx TextRun option.
export function run(text, opts = {}) {
  const o = { font: FONT, scale: SCALE, kern: 2, color: COLOR.text, size: SIZE.body, ...opts };
  if (!o.children) o.text = text;
  return new TextRun(o);
}

// Text with **bold** markers → runs.
export function richRuns(text, opts = {}) {
  if (text !== null && typeof text === 'object') {
    // Unquoted YAML text containing ": " is parsed as a key/value pair.
    throw new Error(`A line in the YAML was read as "key: value" — wrap it in double quotes:\n  ${JSON.stringify(text)}`);
  }
  return String(text ?? '')
    .split(/(\*\*[^*]+\*\*)/g)
    .filter(Boolean)
    .map((part) =>
      part.startsWith('**') && part.endsWith('**') && part.length > 4
        ? run(part.slice(2, -2), { ...opts, bold: true })
        : run(part, opts),
    );
}

// Vertical rhythm: returns the spacing of each new paragraph so that the
// baseline-to-baseline distance from the previous line matches PITCH.
export function createFlow() {
  let prev = null;
  const warnings = [];
  function spacing(type, { extra = 0 } = {}) {
    const L = LINE[type];
    if (L === undefined) throw new Error(`Unknown paragraph type: ${type}`);
    let before = 0;
    if (prev) {
      const target = PITCH[`${prev.type}>${type}`] ?? PITCH[`*>${type}`] ?? PITCH.default;
      before = target - 0.2 * prev.L - prev.extra - 0.8 * L;
      if (before < -0.01) warnings.push(`${prev.type}>${type}: needs ${before.toFixed(2)}pt`);
      before = Math.max(0, before);
    }
    prev = { type, L, extra };
    return { before: tw(before), after: 0, line: tw(L), lineRule: LineRuleType.EXACT };
  }
  return { spacing, warnings };
}

const ruleBorder = () => ({ bottom: { style: BorderStyle.SINGLE, size: RULE.size, color: COLOR.rule, space: RULE.space } });

// Centered name + contact line, identical on the resume and the cover letter.
// `rule: true` draws a full-width line under the contact line (cover letter).
export function header(data, flow, { rule = false } = {}) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: flow.spacing('name'),
      children: [run(data.name, { bold: true, size: SIZE.name })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: flow.spacing('contact', rule ? { extra: RULE.extra } : {}),
      ...(rule ? { border: ruleBorder() } : {}),
      children: [run(data.contact.join(' | '), { size: SIZE.contact })],
    }),
  ];
}

// Section title with the full-width rule under it.
export function sectionHeading(title, flow) {
  return new Paragraph({
    keepNext: true,
    spacing: flow.spacing('heading', { extra: RULE.extra }),
    indent: { left: -RULE.overhang, right: -RULE.overhang, firstLine: RULE.overhang },
    border: ruleBorder(),
    children: [run(title, { bold: true, size: SIZE.heading, color: COLOR.heading })],
  });
}

export function docDefaults() {
  return {
    default: {
      document: {
        run: { font: FONT, size: SIZE.body, color: COLOR.text },
        paragraph: { spacing: { before: 0, after: 0 } },
      },
    },
  };
}
