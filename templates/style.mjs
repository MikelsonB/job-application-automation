// Visual style of the resume, measured from master/Mikelson_Resume_Original.pdf.
// Do not change these values unless you want to change the look of every document.

// The original renders every font at 97% of its nominal size. Word only accepts
// half-point sizes, so we keep the nominal size and squeeze the width with
// `scale: 97` — the text width (and so the line breaks) matches the original.
export const FONT = 'Raleway Thin'; // Google Fonts' static Raleway (weights 400/700) is named "Raleway Thin"
export const SCALE = 97;

export const COLOR = {
  text: '000000',
  bulletText: '0D0D0D',
  heading: '262626',
  date: '7F7F7F',
  rule: '000000',
};

// Sizes in half-points (docx unit)
export const SIZE = {
  name: 36,      // 17.46pt in the original
  contact: 22,   // 10.67pt
  heading: 28,   // 13.58pt
  entry: 22,     // company / school / project rows, skills
  body: 20,      // bullets, project descriptions, project tech list (9.7pt)
  detail: 21,    // education detail line (10.24pt)
  bullet: 25,    // bullet glyph
};

// US Letter, in twips (1pt = 20 twips)
export const PAGE = {
  width: 12240,
  height: 15840,
  margin: { top: 684, bottom: 640, left: 720, right: 720 },
};
export const TEXT_WIDTH = PAGE.width - PAGE.margin.left - PAGE.margin.right;

// Bullet list: glyph 17.46pt from the margin, text 34.92pt from the margin
export const BULLET_INDENT = { left: 698, hanging: 349 };

// Section heading rule (paragraph bottom border)
export const RULE = {
  size: 4,        // eighths of a point → 0.5pt
  space: 2,       // pt between the heading text and the rule
  overhang: 0,    // extra twips the rule sticks out past the margins (Word already adds 1.5pt, like the original)
  extra: 2.4,     // pt the border adds to the paragraph height (measured in Word)
};

// Line height (pt) of each paragraph type. Every paragraph uses "exact" line
// spacing; Word puts the baseline at 80% of the line height.
export const LINE = {
  name: 13,     // smaller than the glyphs on purpose: Word doesn't clip, and it keeps the name close to the contact line
  contact: 13,
  heading: 17,
  job: 13,
  bullet: 11.45,
  school: 13,
  detail: 12.6,
  project: 13,
  desc: 13.1,
  skill: 12.6,
  // cover letter
  date: 14.5,
  recipient: 14.5,
  subject: 14.5,
  salutation: 14.5,
  para: 14.5,
  closing: 14.5,
  signature: 14.5,
};

// Distance (pt) from the baseline of the previous line to the baseline of the
// next paragraph, measured on the original. Keys are "previous>current".
export const PITCH = {
  'name>contact': 13.13,
  'contact>heading': 21.75,
  'heading>job': 18.0,
  'heading>school': 18.0,
  'heading>project': 20.25,
  'heading>skill': 20.25,
  'job>bullet': 15.75,
  'job>job': 15.0,
  'bullet>bullet': 12.0,
  'bullet>job': 15.0,
  'bullet>heading': 21.38,
  'school>detail': 12.75,
  'school>school': 15.0,
  'detail>detail': 12.75,
  'detail>school': 15.0,
  'detail>heading': 22.5,
  'project>desc': 13.3,
  'desc>project': 13.88,
  'desc>heading': 23.25,
  'skill>skill': 12.6,
  '*>heading': 22.0,
  // cover letter
  'contact>date': 36.0,
  'date>recipient': 29.0,
  'recipient>recipient': 14.5,
  'recipient>subject': 29.0,
  'recipient>salutation': 29.0,
  'subject>salutation': 24.0,
  'salutation>para': 24.0,
  'para>para': 22.0,
  'para>closing': 24.0,
  'closing>signature': 36.0,
  default: 15.0,
};
