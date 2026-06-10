// Generates packages/core/src/text/bidi/char-data.generated.ts from the Unicode
// Character Database (UCD). Run: `node packages/core/scripts/gen-bidi-char-data.mjs`
//
// Emits, as gap-free run-length tables + flat pair arrays:
//   - Bidi_Class for every code point (DerivedBidiClass.txt + its @missing block
//     defaults for unassigned code points),
//   - Bidi_Mirroring_Glyph pairs (BidiMirroring.txt),
//   - Bidi_Paired_Bracket + type (BidiBrackets.txt).
//
// The generated file is data straight from the UCD — never hand-edit it.

import { writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const VERSION = '17.0.0';
const BASE = `https://www.unicode.org/Public/${VERSION}/ucd`;
const OUT = join(
  dirname(fileURLToPath(import.meta.url)),
  '..',
  'src',
  'text',
  'bidi',
  'char-data.generated.ts',
);
const MAX_CP = 0x110000;

// Canonical class index order. Index is what we store; names map back in char-data.ts.
const CLASS_NAMES = [
  'L', 'R', 'AL',
  'EN', 'ES', 'ET', 'AN', 'CS', 'NSM', 'BN',
  'B', 'S', 'WS', 'ON',
  'LRE', 'LRO', 'RLE', 'RLO', 'PDF',
  'LRI', 'RLI', 'FSI', 'PDI',
];
const CLASS_INDEX = Object.fromEntries(CLASS_NAMES.map((n, i) => [n, i]));

// Long Bidi_Class value names (used by @missing lines) -> abbreviations.
const LONG_TO_ABBR = {
  Left_To_Right: 'L', Right_To_Left: 'R', Arabic_Letter: 'AL',
  European_Number: 'EN', European_Separator: 'ES', European_Terminator: 'ET',
  Arabic_Number: 'AN', Common_Separator: 'CS', Nonspacing_Mark: 'NSM',
  Boundary_Neutral: 'BN', Paragraph_Separator: 'B', Segment_Separator: 'S',
  White_Space: 'WS', Other_Neutral: 'ON',
  Left_To_Right_Embedding: 'LRE', Left_To_Right_Override: 'LRO',
  Right_To_Left_Embedding: 'RLE', Right_To_Left_Override: 'RLO',
  Pop_Directional_Format: 'PDF',
  Left_To_Right_Isolate: 'LRI', Right_To_Left_Isolate: 'RLI',
  First_Strong_Isolate: 'FSI', Pop_Directional_Isolate: 'PDI',
};

async function fetchText(path) {
  const url = `${BASE}/${path}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`GET ${url} -> ${res.status}`);
  return res.text();
}

function abbr(name) {
  const a = LONG_TO_ABBR[name] ?? name; // assigned lines already use abbreviations
  if (!(a in CLASS_INDEX)) throw new Error(`unknown Bidi_Class: ${name}`);
  return a;
}

function buildBidiClass(text) {
  // Default everything to L, then apply @missing block defaults (in file order),
  // then overlay explicit assigned ranges. Result covers [0, MAX_CP) with no gaps.
  const cls = new Uint8Array(MAX_CP).fill(CLASS_INDEX.L);

  const apply = (start, end, idx) => {
    for (let cp = start; cp <= end && cp < MAX_CP; cp++) cls[cp] = idx;
  };

  for (const raw of text.split('\n')) {
    const missing = raw.match(/^#\s*@missing:\s*([0-9A-Fa-f]+)\.\.([0-9A-Fa-f]+)\s*;\s*(\w+)/);
    if (missing) {
      apply(parseInt(missing[1], 16), parseInt(missing[2], 16), CLASS_INDEX[abbr(missing[3])]);
    }
  }
  for (const raw of text.split('\n')) {
    const line = raw.split('#')[0].trim();
    if (!line) continue;
    const m = line.match(/^([0-9A-Fa-f]+)(?:\.\.([0-9A-Fa-f]+))?\s*;\s*(\w+)$/);
    if (!m) continue;
    const start = parseInt(m[1], 16);
    const end = m[2] ? parseInt(m[2], 16) : start;
    apply(start, end, CLASS_INDEX[abbr(m[3])]);
  }

  // Run-length compress into gap-free ranges: range i covers [starts[i], starts[i+1]).
  const starts = [0];
  const classes = [cls[0]];
  for (let cp = 1; cp < MAX_CP; cp++) {
    if (cls[cp] !== classes[classes.length - 1]) {
      starts.push(cp);
      classes.push(cls[cp]);
    }
  }
  return { starts, classes };
}

function buildMirror(text) {
  const flat = [];
  for (const raw of text.split('\n')) {
    const line = raw.split('#')[0].trim();
    if (!line) continue;
    const m = line.match(/^([0-9A-Fa-f]+)\s*;\s*([0-9A-Fa-f]+)$/);
    if (!m) continue;
    flat.push(parseInt(m[1], 16), parseInt(m[2], 16));
  }
  return flat;
}

function buildBrackets(text) {
  const flat = [];
  for (const raw of text.split('\n')) {
    const line = raw.split('#')[0].trim();
    if (!line) continue;
    const m = line.match(/^([0-9A-Fa-f]+)\s*;\s*([0-9A-Fa-f]+)\s*;\s*([oc])$/);
    if (!m) continue;
    flat.push(parseInt(m[1], 16), parseInt(m[2], 16), m[3] === 'o' ? 0 : 1);
  }
  return flat;
}

function fmtArray(nums, perLine = 16) {
  const out = [];
  for (let i = 0; i < nums.length; i += perLine) {
    out.push('  ' + nums.slice(i, i + perLine).join(', ') + ',');
  }
  return out.join('\n');
}

async function main() {
  const [derived, mirroring, brackets] = await Promise.all([
    fetchText('extracted/DerivedBidiClass.txt'),
    fetchText('BidiMirroring.txt'),
    fetchText('BidiBrackets.txt'),
  ]);

  const { starts, classes } = buildBidiClass(derived);
  const mirror = buildMirror(mirroring);
  const bracket = buildBrackets(brackets);

  const body = `// AUTO-GENERATED from the Unicode Character Database (UCD ${VERSION}).
// Source: ${BASE}/extracted/DerivedBidiClass.txt, ${BASE}/BidiMirroring.txt,
//         ${BASE}/BidiBrackets.txt
// DO NOT EDIT — regenerate via packages/core/scripts/gen-bidi-char-data.mjs
/* eslint-disable */

export const UNICODE_VERSION = '${VERSION}';

/** Canonical Bidi_Class index order. */
export const BIDI_CLASS_NAMES = [
  ${CLASS_NAMES.map((n) => `'${n}'`).join(', ')},
] as const;

/** Range starts (sorted, gap-free): range i covers [BIDI_RANGE_STARTS[i], BIDI_RANGE_STARTS[i+1]). */
export const BIDI_RANGE_STARTS: number[] = [
${fmtArray(starts)}
];

/** Bidi_Class index for the range beginning at the matching BIDI_RANGE_STARTS entry. */
export const BIDI_RANGE_CLASS: number[] = [
${fmtArray(classes)}
];

/** Flat [codePoint, mirrorGlyph, ...] pairs from BidiMirroring.txt. */
export const MIRROR_FLAT: number[] = [
${fmtArray(mirror)}
];

/** Flat [codePoint, pairedBracket, type(0=open,1=close), ...] triples from BidiBrackets.txt. */
export const BRACKET_FLAT: number[] = [
${fmtArray(bracket, 15)}
];
`;

  await writeFile(OUT, body, 'utf8');
  console.log(
    `wrote ${OUT}\n  bidi ranges: ${starts.length}\n  mirror pairs: ${mirror.length / 2}\n  brackets: ${bracket.length / 3}`,
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
