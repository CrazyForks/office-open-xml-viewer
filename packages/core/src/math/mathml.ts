import type { MathNode, MathStyle } from '../types/math';

// Convert the shared OMML AST to MathML for MathJax. MathJax owns the typesetting
// (spacing, italic/upright, stretchy glyphs), so this layer only maps structure.

const ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ESC[c]);

function variant(style: MathStyle): string | null {
  switch (style) {
    case 'roman':
      return 'normal';
    case 'bold':
      return 'bold';
    case 'boldItalic':
      return 'bold-italic';
    case 'italic':
      return null; // MathML default for <mi> single letters is italic
  }
}

const BIN = '+−±∓×÷·∗⋅∘∙*/';
const REL = '=≠<>≤≥≈≡∼≅≃→←↔⇒∈∉⊂⊆⊃⊇∝≪≫⊥≔';
const OPEN = '([{⟨⌈⌊';
const CLOSE = ')]}⟩⌉⌋';
const PUNCT = ',;';

/** Tokenize a run's text into MathML token elements (mi/mn/mo) by character class. */
function runToMathML(text: string, style: MathStyle): string {
  const mv = variant(style);
  const mvAttr = mv ? ` mathvariant="${mv}"` : '';
  let out = '';
  let numBuf = '';
  const flushNum = () => {
    if (numBuf) {
      out += `<mn${mvAttr}>${esc(numBuf)}</mn>`;
      numBuf = '';
    }
  };
  for (const ch of text) {
    if (ch === ' ') {
      flushNum();
      continue;
    }
    if (ch >= '0' && ch <= '9') {
      numBuf += ch;
      continue;
    }
    flushNum();
    if (BIN.includes(ch) || REL.includes(ch) || PUNCT.includes(ch)) {
      out += `<mo>${esc(ch)}</mo>`;
    } else if (OPEN.includes(ch)) {
      out += `<mo fence="true" stretchy="false">${esc(ch)}</mo>`;
    } else if (CLOSE.includes(ch)) {
      out += `<mo fence="true" stretchy="false">${esc(ch)}</mo>`;
    } else {
      out += `<mi${mvAttr}>${esc(ch)}</mi>`;
    }
  }
  flushNum();
  return out;
}

function seq(nodes: MathNode[]): string {
  return nodes.map(nodeToMathML).join('');
}

/** Wrap a node list in an mrow so it forms one argument. */
function row(nodes: MathNode[]): string {
  return `<mrow>${seq(nodes)}</mrow>`;
}

function nodeToMathML(node: MathNode): string {
  switch (node.kind) {
    case 'run':
      return runToMathML(node.text, node.style);
    case 'group':
      return row(node.items);
    case 'fraction':
      return `<mfrac${node.bar === false ? ' linethickness="0"' : ''}>${row(node.num)}${row(node.den)}</mfrac>`;
    case 'sup':
      return `<msup>${row(node.base)}${row(node.sup ?? [])}</msup>`;
    case 'sub':
      return `<msub>${row(node.base)}${row(node.sub ?? [])}</msub>`;
    case 'subSup':
      return `<msubsup>${row(node.base)}${row(node.sub ?? [])}${row(node.sup ?? [])}</msubsup>`;
    case 'nary':
      return naryToMathML(node);
    case 'delimiter':
      return delimiterToMathML(node);
    case 'radical':
      return node.index && node.index.length
        ? `<mroot>${row(node.radicand)}${row(node.index)}</mroot>`
        : `<msqrt>${seq(node.radicand)}</msqrt>`;
    case 'limit':
      return limitToMathML(node);
    case 'array':
      return arrayToMathML(node);
    case 'groupChr': {
      const b = row(node.base);
      const mo = `<mo stretchy="true">${esc(node.char)}</mo>`;
      return node.pos === 'top' ? `<mover>${b}${mo}</mover>` : `<munder>${b}${mo}</munder>`;
    }
    case 'bar': {
      // Tight stretchy overline/underline (like \overline). NOTE: accent="true" adds a
      // gap above the base in MathJax, so we omit it — the stretchy bar hugs the base.
      const b = row(node.base);
      const barOp = '<mo stretchy="true">&#x00AF;</mo>';
      return node.pos === 'bot' ? `<munder>${b}${barOp}</munder>` : `<mover>${b}${barOp}</mover>`;
    }
    case 'accent':
      return accentToMathML(node);
    case 'func':
      return `<mrow>${row(node.name)}<mo>&#x2061;</mo>${row(node.arg)}</mrow>`;
  }
}

// Map an OMML accent character to a MathJax-friendly accent. Combining marks (zero
// advance) float when used as <mo> content, so we translate to spacing equivalents.
const ACCENT_MAP: Record<string, string> = {
  '̀': '`', // grave
  '́': '´', // acute
  '̂': '^', // circumflex / hat
  '̃': '~', // tilde
  '̆': '˘', // breve
  '̇': '˙', // dot above
  '̈': '¨', // diaeresis
  '̌': 'ˇ', // caron
  '⃗': '→', // vector arrow
  '⃖': '←', // left arrow
};
// Overline / macron family → rendered as a tight stretchy bar (\overline), no accent gap.
const OVERLINE_CHARS = new Set(['̅', '̄', '¯', '‾', '̲', '̳']);

function accentToMathML(node: Extract<MathNode, { kind: 'accent' }>): string {
  const b = row(node.base);
  if (OVERLINE_CHARS.has(node.char)) {
    return `<mover>${b}<mo stretchy="true">&#x00AF;</mo></mover>`;
  }
  const ch = ACCENT_MAP[node.char] ?? node.char;
  const stretchy = ch === '→' || ch === '←' ? 'true' : 'false';
  return `<mover accent="true">${b}<mo stretchy="${stretchy}">${esc(ch)}</mo></mover>`;
}

function limitToMathML(node: Extract<MathNode, { kind: 'limit' }>): string {
  const base = row(node.base);
  const lower = node.lower && node.lower.length ? row(node.lower) : null;
  const upper = node.upper && node.upper.length ? row(node.upper) : null;
  if (lower && upper) return `<munderover>${base}${lower}${upper}</munderover>`;
  if (lower) return `<munder>${base}${lower}</munder>`;
  if (upper) return `<mover>${base}${upper}</mover>`;
  return base;
}

function arrayToMathML(node: Extract<MathNode, { kind: 'array' }>): string {
  const cols = Math.max(1, ...node.rows.map((r) => r.length));
  let colalign: string;
  if (node.align === 'eq') {
    // alternating right/left so cells align at the relation (& marker)
    colalign = Array.from({ length: cols }, (_, i) => (i % 2 === 0 ? 'right' : 'left')).join(' ');
  } else if (node.align === 'left') {
    colalign = 'left';
  } else {
    colalign = 'center';
  }
  const rows = node.rows
    .map((cells) => `<mtr>${cells.map((c) => `<mtd>${seq(c)}</mtd>`).join('')}</mtr>`)
    .join('');
  return `<mtable columnalign="${colalign}" rowspacing="0.2em" columnspacing="0.3em">${rows}</mtable>`;
}

function naryToMathML(node: Extract<MathNode, { kind: 'nary' }>): string {
  const op = `<mo largeop="true" movablelimits="true">${esc(node.op)}</mo>`;
  const sub = node.sub ?? [];
  const sup = node.sup ?? [];
  let operator: string;
  if (sub.length && sup.length) operator = `<munderover>${op}${row(sub)}${row(sup)}</munderover>`;
  else if (sub.length) operator = `<munder>${op}${row(sub)}</munder>`;
  else if (sup.length) operator = `<mover>${op}${row(sup)}</mover>`;
  else operator = op;
  return `<mrow>${operator}${seq(node.body)}</mrow>`;
}

function delimiterToMathML(node: Extract<MathNode, { kind: 'delimiter' }>): string {
  const beg = `<mo fence="true" stretchy="true">${esc(node.begChar || '(')}</mo>`;
  const end = `<mo fence="true" stretchy="true">${esc(node.endChar || ')')}</mo>`;
  const inner = node.items.map((g) => row(g)).join('<mo separator="true">,</mo>');
  return `<mrow>${beg}${inner}${end}</mrow>`;
}

/** Full MathML document string for a formula. `display` selects block vs inline. */
export function mathToMathML(nodes: MathNode[], display: boolean): string {
  const mode = display ? 'block' : 'inline';
  return `<math xmlns="http://www.w3.org/1998/Math/MathML" display="${mode}">${seq(nodes)}</math>`;
}
