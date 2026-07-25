import type { DocParagraph } from '../types.js';

export function isInklessParagraph(paragraph: Readonly<DocParagraph>): boolean {
  return !(paragraph.runs ?? []).some((candidate) => {
    if (candidate.type === 'text') return candidate.text.length > 0;
    return true;
  });
}

/** §17.3.1.29 and §17.3.2.41: a vanished paragraph mark with no inline
 * content owns a source occurrence but contributes neither flow nor ink. */
export function isFullyHiddenParagraph(paragraph: Readonly<DocParagraph>): boolean {
  return paragraph.markVanish === true && isInklessParagraph(paragraph);
}
