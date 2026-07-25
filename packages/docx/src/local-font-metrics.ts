import {
  loadLocalFontMetrics,
  normalizeLocalFontMetricFamily,
  type LoadedLocalFontMetrics,
  type LocalFontMetricRequest,
} from '@silurus/ooxml-core';
import type { DocxDocumentModel } from './types.js';
import { docxRenderedTextUsages } from './document-content.js';

/** `word-east-asian-grid-line-allocation` uses 1.3 × the selected Meiryo face's
 * hhea design box as its compatibility floor. Resolving the exact local face
 * keeps the rule version-safe without a version-specific constant. Meiryo UI is
 * deliberately excluded because it is a distinct family with different
 * metrics and no registered evidence for this multiplier. */
export function docxLocalMetricRequests(
  doc: DocxDocumentModel,
): LocalFontMetricRequest[] {
  // ECMA-376 §17.8.3.3: an embedded face is authoritative for its exact
  // family/weight/style tuple. Other tuples remain unavailable unless they
  // have their own positively identified exact face route.
  const embeddedTuples = new Set(
    (doc.embeddedFonts ?? [])
      .map((font) => {
        const weight = font.style === 'bold' || font.style === 'boldItalic' ? 700 : 400;
        const style = font.style === 'italic' || font.style === 'boldItalic' ? 'italic' : 'normal';
        return `${normalizeLocalFontMetricFamily(font.fontName)}:${weight}:${style}`;
      }),
  );
  const requests: LocalFontMetricRequest[] = [];
  const seen = new Set<string>();
  const addCandidate = (familyValue: string | null | undefined): void => {
    const family = familyValue?.trim();
    if (!family) return;
    const normalized = normalizeLocalFontMetricFamily(family);
    // This reviewed compatibility profile is the only mapping for which DOCX
    // knows an exact local full-face name. Arbitrary w:rFonts family text is not
    // a CSS local() identity and must use embedded/web/generic routing instead.
    const isMeiryo = normalized === 'meiryo' || family === 'メイリオ';
    if (!isMeiryo) return;
    const key = `${normalized}:400:normal`;
    if (embeddedTuples.has(key)) return;
    if (seen.has(key)) return;
    seen.add(key);
    requests.push({ family, localNames: ['Meiryo'], lineHeightMultiplier: 1.3 });
  };
  for (const usage of docxRenderedTextUsages(doc)) {
    if (usage.bold || usage.italic) continue;
    for (const authoredFamily of usage.fontFamilies) {
      addCandidate(authoredFamily);
    }
  }
  return requests;
}

export function loadDocxLocalFontMetrics(
  doc: DocxDocumentModel,
): Promise<LoadedLocalFontMetrics> {
  return loadLocalFontMetrics(docxLocalMetricRequests(doc));
}
