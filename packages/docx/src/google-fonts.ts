import {
  classifyCjkFont,
  scriptPreloadNamesForText,
  GOOGLE_FONT_SUBSTITUTES,
  SCRIPT_GOOGLE_FONTS,
  type FontPreloadEntry,
} from '@silurus/ooxml-core';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocxDocumentModel,
} from './types.js';

/** Theme-referenced typefaces commonly used by DOCX templates.
 *
 *  {@link GOOGLE_FONT_SUBSTITUTES} supplies the Office substitutes (Calibri →
 *  Carlito, Cambria → Caladea), the popular free web fonts and the Arabic Noto
 *  fallbacks — shared with pptx/xlsx. {@link SCRIPT_GOOGLE_FONTS} adds the
 *  CJK (KR/SC/TC/JP) / Cyrillic / Thai / Devanagari / Hebrew Noto faces the
 *  renderer appends to the font chain (CJK ordered by document language). Both
 *  load only when `useGoogleFonts` is on — no binaries ship in the bundle. DOCX
 *  currently has no format-specific additions. */
export const DOCX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  ...GOOGLE_FONT_SUBSTITUTES,
  ...SCRIPT_GOOGLE_FONTS,
};

/** Yield every rendered text string in the document model so the preload step
 *  can detect which scripts are actually present. Walks the body, headers /
 *  footers, footnotes / endnotes and nested tables (text-only). Comments are
 *  not painted on the page, so they are excluded. */
function* docxTextRuns(doc: DocxDocumentModel): Generator<string> {
  const fromRuns = function* (runs: DocParagraph['runs']): Generator<string> {
    for (const r of runs) {
      if (r.type === 'text') yield r.text;
      else if (r.type === 'field') yield r.fallbackText;
      else if (r.type === 'shape') {
        for (const t of r.textBlocks ?? []) yield t.text;
      }
    }
  };
  const walk = function* (el: BodyElement): Generator<string> {
    if ('runs' in el) yield* fromRuns((el as DocParagraph).runs);
    if ('rows' in el) {
      for (const row of (el as DocTable).rows) {
        for (const cell of row.cells) {
          for (const child of cell.content) yield* walk(child as BodyElement);
        }
      }
    }
  };
  const walkBody = function* (body: BodyElement[] | undefined): Generator<string> {
    for (const el of body ?? []) yield* walk(el);
  };

  yield* walkBody(doc.body);
  for (const hf of [doc.headers, doc.footers]) {
    for (const which of [hf?.default, hf?.first, hf?.even]) {
      yield* walkBody(which?.body);
    }
  }
  for (const note of [...(doc.footnotes ?? []), ...(doc.endnotes ?? [])]) {
    yield* walkBody(note.content);
  }
}

/**
 * The font-family names to preload for a document: the theme major/minor fonts,
 * plus only the script-fallback Noto faces whose script the document's TEXT
 * actually contains ({@link scriptPreloadNamesForText}). The renderer's font
 * fallback chains still END with the full Noto set, but eagerly fetching the
 * multi-MB CJK families for a document that has no CJK glyphs would block first
 * paint for nothing; an un-preloaded face loads lazily if it ever proves needed.
 *
 * Single source of truth shared by the main-thread `load()` and the render
 * worker. Both derive the set from the SAME parsed {@link DocxDocumentModel}, so
 * they preload an identical set — worker/main rendering must stay
 * pixel-equivalent. (Fonts must also be loaded before pagination, which measures
 * text; both callers await this before paginating.)
 */
export function docxFontPreloadNames(
  doc: DocxDocumentModel,
): (string | null | undefined)[] {
  const cjkLang =
    classifyCjkFont(doc.majorFont) ?? classifyCjkFont(doc.minorFont) ?? null;
  return [
    doc.majorFont,
    doc.minorFont,
    ...scriptPreloadNamesForText(docxTextRuns(doc), cjkLang),
  ];
}
