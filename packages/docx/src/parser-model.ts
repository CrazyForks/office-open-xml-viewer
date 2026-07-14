import type { DocxDocumentModel, DocxTextRun, FieldRun } from './types.js';

/** Parser-emitted metadata intentionally kept outside the stable public model.
 * Ordinary text and field results share these resolved WordprocessingML axes. */
export interface InternalRunSlotMetadata {
  fontFamilyEastAsia?: string | null;
  fontHint?: 'default' | 'eastAsia' | 'cs';
  rtl?: boolean;
  cs?: boolean;
  fontFamilyCs?: string | null;
  fontSizeCs?: number;
  boldCs?: boolean;
  italicCs?: boolean;
  langBidi?: string;
  langEastAsia?: string;
}

type TextOnlyMetadata = Pick<
  DocxTextRun,
  | 'ruby' | 'revision' | 'hyperlink' | 'hyperlinkAnchor'
  | 'underlineStyle' | 'underlineColor' | 'colorAuto' | 'border'
  | 'snapToGrid' | 'charSpacing' | 'charScale' | 'fitTextVal' | 'fitTextId'
  | 'position' | 'kerning' | 'eastAsianVert' | 'eastAsianVertCompress'
>;

export type InternalTextRun = DocxTextRun & InternalRunSlotMetadata;
export type InternalFieldRun = FieldRun & Partial<TextOnlyMetadata> & InternalRunSlotMetadata;
export type InternalTextBearingRun = InternalTextRun | InternalFieldRun;

export interface InternalDocxDocumentModel extends DocxDocumentModel {
  fontFamilyCharsets?: Record<string, string>;
}

export function internalFieldRun(run: FieldRun): InternalFieldRun {
  return run as InternalFieldRun;
}

export function internalDocumentModel(doc: DocxDocumentModel): InternalDocxDocumentModel {
  return doc as InternalDocxDocumentModel;
}
