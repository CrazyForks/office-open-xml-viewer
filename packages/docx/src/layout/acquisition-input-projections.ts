import type {
  DocParagraph,
  DocTable,
  NumberingInfo,
} from '../types.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type {
  NumberingMarkerShapeInput,
  SourceRef,
  TableColumnLayoutInput,
  TableFormatInput,
} from './types.js';

/** Parser-owned fact projections needed by otherwise parser-independent body
 * acquisition. Layout consumes this required capability record instead of
 * importing parser-model implementation or private wire fields. */
export interface BodyAcquisitionInputProjections {
  readonly numberingMarkerShapeInput: (
    numbering: NumberingInfo,
    fallbackFontSizePt: number,
  ) => NumberingMarkerShapeInput;
  readonly paragraphMarkShapeInput: (
    paragraph: DocParagraph,
  ) => NumberingMarkerShapeInput | undefined;
  readonly tableFormatInput: (
    table: Readonly<DocTable>,
  ) => TableFormatInput;
  readonly tableColumnLayoutInput: (
    table: Readonly<DocTable>,
    availableWidthPt: number,
    intrinsicWidths: (
      cell: Readonly<DocTable['rows'][number]['cells'][number]>,
    ) => Readonly<{ minWidthPt: number; maxWidthPt: number }>,
    maximumWidthPt?: number,
  ) => TableColumnLayoutInput;
  readonly tableParticipatesInOrdinaryFlow: (
    table: Readonly<DocTable>,
  ) => boolean;
  readonly paragraphAcquisitionInput: (
    paragraph: DocParagraph,
    source: SourceRef,
  ) => ParagraphAcquisitionInput;
}
