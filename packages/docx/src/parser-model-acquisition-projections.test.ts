import { describe, expect, it } from 'vitest';
import {
  bodyAcquisitionInputProjections,
  normalizeInternalDocumentModel,
  numberingMarkerShapeInput,
  paragraphAcquisitionInput,
  paragraphMarkShapeInput,
  tableColumnLayoutInput,
  tableFormatInput,
  tableParticipatesInOrdinaryFlow,
} from './parser-model.js';
import type { DocParagraph, DocxDocumentModel } from './types.js';
import type { SourceRef } from './layout/types.js';

const paragraph = (): DocParagraph => ({
  type: 'paragraph',
  alignment: 'left',
  indentLeft: 0,
  indentRight: 0,
  indentFirst: 0,
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [{
    type: 'text',
    text: 'cached parser fact',
    fontSize: 10,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    color: null,
    fontFamily: null,
    isLink: false,
    background: null,
    vertAlign: null,
  }],
} as unknown as DocParagraph);

const document = (bodyParagraph: DocParagraph): DocxDocumentModel => ({
  section: {
    pageWidth: 612,
    pageHeight: 792,
    marginTop: 72,
    marginRight: 72,
    marginBottom: 72,
    marginLeft: 72,
    headerDistance: 36,
    footerDistance: 36,
  },
  body: [bodyParagraph],
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
} as unknown as DocxDocumentModel);

const bodySource = (path: number[] = [0]): SourceRef => ({
  story: 'body',
  storyInstance: 'body',
  path,
});

describe('parser-to-body-acquisition projection capability', () => {
  it('is one frozen identity-preserving record without compatibility wrappers', () => {
    expect(Object.isFrozen(bodyAcquisitionInputProjections)).toBe(true);
    expect(Object.keys(bodyAcquisitionInputProjections).sort()).toEqual([
      'numberingMarkerShapeInput',
      'paragraphAcquisitionInput',
      'paragraphMarkShapeInput',
      'tableColumnLayoutInput',
      'tableFormatInput',
      'tableParticipatesInOrdinaryFlow',
    ]);
    expect(bodyAcquisitionInputProjections.numberingMarkerShapeInput)
      .toBe(numberingMarkerShapeInput);
    expect(bodyAcquisitionInputProjections.paragraphAcquisitionInput)
      .toBe(paragraphAcquisitionInput);
    expect(bodyAcquisitionInputProjections.paragraphMarkShapeInput)
      .toBe(paragraphMarkShapeInput);
    expect(bodyAcquisitionInputProjections.tableColumnLayoutInput)
      .toBe(tableColumnLayoutInput);
    expect(bodyAcquisitionInputProjections.tableFormatInput).toBe(tableFormatInput);
    expect(bodyAcquisitionInputProjections.tableParticipatesInOrdinaryFlow)
      .toBe(tableParticipatesInOrdinaryFlow);
  });

  it('reuses only one document-scoped paragraph/source parser-fact projection', () => {
    const modelParagraph = paragraph();
    const firstDocument = normalizeInternalDocumentModel(document(modelParagraph));
    const project = firstDocument.bodyModelGateway.acquisitionInputs.paragraphAcquisitionInput;

    const first = project(modelParagraph, bodySource());
    const equivalentSource = project(modelParagraph, bodySource());
    const otherSource = project(modelParagraph, bodySource([1]));
    const equivalentParagraph = project(paragraph(), bodySource());

    expect(equivalentSource).toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(otherSource).not.toBe(first);
    expect(equivalentParagraph).not.toBe(first);

    const secondDocument = normalizeInternalDocumentModel(document(modelParagraph));
    expect(secondDocument.bodyModelGateway.acquisitionInputs.paragraphAcquisitionInput(
      modelParagraph,
      bodySource(),
    )).not.toBe(first);
  });
});
