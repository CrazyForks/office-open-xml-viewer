import { describe, expect, it } from 'vitest';
import {
  computeSectionColumns,
  resolveDocumentLayoutSettings,
  resolveSectionLayoutContext,
} from './layout-context.js';
import { computeSectionColumns as computeColumns } from './layout-context.js';
import { resolveBodyParagraphLayoutContext } from './layout/acquisition-state.js';
import { bodyAcquisitionInputProjections } from './parser-model.js';
import type {
  DocParagraph,
  DocxDocumentModel,
  SectionProps,
} from './types.js';

const section = (): SectionProps => ({
  pageWidth: 200,
  pageHeight: 300,
  marginTop: 20,
  marginRight: 20,
  marginBottom: 20,
  marginLeft: 20,
  headerDistance: 10,
  footerDistance: 10,
  titlePage: false,
  evenAndOddHeaders: false,
});

const paragraph = (): DocParagraph => ({
  alignment: 'left',
  indentLeft: 12,
  indentRight: 6,
  indentFirst: 0,
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [],
  bidi: true,
});

const documentModel = (): DocxDocumentModel => ({
  section: section(),
  body: [],
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
});

describe('body layout context integration', () => {
  it('uses the shared section-column implementation', () => {
    expect(computeColumns).toBe(computeSectionColumns);
  });

  it('resolves body paragraph physical indents through the shared context', () => {
    const document = documentModel();
    const layoutSettings = resolveDocumentLayoutSettings(document);
    const sectionLayout = resolveSectionLayoutContext(layoutSettings, document.section);
    const context = resolveBodyParagraphLayoutContext(
      { layoutSettings, sectionLayout, acquisitionInputs: bodyAcquisitionInputProjections },
      paragraph(),
    );

    expect(context.physicalIndentLeftPt).toBe(6);
    expect(context.physicalIndentRightPt).toBe(12);
  });
});
