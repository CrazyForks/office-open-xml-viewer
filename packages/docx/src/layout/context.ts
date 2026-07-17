import type {
  ColumnGeom,
  ColumnsSpec,
  HeadersFooters,
  LineNumbering,
  PageNumType,
  SectionGeom,
  SectionProps,
} from '../types.js';
import {
  computeSectionColumns,
  type SectionLayoutContext,
} from '../layout-context.js';
import type { DeepReadonly } from './types.js';

/**
 * Section facts that must change atomically at a page-flow boundary. Keeping the
 * occurrence identity beside geometry prevents two equal-looking consecutive
 * sections from being mistaken for the same section by page-number/header logic.
 */
export interface PageFlowSectionContext {
  readonly sectionOccurrenceId: string;
  readonly geometry: Readonly<SectionGeom>;
  readonly columns: readonly Readonly<ColumnGeom>[];
  readonly textDirection: string;
}

export function createPageFlowSectionContext(input: Readonly<{
  sectionOccurrenceId: string;
  geometry: SectionGeom;
  columns: readonly Readonly<ColumnGeom>[];
  textDirection: string;
}>): PageFlowSectionContext {
  if (input.sectionOccurrenceId.length === 0) {
    throw new RangeError('Section occurrence id must not be empty');
  }
  if (input.columns.length === 0) {
    throw new RangeError('A page-flow section requires at least one column');
  }
  return Object.freeze({
    sectionOccurrenceId: input.sectionOccurrenceId,
    geometry: Object.freeze({ ...input.geometry }),
    columns: Object.freeze(input.columns.map((column) => Object.freeze({ ...column }))),
    textDirection: input.textDirection,
  });
}

/** §17.6.11 permits signed top/bottom margins, but body flow uses their distance
 * from the page edge; the sign controls header/footer overlap separately. */
export function sectionContentStartBlockPt(section: PageFlowSectionContext): number {
  return sectionBodyInsetPt(section.geometry.marginTop);
}

/** Signed top/bottom margins retain overlap policy; body placement uses distance. */
export function sectionBodyInsetPt(marginPt: number): number {
  return Math.abs(marginPt);
}

/** Physical-to-logical quarter turn for vertical section body layout. */
export function logicalSectionGeometry(physical: SectionGeom): SectionGeom {
  return {
    pageWidth: physical.pageHeight,
    pageHeight: physical.pageWidth,
    marginLeft: physical.marginTop,
    marginTop: physical.marginRight,
    marginRight: physical.marginBottom,
    marginBottom: physical.marginLeft,
    headerDistance: physical.headerDistance,
    footerDistance: physical.footerDistance,
  };
}

/** Inverse logical-to-physical quarter turn for a vertical section page box. */
export function physicalSectionGeometry(logical: SectionGeom): SectionGeom {
  return {
    pageWidth: logical.pageHeight,
    pageHeight: logical.pageWidth,
    marginTop: logical.marginLeft,
    marginRight: logical.marginTop,
    marginBottom: logical.marginRight,
    marginLeft: logical.marginBottom,
    headerDistance: logical.headerDistance,
    footerDistance: logical.footerDistance,
  };
}

export function isVerticalSectionDirection(textDirection: string): boolean {
  return textDirection === 'tbRl'
    || textDirection === 'tbRlV'
    || textDirection === 'tbLrV'
    || textDirection === 'btLr';
}

export interface SectionPageLayoutPolicy {
  readonly physicalGeometry: Readonly<SectionGeom>;
  readonly columns: DeepReadonly<ColumnsSpec> | null;
  readonly textDirection: string;
  readonly gutterPt: number;
  readonly rtlGutter: boolean;
  readonly mirrorMargins: boolean;
  readonly gutterAtTop: boolean;
  readonly bookFoldPrinting: boolean;
  readonly bookFoldRevPrinting: boolean;
  readonly printTwoOnOne: boolean;
}

export function effectivePhysicalSectionGeometry(
  policy: SectionPageLayoutPolicy,
  pageIndex: number,
): SectionGeom {
  if (!Number.isInteger(pageIndex) || pageIndex < 0) {
    throw new RangeError('Physical page index must be a non-negative integer');
  }
  let { marginTop, marginRight, marginBottom, marginLeft } = policy.physicalGeometry;
  const imposed = policy.bookFoldPrinting
    || policy.bookFoldRevPrinting
    || policy.printTwoOnOne;
  if (!imposed) {
    if (policy.gutterAtTop && !policy.mirrorMargins) marginTop += policy.gutterPt;
    else if (policy.rtlGutter) marginRight += policy.gutterPt;
    else marginLeft += policy.gutterPt;
  }
  if (policy.mirrorMargins && pageIndex % 2 === 1) {
    [marginLeft, marginRight] = [marginRight, marginLeft];
  }
  return { ...policy.physicalGeometry, marginTop, marginRight, marginBottom, marginLeft };
}

export function resolveSectionContextForPage(
  base: SectionLayoutContext,
  policy: SectionPageLayoutPolicy,
  pageIndex: number,
): SectionLayoutContext {
  const physical = effectivePhysicalSectionGeometry(policy, pageIndex);
  const geometry = isVerticalSectionDirection(policy.textDirection)
    ? logicalSectionGeometry(physical)
    : physical;
  return Object.freeze({
    ...base,
    geometry: Object.freeze(geometry),
    columns: Object.freeze(computeSectionColumns({
      ...geometry,
      titlePage: false,
      evenAndOddHeaders: false,
      columns: policy.columns,
    } as SectionProps).map((column) => Object.freeze(column))),
  });
}

export interface SectionPlacementFacts {
  readonly sectionId: string;
  readonly vAlign: string | null;
  readonly lineNumbering: Readonly<LineNumbering> | null;
  readonly docGridType: string | null;
  readonly docGridLinePitch: number | null;
  readonly docGridCharSpace: number | null;
  readonly gutterPt: number | null;
  readonly rtlGutter: boolean | null;
  readonly pageBordersAuthored: boolean;
  readonly pageBorders: Readonly<import('../types.js').PageBorders> | null;
  readonly pageGeometry: Readonly<Partial<SectionGeom>> | null;
}

/**
 * One lexical section occurrence in body order. Equal section properties do not
 * make two occurrences interchangeable: page numbering, title-page selection,
 * and line-number restart rules are occurrence-sensitive.
 */
export interface BodySectionOccurrence {
  readonly sectionOccurrenceId: string;
  readonly ordinal: number;
  /** First body item owned by this occurrence. */
  readonly startBodyIndex: number;
  /** Last body item owned by this occurrence, inclusive. */
  readonly endBodyIndex: number;
  /** The paragraph-owned sectPr marker which terminates this occurrence. */
  readonly markerBodyIndex: number | null;
  readonly final: boolean;
  /** ECMA-376 §17.6.22: how this section starts relative to its predecessor. */
  readonly startType: string;
  readonly columns: ColumnsSpec | null;
  /** Physical §17.6.13/§17.6.11 page box; writing-mode transforms happen later. */
  readonly geometry: SectionGeom;
  readonly textDirection: string | null;
  readonly pageNumType: PageNumType | null;
  readonly headers: HeadersFooters;
  readonly footers: HeadersFooters;
  readonly titlePage: boolean;
  readonly vAlign: string | null;
  readonly lineNumbering: Readonly<LineNumbering> | null;
  readonly docGridType: string | null;
  readonly docGridLinePitch: number | null;
  readonly docGridCharSpace: number | null;
  readonly gutterPt: number;
  readonly rtlGutter: boolean;
  readonly pageBordersAuthored: boolean;
  readonly pageBorders: Readonly<import('../types.js').PageBorders> | null;
  readonly placement: SectionPlacementFacts;
}

/** Complete parser-boundary projection; layout does not scan document nodes. */
export interface BodySectionIndexInput {
  readonly bodyLength: number;
  readonly occurrences: readonly BodySectionOccurrence[];
}

export interface BodySectionIndex {
  readonly occurrences: readonly BodySectionOccurrence[];
  /** Accepts body.length as the insertion point owned by the final section. */
  sectionAtBodyIndex(bodyIndex: number): BodySectionOccurrence;
}

export function sectionPageBox(section: SectionProps): SectionGeom {
  return {
    pageWidth: section.pageWidth,
    pageHeight: section.pageHeight,
    marginTop: section.marginTop,
    marginRight: section.marginRight,
    marginBottom: section.marginBottom,
    marginLeft: section.marginLeft,
    headerDistance: section.headerDistance,
    footerDistance: section.footerDistance,
  };
}

export function defaultSectionGeometry(): SectionGeom {
  return {
    pageWidth: 612, pageHeight: 792,
    marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
    headerDistance: 36, footerDistance: 36,
  };
}

/** Resolve section-owned facts after parser acquisition, without retaining the
 * document wrapper across the canonical layout boundary. */
export function resolveAcquiredSectionLayoutContext(
  section: SectionProps,
): SectionLayoutContext {
  const gridKind = section.docGridType === 'lines'
    || section.docGridType === 'linesAndChars'
    || section.docGridType === 'snapToChars'
    ? section.docGridType
    : 'none';
  return Object.freeze({
    geometry: Object.freeze(sectionPageBox(section)),
    columns: Object.freeze(computeSectionColumns(section).map((column) => Object.freeze(column))),
    columnSeparator: section.columns?.sep === true,
    grid: Object.freeze({
      kind: gridKind,
      linePitchPt: section.docGridLinePitch ?? null,
      charSpacePt: section.docGridCharSpace == null ? null : section.docGridCharSpace / 4096,
    }),
    textDirection: section.textDirection ?? 'lrTb',
    verticalAlignment: section.vAlign ?? 'top',
    ...(section.lineNumbering === null || section.lineNumbering === undefined
      ? {}
      : { lineNumbering: Object.freeze({ ...section.lineNumbering }) }),
  });
}

/**
 * Index a complete §17.6.18/§17.6.17 occurrence projection. Construction is
 * O(occurrences + body), and subsequent source-index lookup is one array access.
 */
export function createBodySectionIndex(input: BodySectionIndexInput): BodySectionIndex {
  if (!Number.isInteger(input.bodyLength) || input.bodyLength < 0 || input.occurrences.length === 0) {
    throw new RangeError('A body section index requires a non-negative length and occurrences');
  }
  const occurrenceOrdinalByBodyIndex = new Array<number>(input.bodyLength + 1);
  let expectedStart = 0;
  input.occurrences.forEach((occurrence, ordinal) => {
    const last = ordinal === input.occurrences.length - 1;
    if (
      occurrence.ordinal !== ordinal
      || occurrence.startBodyIndex !== expectedStart
      || occurrence.endBodyIndex !== (last ? input.bodyLength - 1 : occurrence.markerBodyIndex)
      || occurrence.final !== last
      || (last ? occurrence.markerBodyIndex !== null : occurrence.markerBodyIndex === null)
    ) {
      throw new RangeError(`Invalid section occurrence ${ordinal}`);
    }
    for (
      let ownedIndex = occurrence.startBodyIndex;
      ownedIndex <= occurrence.endBodyIndex;
      ownedIndex += 1
    ) {
      occurrenceOrdinalByBodyIndex[ownedIndex] = ordinal;
    }
    expectedStart = occurrence.endBodyIndex + 1;
  });
  const finalOrdinal = input.occurrences.length - 1;
  occurrenceOrdinalByBodyIndex[input.bodyLength] = finalOrdinal;
  const retainedOccurrences = Object.freeze([...input.occurrences]);
  const retainedOrdinals = Object.freeze(occurrenceOrdinalByBodyIndex);
  return Object.freeze({
    occurrences: retainedOccurrences,
    sectionAtBodyIndex(bodyIndex: number): BodySectionOccurrence {
      if (!Number.isInteger(bodyIndex) || bodyIndex < 0 || bodyIndex >= retainedOrdinals.length) {
        throw new RangeError(`Body index ${bodyIndex} is outside the retained section index`);
      }
      return retainedOccurrences[retainedOrdinals[bodyIndex]!]!;
    },
  });
}
