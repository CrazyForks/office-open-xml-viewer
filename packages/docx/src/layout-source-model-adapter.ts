import type { DocParagraph, DocRun, DocxDocumentModel, HeadersFooters } from './types.js';
import { docxRenderedFontFamilies } from './document-content.js';
import { docxFontPreloadNames } from './google-fonts.js';
import { resolveDocumentLayoutSettings } from './layout-context.js';
import { getDefaultFontSize } from './line-layout.js';
import { docDefaultFontSizePt } from './layout/measurement-environment.js';
import { docxLocalMetricRequests } from './local-font-metrics.js';
import {
  normalizeInternalDocumentModel,
  tableSourceAcquisitionInput,
  type InternalShapeRun,
} from './parser-model.js';
import type { BodyAcquisitionInputProjections } from './layout/acquisition-input-projections.js';
import {
  sealLayoutSourceStore,
  type LayoutSourceStore,
  type LayoutParagraphAcquisitionFact,
  type LayoutParagraphBlock,
  type LayoutStoryBlock,
  type LayoutStoryNote,
  type LayoutTableBlock,
  type LayoutTableAcquisitionFact,
} from './layout/layout-source-store.js';
import { paginatedFlowHasPaginationDependentFields } from './layout/pagination-fields.js';
import { deepFreezePlainData } from './layout/plain-data.js';
import { projectDocumentSnapshotResources } from './layout/production-paint-resources.js';
import { sourceKey } from './layout/source-key.js';
import type { BodyElement, DocNote } from './types.js';
import type { SourceRef } from './layout/types.js';
import type { ParagraphLayoutSource } from './layout/text.js';
import { documentRequiresDomVerticalGlyphLayout } from './vertical-render-capability.js';
import { defaultSectionGeometry } from './layout/context.js';

export interface LayoutSourceModelAdapter {
  /** Stable public compatibility model, intentionally outside the source graph. */
  readonly document: DocxDocumentModel;
  readonly source: LayoutSourceStore;
}

const adapters = new WeakMap<object, LayoutSourceModelAdapter>();

function compatibleDocumentModel(input: DocxDocumentModel): DocxDocumentModel {
  const defaults = defaultSectionGeometry();
  const authored = input.section ?? {} as DocxDocumentModel['section'];
  const finiteOrDefault = (value: number | undefined, fallback: number): number => (
    Number.isFinite(value) ? value as number : fallback
  );
  const parts = (value: HeadersFooters | undefined): HeadersFooters => ({
    default: value?.default ?? null,
    first: value?.first ?? null,
    even: value?.even ?? null,
  });
  return {
    ...input,
    body: input.body ?? [],
    section: {
      ...defaults,
      ...authored,
      pageWidth: finiteOrDefault(authored.pageWidth, defaults.pageWidth),
      pageHeight: finiteOrDefault(authored.pageHeight, defaults.pageHeight),
      marginTop: finiteOrDefault(authored.marginTop, defaults.marginTop),
      marginRight: finiteOrDefault(authored.marginRight, defaults.marginRight),
      marginBottom: finiteOrDefault(authored.marginBottom, defaults.marginBottom),
      marginLeft: finiteOrDefault(authored.marginLeft, defaults.marginLeft),
      headerDistance: finiteOrDefault(authored.headerDistance, defaults.headerDistance),
      footerDistance: finiteOrDefault(authored.footerDistance, defaults.footerDistance),
    },
    headers: parts(input.headers),
    footers: parts(input.footers),
  };
}

function forEachPart(
  parts: HeadersFooters | undefined,
  story: 'header' | 'footer',
  instancePrefix: string | null,
  visit: (body: BodyElement[], source: SourceRef) => void,
): void {
  if (!parts) return;
  for (const kind of ['default', 'first', 'even'] as const) {
    const part = parts[kind];
    if (!part) continue;
    visit(part.body, {
      story,
      storyInstance: instancePrefix === null ? kind : `${instancePrefix}:${kind}`,
      path: [],
    });
  }
}

function traverseParagraphSources(
  document: DocxDocumentModel,
  projections: BodyAcquisitionInputProjections,
  retain: (paragraph: DocParagraph, source: SourceRef, input: ReturnType<BodyAcquisitionInputProjections['paragraphAcquisitionInput']>) => void,
  retainTable: (table: Extract<BodyElement, { type: 'table' }>, source: SourceRef) => void = () => {},
  retainStory: (body: BodyElement[], source: SourceRef) => void = () => {},
): void {
  const visitBody = (body: BodyElement[], root: SourceRef, prefix: number[] = []): void => {
    body.forEach((element, elementIndex) => {
      const path = [...prefix, elementIndex];
      if (element.type === 'paragraph') {
        const paragraphSource: SourceRef = { ...root, path };
        const retained = projections.paragraphAcquisitionInput(element, paragraphSource);
        retain(element, paragraphSource, retained);

        // Authored acquisition indices include unavailable drawings omitted from
        // the public run union. Advance public identity only for retained runs so
        // text-box story identities remain aligned with parser acquisition.
        let publicRunIndex = 0;
        retained.runs.forEach((run, authoredRunIndex) => {
          if (run.type === 'unavailableDrawing') return;
          const publicRun = element.runs[publicRunIndex++] as DocRun | undefined;
          if (run.type !== 'shape' || publicRun?.type !== 'shape') return;
          const content = (publicRun as InternalShapeRun).textBoxContent;
          if (!content) return;
          const textBoxRoot: SourceRef = {
            story: 'textbox',
            storyInstance: `${root.story}:${root.storyInstance}:${path.join('.')}.${authoredRunIndex}`,
            path: [],
          };
          retainStory(content as BodyElement[], textBoxRoot);
          visitBody(content as BodyElement[], textBoxRoot);
        });
      } else if (element.type === 'table') {
        retainTable(element, { ...root, path });
        element.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
          visitBody(cell.content as BodyElement[], root, [...path, rowIndex, cellIndex]);
        }));
      } else if (element.type === 'sectionBreak') {
        forEachPart(element.headers, 'header', `section:${elementIndex}`, visitBody);
        forEachPart(element.footers, 'footer', `section:${elementIndex}`, visitBody);
      }
    });
  };
  visitBody(document.body, { story: 'body', storyInstance: 'body', path: [] });
  forEachPart(document.headers, 'header', null, visitBody);
  forEachPart(document.footers, 'footer', null, visitBody);
  for (const note of document.footnotes ?? []) {
    visitBody(note.content, { story: 'footnote', storyInstance: note.id, path: [] });
  }
  for (const note of document.endnotes ?? []) {
    visitBody(note.content, { story: 'endnote', storyInstance: note.id, path: [] });
  }
}

function storyEntries(document: DocxDocumentModel): { source: SourceRef; body: BodyElement[] }[] {
  const entries: { source: SourceRef; body: BodyElement[] }[] = [];
  const add = (body: BodyElement[], source: SourceRef): void => { entries.push({ source, body }); };
  forEachPart(document.headers, 'header', null, add);
  forEachPart(document.footers, 'footer', null, add);
  document.body.forEach((element, elementIndex) => {
    if (element.type !== 'sectionBreak') return;
    forEachPart(element.headers, 'header', `section:${elementIndex}`, add);
    forEachPart(element.footers, 'footer', `section:${elementIndex}`, add);
  });
  return entries;
}

function storyBodies(document: DocxDocumentModel): BodyElement[][] {
  return storyEntries(document).map(({ body }) => body);
}

function frozenNotes(notes: DocNote[] | undefined): readonly DocNote[] {
  return notes ?? Object.freeze([]);
}

function canonicalLayoutBody(
  body: readonly BodyElement[],
  root: SourceRef,
  paragraphInputs: ReadonlyMap<string, ReturnType<BodyAcquisitionInputProjections['paragraphAcquisitionInput']>>,
  prefix: number[] = [],
): LayoutStoryBlock[] {
  const canonicalParts = (
    parts: HeadersFooters | undefined,
    story: 'header' | 'footer',
    instancePrefix: string,
  ): HeadersFooters | undefined => {
    if (!parts) return parts;
    return Object.fromEntries((['default', 'first', 'even'] as const).map((kind) => {
      const part = parts[kind];
      return [kind, !part ? null : {
        ...structuredClone(Object.fromEntries(
          Object.entries(part).filter(([key]) => key !== 'body'),
        )),
        body: canonicalLayoutBody(part.body, {
          story, storyInstance: `${instancePrefix}:${kind}`, path: [],
        }, paragraphInputs),
      }];
    })) as unknown as HeadersFooters;
  };
  return body.map((element, elementIndex): LayoutStoryBlock => {
    const path = [...prefix, elementIndex];
    if (element.type === 'paragraph') {
      const retained = paragraphInputs.get(sourceKey({ ...root, path }));
      if (!retained) throw new Error(`Missing canonical paragraph source: ${sourceKey({ ...root, path })}`);
      return retained as LayoutParagraphBlock;
    }
    if (element.type === 'table') {
      // Detach one bounded logical table at a time. This replaces the former
      // whole-document structuredClone while preserving the store's strict
      // non-aliasing contract with the mutable compatibility model.
      const detached = structuredClone(element);
      const { __tableLayout: _tableLayout, ...table } = detached as typeof detached & { __tableLayout?: unknown };
      return {
      ...table,
      rows: detached.rows.map((row, rowIndex) => {
        const { __tableRowLayout: _rowLayout, ...retainedRow } = row as typeof row & { __tableRowLayout?: unknown };
        return {
        ...retainedRow,
        cells: row.cells.map((cell, cellIndex) => {
          const { __tableCellLayout: _cellLayout, ...retainedCell } = cell as typeof cell & { __tableCellLayout?: unknown };
          return {
          ...retainedCell,
          content: canonicalLayoutBody(
            cell.content as BodyElement[], root, paragraphInputs,
            [...path, rowIndex, cellIndex],
          ),
        };
        }),
      };
      }),
    } as unknown as LayoutTableBlock;
    }
    if (element.type !== 'sectionBreak') return structuredClone(element);
    const {
      __sectionPlacement: _sectionPlacement,
      headers: _headers,
      footers: _footers,
      ...sectionBreak
    } = element as typeof element & { __sectionPlacement?: unknown };
    return {
      ...structuredClone(sectionBreak),
      headers: canonicalParts(element.headers, 'header', `section:${elementIndex}`),
      footers: canonicalParts(element.footers, 'footer', `section:${elementIndex}`),
    } as unknown as LayoutStoryBlock;
  });
}

function canonicalSection(
  section: DocxDocumentModel['section'],
): DocxDocumentModel['section'] {
  const { __sectionPlacement: _sectionPlacement, ...retained } = section as typeof section & {
    __sectionPlacement?: unknown;
  };
  return structuredClone(retained);
}

function canonicalFinalParts(
  parts: HeadersFooters,
  story: 'header' | 'footer',
  paragraphInputs: ReadonlyMap<string, ReturnType<BodyAcquisitionInputProjections['paragraphAcquisitionInput']>>,
): HeadersFooters {
  return Object.fromEntries((['default', 'first', 'even'] as const).map((kind) => {
    const part = parts[kind];
    return [kind, !part ? null : {
      ...structuredClone(Object.fromEntries(
        Object.entries(part).filter(([key]) => key !== 'body'),
      )),
      body: canonicalLayoutBody(part.body, { story, storyInstance: kind, path: [] }, paragraphInputs),
    }];
  })) as unknown as HeadersFooters;
}

/**
 * The only full-model adapter. It snapshots parser facts once, captures
 * identity-only sidecars by SourceRef, and seals a model-free source store.
 */
export function layoutSourceModelAdapter(input: DocxDocumentModel): LayoutSourceModelAdapter {
  const cached = adapters.get(input);
  if (cached) return cached;

  const publicInput = normalizeInternalDocumentModel(compatibleDocumentModel(input));
  // Acquire while parser sidecars are still attached. A structured clone cannot
  // carry WeakMap-owned unavailable-drawing facts, so deriving this projection
  // from the detached compatibility clone would misclassify such paragraphs.
  const bodyLayoutInput = publicInput.bodyLayoutInput;
  const paragraphInputs = new Map<string, ReturnType<BodyAcquisitionInputProjections['paragraphAcquisitionInput']>>();
  const paragraphFacts: LayoutParagraphAcquisitionFact[] = [];
  traverseParagraphSources(publicInput.document, publicInput.bodyModelGateway.acquisitionInputs, (paragraph, source, retained) => {
    const key = sourceKey(source);
    if (paragraphInputs.has(key)) throw new Error(`Duplicate paragraph source: ${key}`);
    paragraphInputs.set(key, retained);
    let publicRunIndex = 0;
    const publicAnchorBridges = retained.runs.map((run, authoredRunIndex) => {
      if (run.type === 'unavailableDrawing') return null;
      const publicRun = paragraph.runs[publicRunIndex++];
      return publicRun
        ? publicInput.bodyModelGateway.publicAnchorBridge(publicRun, source, authoredRunIndex)
        : null;
    });
    paragraphFacts.push(Object.freeze({
      source: deepFreezePlainData({ ...source, path: [...source.path] }),
      publicAnchorBridges: Object.freeze(publicAnchorBridges),
      numberingMarkerFallbackFontSizePt: paragraph.numbering
        ? getDefaultFontSize(paragraph)
        : null,
    }));
  });

  // Paragraph acquisition already snapshots one logical paragraph at a time;
  // table/story canonicalization below does the same for its bounded unit. Do
  // not clone the complete compatibility model here: that transiently doubled
  // every body, story, and resource graph immediately before sealing.
  const privateInput = publicInput;
  const privateDocument = privateInput.document;
  const privateProjections = privateInput.bodyModelGateway.acquisitionInputs;
  const projections: BodyAcquisitionInputProjections = Object.freeze({
    ...privateProjections,
    paragraphAcquisitionInput(_paragraph: ParagraphLayoutSource, source: SourceRef) {
      const retained = paragraphInputs.get(sourceKey(source));
      if (!retained) throw new Error(`Unknown paragraph acquisition source: ${sourceKey(source)}`);
      return retained;
    },
  });
  const tableFacts: LayoutTableAcquisitionFact[] = [];
  const textBoxStories: { source: SourceRef; body: BodyElement[] }[] = [];
  traverseParagraphSources(privateDocument, projections, () => {}, (table, source) => {
    tableFacts.push(Object.freeze({
      source: deepFreezePlainData({ ...source, path: [...source.path] }),
      input: tableSourceAcquisitionInput(table),
    }));
  }, (body, source) => { textBoxStories.push({ body, source }); });
  const documentLayoutSettings = resolveDocumentLayoutSettings(privateDocument);
  const resources = projectDocumentSnapshotResources(
    privateDocument,
    projections,
    privateInput.mathOccurrences,
    (paragraph) => {
    const numbering = paragraph.numbering;
    if (!numbering) throw new Error('Picture-bullet metadata requires numbering');
    const marker = projections.numberingMarkerShapeInput(numbering, getDefaultFontSize(paragraph));
    return {
      widthPt: numbering.picBulletWidthPt ?? marker.fontSizePt,
      heightPt: numbering.picBulletHeightPt ?? marker.fontSizePt,
    };
  });
  const footnotes = frozenNotes(privateDocument.footnotes);
  const endnotes = frozenNotes(privateDocument.endnotes);
  const hasPaginationFields = paginatedFlowHasPaginationDependentFields(
    privateDocument.body,
    footnotes,
    [...storyBodies(privateDocument), ...endnotes.map((note) => note.content)],
  );
  const requiresDomVerticalGlyphLayout = documentRequiresDomVerticalGlyphLayout(privateDocument);
  const fatalParse = privateDocument.parseError === undefined ? null : {
    message: privateDocument.parseError,
    pageSize: {
      widthPt: privateDocument.section.pageWidth,
      heightPt: privateDocument.section.pageHeight,
    },
  };
  const fonts = {
    familyClasses: { ...(privateDocument.fontFamilyClasses ?? {}) },
    familyPitches: { ...(privateDocument.fontFamilyPitches ?? {}) },
    majorFamily: privateDocument.majorFont ?? null,
    minorFamily: privateDocument.minorFont ?? null,
    embeddedFonts: [...(privateDocument.embeddedFonts ?? [])],
    renderedFamilies: docxRenderedFontFamilies(privateDocument),
    preloadNames: docxFontPreloadNames(privateDocument),
    localMetricRequests: docxLocalMetricRequests(privateDocument),
    defaultBodyFontSizePt: docDefaultFontSizePt(privateDocument),
  };

  const canonicalDocument = {
    ...privateDocument,
    body: canonicalLayoutBody(
      privateDocument.body,
      { story: 'body', storyInstance: 'body', path: [] },
      paragraphInputs,
    ),
    headers: canonicalFinalParts(privateDocument.headers, 'header', paragraphInputs),
    footers: canonicalFinalParts(privateDocument.footers, 'footer', paragraphInputs),
    footnotes: footnotes.map(({ content, ...note }) => ({
      ...structuredClone(note),
      content: canonicalLayoutBody(content, {
        story: 'footnote', storyInstance: note.id, path: [],
      }, paragraphInputs),
    })),
    endnotes: endnotes.map(({ content, ...note }) => ({
      ...structuredClone(note),
      content: canonicalLayoutBody(content, {
        story: 'endnote', storyInstance: note.id, path: [],
      }, paragraphInputs),
    })),
  } as unknown as DocxDocumentModel;
  const canonicalTextBoxStories = textBoxStories.map(({ source, body }) => ({
    source,
    body: canonicalLayoutBody(body, source, paragraphInputs),
  }));

  deepFreezePlainData(paragraphInputs as unknown as object);
  // Project the body input before sealing the store. The source store must own
  // only the resulting plain snapshot; retaining a lazy closure here would keep
  // the complete private parser model reachable for the store lifetime.
  const source = sealLayoutSourceStore({
    bodyLayoutInput,
    blockRepository: {
      body: canonicalDocument.body as unknown as readonly LayoutStoryBlock[],
      stories: [
        ...storyEntries(canonicalDocument).map(({ source, body }) => ({
          source,
          body: body as unknown as readonly LayoutStoryBlock[],
        })),
        ...canonicalTextBoxStories,
      ],
      footnotes: (canonicalDocument.footnotes ?? []) as unknown as readonly LayoutStoryNote[],
      endnotes: (canonicalDocument.endnotes ?? []) as unknown as readonly LayoutStoryNote[],
    },
    section: canonicalSection(privateDocument.section),
    documentLayoutFacts: deepFreezePlainData({
      ...documentLayoutSettings,
      kinsoku: {
        enabled: documentLayoutSettings.kinsoku.enabled,
        lineStartForbidden: [...documentLayoutSettings.kinsoku.lineStartForbidden].sort((a, b) => a - b),
        lineEndForbidden: [...documentLayoutSettings.kinsoku.lineEndForbidden].sort((a, b) => a - b),
      },
    }),
    fonts: deepFreezePlainData(fonts),
    fontFamilyCharsets: privateInput.fontFamilyCharsets,
    acquisitionFacts: Object.freeze({
      paragraphs: Object.freeze(paragraphFacts),
      tables: Object.freeze(tableFacts),
    }),
    mathOccurrences: privateInput.mathOccurrences,
    imageMetadata: resources.imageMetadata,
    paintDescriptors: resources.paintResources.descriptors,
    hasPaginationFields,
    requiresDomVerticalGlyphLayout,
    fatalParse: fatalParse === null ? null : deepFreezePlainData(fatalParse),
  });
  const adapter = Object.freeze({ document: publicInput.document, source });
  adapters.set(input, adapter);
  adapters.set(publicInput.document, adapter);
  return adapter;
}

export function layoutSourceStore(input: DocxDocumentModel): LayoutSourceStore {
  return layoutSourceModelAdapter(input).source;
}
