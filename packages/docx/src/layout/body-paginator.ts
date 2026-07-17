import type { SectionLayoutContext } from '../layout-context.js';
import type {
  BodyLayoutInput,
  BodyParagraphSourceInput,
  BodySectionLayoutInput,
  BodyTableSourceInput,
} from './body-layout-input.js';
import type {
  BodyAcquisitionLocation,
  BodyLayoutSession,
  BodyTableContinuationCursor,
} from './body-layout-kernel.js';
import {
  commitPageFlowTransition,
  createBodyPaginationState,
  createCanonicalPageDraft,
  addPageFootnoteReserve,
  markBodySourceConsumed,
  setBodyBalanceTarget,
  type BodyPageTransitionFactory,
  type BodyPaginationState,
  type CanonicalPageDraft,
} from './body-pagination.js';
import { deepFreezeDocumentLayout, assertDocumentLayout } from './invariants.js';
import {
  bodyLayoutKernelOf,
  createFieldAcquisitionServicesView,
} from './runtime-state.js';
import {
  accumulatePagePaintNode,
  accumulatePageSectionRegion,
  bodyFlowDomainId,
  createParityBlankLayoutPage,
  finalizeLayoutPage,
  type PageSectionRegionInput,
} from './page-factory.js';
import {
  projectBodyOccurrence,
  projectedNestedOccurrenceId,
} from './occurrence-projection.js';
import {
  advanceColumnOrPage,
  applyAuthoredBreak,
  beginSection,
  createPageFlowState,
  placeFlowNode,
  type PageFlowState,
} from './paginator.js';
import {
  createPageFlowSectionContext,
  resolveSectionContextForPage,
} from './context.js';
import { uprightPhysicalExtent, writingModeFromTextDirection } from './coordinate-space.js';
import { selectParagraphFragment, type ParagraphFragmentCursor } from './paragraph-pagination.js';
import { paragraphGapAdjustment } from './paragraph-spacing.js';
import { footnoteIdsInRetainedSlice } from './note-reference-ownership.js';
import { composeCanonicalSectionFlow } from './section-flow-composition.js';
import type { BodyFlowAllocation } from './section-flow-composition.js';
import {
  wordActiveColumnBreakIndexes,
  wordFlowNeutralPreBreakAnchorParagraph,
  wordPreBreakHostAnchorExtentPt,
  wordPreBreakInlineDrawingResource,
} from './body-pagination-compatibility.js';
import { bodyOccurrenceKey } from './source-key.js';
import {
  convergeHeaderFooterReserves,
  headerFooterOverflowReservePt,
  reservedBodyInterval,
  selectedHeaderFooterStory,
  type HeaderFooterReserve,
  type ReservedBodyInterval,
} from './header-footer-reserve.js';
import type { LayoutOptions } from './options.js';
import type {
  DeepReadonly,
  DocumentLayout,
  LayoutServices,
  PaintNode,
  ParagraphLayout,
  SourceRef,
  TableLayout,
  FloatRegistryDeltaPt,
} from './types.js';

function nestedFloatingOccurrenceIds(layout: TableLayout): ReadonlySet<string> {
  const ids = new Set<string>();
  const visit = (table: TableLayout) => {
    for (const resolved of table.resolvedFloatingTables ?? []) {
      ids.add(resolved.occurrenceId);
      visit(resolved.child);
    }
  };
  visit(layout);
  return ids;
}

function bindTableFloatDeltaToAcceptedOccurrence(
  delta: FloatRegistryDeltaPt,
  layout: TableLayout,
  ownerOccurrenceId: string,
): FloatRegistryDeltaPt {
  const nestedIds = nestedFloatingOccurrenceIds(layout);
  return Object.freeze({
    ...delta,
    entries: Object.freeze(delta.entries.map((entry) => {
      const occurrenceId = nestedIds.has(entry.occurrenceId)
        ? projectedNestedOccurrenceId(ownerOccurrenceId, entry.occurrenceId)
        : layout.ordinaryFlow
          ? null
          : ownerOccurrenceId;
      if (occurrenceId === null) return entry;
      return Object.freeze({ ...entry, occurrenceId, exclusionId: occurrenceId });
    })),
  });
}

function paragraphFloatDeltaForAcceptedFragment(
  delta: FloatRegistryDeltaPt,
  fragment: ParagraphLayout,
): FloatRegistryDeltaPt | null {
  const acceptedAnchorOccurrenceIds = new Set(fragment.drawings.flatMap((drawing) => {
    const occurrenceId = drawing.anchorLayer?.acquisitionOccurrenceId
      ?? drawing.anchorLayer?.occurrenceId;
    return occurrenceId === undefined ? [] : [occurrenceId];
  }));
  const entries = delta.entries.filter((entry) =>
    acceptedAnchorOccurrenceIds.has(entry.occurrenceId));
  if (entries.length === 0) return null;
  return Object.freeze({
    ...delta,
    entries: Object.freeze(entries),
    nextParagraphId: delta.baseNextParagraphId + entries.length,
  });
}

function ownerMap(input: BodyLayoutInput): Map<string, BodySectionLayoutInput> {
  const owners = new Map([[input.initialSection.sectionOccurrenceId, input.initialSection]]);
  for (let entryIndex = 0; entryIndex < input.sequence.length; entryIndex += 1) {
    const entry = input.sequence[entryIndex]!;
    if (entry.kind === 'begin-section') owners.set(entry.section.sectionOccurrenceId, entry.section);
  }
  return owners;
}

function sectionContextForPage(owner: BodySectionLayoutInput, pageIndex: number) {
  return resolveSectionContextForPage(owner.context as SectionLayoutContext, owner.pageLayout, pageIndex);
}

function flowSection(owner: BodySectionLayoutInput, pageIndex: number) {
  const context = sectionContextForPage(owner, pageIndex);
  return createPageFlowSectionContext({
    sectionOccurrenceId: owner.sectionOccurrenceId,
    geometry: context.geometry,
    columns: context.columns,
    textDirection: context.textDirection,
  });
}

function pageRegion(
  owner: BodySectionLayoutInput,
  pageIndex: number,
  interval: ReservedBodyInterval,
  blockStartPt = interval.blockStartPt,
): PageSectionRegionInput {
  const context = sectionContextForPage(owner, pageIndex);
  return Object.freeze({
    id: `page:${pageIndex}:section:${encodeURIComponent(owner.sectionOccurrenceId)}`,
    sectionOccurrenceId: owner.sectionOccurrenceId,
    section: context,
    pageBorders: owner.pageBordersAuthored ? owner.pageBorders : null,
    writingMode: writingModeFromTextDirection(context.textDirection),
    blockStartPt,
    blockEndPt: interval.blockEndPt,
    columns: Object.freeze(context.columns.map((column) => Object.freeze({
      inlineStartPt: column.xPt,
      inlineExtentPt: column.wPt,
    }))),
  });
}

function physicalPage(
  section: DeepReadonly<SectionLayoutContext>,
  interval: ReservedBodyInterval,
) {
  const writingMode = writingModeFromTextDirection(section.textDirection);
  const extent = uprightPhysicalExtent({
    widthPt: section.geometry.pageWidth,
    heightPt: section.geometry.pageHeight,
  }, writingMode);
  if (writingMode !== 'horizontal-tb') {
    return Object.freeze({
      ...extent,
      contentTopPt: 0,
      contentBottomPt: extent.heightPt,
    });
  }
  return Object.freeze({
    ...extent,
    contentTopPt: interval.blockStartPt,
    contentBottomPt: interval.blockEndPt,
  });
}

function pageBodyInterval(
  owner: BodySectionLayoutInput,
  pageIndex: number,
  reserve: HeaderFooterReserve,
): ReservedBodyInterval {
  return reservedBodyInterval(sectionContextForPage(owner, pageIndex).geometry, reserve);
}

function openDraft(
  owner: BodySectionLayoutInput,
  pageIndex: number,
  interval: ReservedBodyInterval,
): CanonicalPageDraft {
  const context = sectionContextForPage(owner, pageIndex);
  return createCanonicalPageDraft({
    kind: 'content', pageIndex,
    physicalPage: physicalPage(context, interval),
    sectionOccurrenceId: owner.sectionOccurrenceId,
    section: context,
    region: pageRegion(owner, pageIndex, interval),
  });
}

function activeRegion(state: BodyPaginationState): PageSectionRegionInput {
  const page = state.pages.at(-1);
  const region = page?.accumulator.sectionRegions.at(-1);
  if (!page || page.kind !== 'content' || !region) throw new Error('Missing active body region');
  return region;
}

function activeBlockEndPt(state: BodyPaginationState): number {
  const region = activeRegion(state);
  const isLastColumn = state.flow.columnIndex >= region.columns.length - 1;
  return state.balanceTargetPt === null || isLastColumn
    ? region.blockEndPt
    : Math.min(region.blockEndPt, region.blockStartPt + state.balanceTargetPt);
}

function acquisitionLocation(state: BodyPaginationState): BodyAcquisitionLocation {
  const region = activeRegion(state);
  const column = region.columns[state.flow.columnIndex];
  if (!column) throw new Error('Missing active body column');
  return Object.freeze({
    pageIndex: state.flow.pageIndex,
    columnIndex: state.flow.columnIndex,
    flowDomainId: bodyFlowDomainId(state.flow.pageIndex, region.id, state.flow.columnIndex),
    section: region.section,
    cursorPt: Object.freeze({ xPt: column.inlineStartPt, yPt: state.flow.cursorBlockPt }),
    availableBounds: Object.freeze({
      xPt: column.inlineStartPt,
      yPt: state.flow.cursorBlockPt,
      widthPt: column.inlineExtentPt,
      heightPt: Math.max(
        0,
        activeBlockEndPt(state) - state.footnoteReservePt - state.flow.cursorBlockPt,
      ),
    }),
  });
}

function transitionFactory(
  owners: ReadonlyMap<string, BodySectionLayoutInput>,
  reserves: readonly HeaderFooterReserve[],
): BodyPageTransitionFactory {
  const owner = (id: string) => {
    const result = owners.get(id);
    if (!result) throw new Error(`Unknown body section ${id}`);
    return result;
  };
  return {
    openContentPage(event) {
      const nextOwner = owner(event.sectionOccurrenceId);
      const reserve = reserves[event.pageIndex] ?? { top: 0, bottom: 0 };
      const interval = pageBodyInterval(nextOwner, event.pageIndex, reserve);
      const flow = createPageFlowState(flowSection(nextOwner, event.pageIndex), {
        pageIndex: event.pageIndex, pageContentStartBlockPt: interval.blockStartPt,
      });
      return { page: openDraft(nextOwner, event.pageIndex, interval), flow };
    },
    openParityBlankPage(event) {
      const blankOwner = owner(event.sectionOccurrenceId);
      const context = sectionContextForPage(blankOwner, event.pageIndex);
      const interval = pageBodyInterval(
        blankOwner,
        event.pageIndex,
        reserves[event.pageIndex] ?? { top: 0, bottom: 0 },
      );
      return createCanonicalPageDraft({
        kind: 'parity-blank', pageIndex: event.pageIndex,
        physicalPage: physicalPage(context, interval),
        sectionOccurrenceId: blankOwner.sectionOccurrenceId,
        section: context,
        pageBorders: blankOwner.pageBordersAuthored ? blankOwner.pageBorders : null,
      });
    },
    openContinuousSectionRegion(page, event, flow) {
      const nextOwner = owner(event.section.sectionOccurrenceId);
      const priorRegions = page.accumulator.sectionRegions;
      const prior = priorRegions.at(-1);
      if (!prior) throw new Error('A continuous section requires a prior region');
      const pageInterval = Object.freeze({
        blockStartPt: page.accumulator.sectionRegions[0]!.blockStartPt,
        blockEndPt: prior.blockEndPt,
      });
      const constrainedRegions = Object.freeze([
        ...priorRegions.slice(0, -1),
        Object.freeze({ ...prior, blockEndPt: flow.regionStartBlockPt }),
      ]);
      return Object.freeze({
        ...page,
        accumulator: accumulatePageSectionRegion(
          Object.freeze({ ...page.accumulator, sectionRegions: constrainedRegions }),
          pageRegion(
            nextOwner,
            flow.pageIndex,
            pageInterval,
            flow.regionStartBlockPt,
          ),
        ),
      });
    },
  };
}

function acceptNode(
  state: BodyPaginationState,
  retained: ParagraphLayout | TableLayout,
  source: SourceRef,
  blockExtentPt: number,
  fragmentStartKey: string,
  acceptedOccurrenceIds: Set<string>,
  allocations: BodyFlowAllocation[],
  placement?: Readonly<{
    coordinateSpace: 'logical-body' | 'upright-physical';
    xPt: number;
    yPt: number;
    sectionFlowOwnership?: 'host-flow' | 'page';
  }>,
): BodyPaginationState {
  const page = state.pages.at(-1);
  if (!page || page.kind !== 'content') throw new Error('Body content requires an active page');
  const region = activeRegion(state);
  const column = region.columns[state.flow.columnIndex]!;
  const flowDomainId = bodyFlowDomainId(state.flow.pageIndex, region.id, state.flow.columnIndex);
  const occurrenceId = bodyOccurrenceKey(source, flowDomainId, fragmentStartKey);
  if (acceptedOccurrenceIds.has(occurrenceId)) {
    throw new Error(`Duplicate body occurrence acceptance: ${occurrenceId}`);
  }
  acceptedOccurrenceIds.add(occurrenceId);
  const projected = projectBodyOccurrence(retained, {
    occurrenceId,
    destination: {
      coordinateSpace: 'logical-page-points',
      flowDomainId,
      translation: {
        // Ordinary table acquisition owns the complete inline placement:
        // physical jc alignment followed by signed tblInd translation. Move
        // that acquisition-local frame into the page column without
        // normalizing its retained X origin away. Explicit out-of-flow
        // placements and paragraphs continue to own an exact destination.
        xPt: placement
          ? placement.xPt - retained.flowBounds.xPt
          : retained.kind === 'table'
            ? column.inlineStartPt
            : column.inlineStartPt - retained.flowBounds.xPt,
        yPt: (placement?.yPt ?? state.flow.cursorBlockPt) - retained.flowBounds.yPt,
      },
    },
  });
  const ownershipRetained = placement?.sectionFlowOwnership === undefined
    ? projected
    : Object.freeze({ ...projected, sectionFlowOwnership: placement.sectionFlowOwnership });
  const contentOwned = ownershipRetained.kind === 'paragraph' && ownershipRetained.ordinaryFlow
    ? Object.freeze({
        ...ownershipRetained,
        flowBounds: Object.freeze({
          ...ownershipRetained.flowBounds,
          yPt: ownershipRetained.flowBounds.yPt + ownershipRetained.spacing.beforePt,
          // Admission owns flow containment; intrinsic line and ink geometry stay retained separately.
          heightPt: Math.max(
            0,
            blockExtentPt
              - ownershipRetained.spacing.beforePt
              - ownershipRetained.spacing.afterPt,
          ),
        }),
      })
    : ownershipRetained;
  const retainedAtDestination = placement?.coordinateSpace === 'upright-physical'
    ? ({
        ...contentOwned,
        ordinaryFlow: false,
        flowBounds: Object.freeze({
          ...contentOwned.flowBounds,
          heightPt: blockExtentPt,
        }),
      } as typeof contentOwned)
    : contentOwned;
  const transition = placeFlowNode(state.flow, retainedAtDestination, blockExtentPt);
  const place = transition.events[0];
  if (!place || place.type !== 'place') throw new Error('Flow placement did not emit an allocation');
  allocations.push(Object.freeze({
    nodeId: retainedAtDestination.id,
    flowDomainId: retainedAtDestination.flowDomainId,
    blockStartPt: place.blockStartPt,
    blockEndPt: place.blockEndPt,
  }));
  const accumulator = accumulatePagePaintNode(page.accumulator, {
    layer: 'body', node: retainedAtDestination,
    ...(placement?.coordinateSpace === 'upright-physical'
      ? { coordinateSpace: 'upright-physical' as const }
      : {}),
  }, true);
  const pages = [...state.pages];
  pages[pages.length - 1] = Object.freeze({ ...page, accumulator });
  return Object.freeze({
    ...state,
    flow: transition.state,
    pages: Object.freeze(pages),
    pageHasConsumedSource: true,
  });
}

function paragraphFragmentStartKey(cursor: ParagraphFragmentCursor): string {
  return cursor.boundary === null
    ? 'root'
    : `paragraph:${cursor.boundary.segIndex}:${cursor.boundary.charOffset}`;
}

function hasFollowingInkContent(input: BodyLayoutInput, startIndex: number): boolean {
  for (let index = startIndex; index < input.sequence.length; index += 1) {
    const entry = input.sequence[index]!;
    if (entry.kind === 'consume-source') continue;
    if (entry.kind === 'authored-break') {
      if (entry.break !== 'lastRenderedPageBreak') return false;
      continue;
    }
    if (entry.kind === 'begin-section') {
      if (entry.section.startType !== 'continuous') return false;
      continue;
    }
    const block = entry.kind === 'adjacent-table-group' ? entry : entry.block;
    if (block.kind !== 'paragraph') return true;
    if (block.pageBreakBefore) return false;
    if (block.inkless !== true) return true;
  }
  return false;
}

function isUndecoratedInklessMark(layout: ParagraphLayout): boolean {
  return layout.paragraphMark !== undefined
    && layout.lines.length === 0
    && layout.shading === undefined
    && layout.borders.length === 0
    && layout.resources.length === 0
    && layout.drawings.length === 0
    && layout.textBoxes.length === 0;
}

function tableCursorKey(cursor: import('./table-pagination.js').TableFragmentCursor): readonly unknown[] {
  return [
    cursor.rowIndex,
    cursor.rowFragmentIndex,
    cursor.cells.map((cell) => [
      cell.blockIndex,
      cell.paragraphLineStart,
      cell.nestedFragmentIndex,
      cell.nestedCursor === null ? null : tableCursorKey(cell.nestedCursor),
    ]),
  ];
}

function tableFragmentStartKey(cursor: BodyTableContinuationCursor | undefined): string {
  if (cursor === undefined) return 'root';
  if (cursor.kind === 'table') return `table:${JSON.stringify(tableCursorKey(cursor.cursor))}`;
  const tableCursor = cursor.cursor.tableCursor;
  return `adjacent-table:${cursor.cursor.tableIndex}:${cursor.cursor.sourceRowIndex}:${JSON.stringify(
    tableCursor === undefined ? null : tableCursorKey(tableCursor),
  )}`;
}

function freshPageExtent(state: BodyPaginationState): number {
  const region = activeRegion(state);
  return activeBlockEndPt(state) - region.blockStartPt;
}

function locationAfter(
  location: BodyAcquisitionLocation,
  blockExtentPt: number,
): BodyAcquisitionLocation {
  const yPt = location.cursorPt.yPt + blockExtentPt;
  const blockEndPt = location.availableBounds.yPt + location.availableBounds.heightPt;
  return Object.freeze({
    ...location,
    cursorPt: Object.freeze({ ...location.cursorPt, yPt }),
    availableBounds: Object.freeze({
      ...location.availableBounds,
      yPt,
      heightPt: Math.max(0, blockEndPt - yPt),
    }),
  });
}

function finalize(state: BodyPaginationState, owners: ReadonlyMap<string, BodySectionLayoutInput>): DocumentLayout {
  const firstAppearance = new Map<string, number>();
  state.pages.forEach((draft) => draft.accumulator.sectionRegions.forEach((region) => {
    if (!firstAppearance.has(region.sectionOccurrenceId)) {
      firstAppearance.set(region.sectionOccurrenceId, draft.accumulator.pageIndex);
    }
  }));
  let displayNumber = 0;
  let priorOwner: string | null = null;
  const pages = state.pages.map((draft) => {
    const owner = owners.get(draft.accumulator.sectionOccurrenceId)!;
    if (owner.sectionOccurrenceId !== priorOwner && owner.pageNumbering.start !== null) {
      displayNumber = owner.pageNumbering.start + (
        draft.accumulator.pageIndex
          - (firstAppearance.get(owner.sectionOccurrenceId) ?? draft.accumulator.pageIndex)
      );
    } else displayNumber += 1;
    priorOwner = owner.sectionOccurrenceId;
    const pageNumber = {
      displayNumber,
      format: owner.pageNumbering.format ?? 'decimal',
      sectionOccurrenceId: owner.sectionOccurrenceId,
    };
    if (draft.kind === 'parity-blank') {
      return createParityBlankLayoutPage({
        pageIndex: draft.accumulator.pageIndex,
        physicalPage: draft.accumulator.physicalPage,
        sectionOccurrenceId: draft.accumulator.sectionOccurrenceId,
        section: draft.accumulator.section,
        pageBorders: draft.accumulator.pageBorders,
        pageNumber,
      });
    }
    return finalizeLayoutPage(draft.accumulator, pageNumber);
  });
  const layout: DocumentLayout = { pages, diagnostics: [] };
  assertDocumentLayout(layout);
  return deepFreezeDocumentLayout(layout) as DocumentLayout;
}

function paginateBodyPass(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
  reserves: readonly HeaderFooterReserve[],
  anchorDestinations: ReadonlyMap<string, Readonly<{
    occurrenceId: string;
    paragraphSource: SourceRef;
    pageIndex: number;
    flowDomainId: string;
  }>> | null,
): Readonly<{
  layout: DocumentLayout;
  session: BodyLayoutSession;
  allocations: readonly BodyFlowAllocation[];
}> {
  const kernel = bodyLayoutKernelOf(services);
  if (!kernel) throw new Error('Body layout kernel is not attached to the supplied services');
  const owners = ownerMap(input);
  const acceptedOccurrenceIds = new Set<string>();
  const allocations: BodyFlowAllocation[] = [];
  const initialReserve = reserves[0] ?? { top: 0, bottom: 0 };
  const initialInterval = pageBodyInterval(input.initialSection, 0, initialReserve);
  const initialFlow = createPageFlowState(flowSection(input.initialSection, 0), {
    pageContentStartBlockPt: initialInterval.blockStartPt,
  });
  let state = createBodyPaginationState(
    initialFlow,
    openDraft(input.initialSection, 0, initialInterval),
  );
  const factory = transitionFactory(owners, reserves);
  const session = kernel.openBodyLayoutSession({
    source: input.source,
    section: input.initialSection.context,
    initialLocation: acquisitionLocation(state),
  }, services, options);
  const pageStartAnchors = (target: BodyPaginationState, startIndex: number) => {
    if (anchorDestinations !== null) {
      const location = acquisitionLocation(target);
      return Object.freeze([...anchorDestinations.values()]
        .filter((destination) => (
          destination.pageIndex === location.pageIndex
          && destination.flowDomainId === location.flowDomainId
        ))
        .map(({ occurrenceId, paragraphSource }) => Object.freeze({
          occurrenceId,
          paragraphSource,
        })));
    }
    const anchors: Array<Readonly<{ occurrenceId: string; paragraphSource: SourceRef }>> = [];
    for (let index = startIndex; index < input.sequence.length; index += 1) {
      const entry = input.sequence[index]!;
      if (entry.kind === 'authored-break' && entry.break !== 'column') break;
      if (entry.kind === 'begin-section' && entry.section.startType !== 'continuous') break;
      if (entry.kind !== 'body-block' || entry.block.kind !== 'paragraph') continue;
      if (index > startIndex && entry.block.pageBreakBefore) break;
      entry.block.pageOwnedAnchorOccurrenceIds?.forEach((occurrenceId) => anchors.push(
        Object.freeze({ occurrenceId, paragraphSource: entry.block.source }),
      ));
    }
    return Object.freeze(anchors);
  };
  const prescanPageAnchors = (target: BodyPaginationState, startIndex: number) => {
    const anchors = pageStartAnchors(target, startIndex);
    if (anchors.length === 0) return;
    if (!session.prescanPageAnchors) {
      throw new Error('Page-owned anchors require canonical prescan acquisition');
    }
    const location = acquisitionLocation(target);
    const delta = session.prescanPageAnchors({
      anchors,
      location,
      availableInlineExtentPt: location.availableBounds.widthPt,
    });
    if (delta) session.commitFloatRegistryDelta(delta);
  };
  prescanPageAnchors(state, 0);
  const approximateBalanceTarget = (
    startIndex: number,
    owner: BodySectionLayoutInput,
  ): number | null => {
    const region = activeRegion(state);
    if (owner.startType !== 'continuous' || region.columns.length < 2) return null;
    const location = acquisitionLocation(state);
    const physicalExtentPt = region.blockEndPt - region.blockStartPt;
    let totalExtentPt = 0;
    let priorParagraph: BodyParagraphSourceInput | null = null;
    let terminated = false;
    for (let index = startIndex; index < input.sequence.length; index += 1) {
      const entry = input.sequence[index]!;
      if (entry.kind === 'begin-section') {
        terminated = true;
        break;
      }
      if (entry.kind === 'consume-source') continue;
      if (entry.kind === 'authored-break') {
        if (entry.break !== 'column') return null;
        priorParagraph = null;
        continue;
      }
      const candidate = entry.kind === 'adjacent-table-group' ? entry : entry.block;
      if (candidate.kind === 'paragraph' && candidate.pageBreakBefore) return null;
      const measured = session.measureFollowingBlock({
        input: candidate,
        location,
        availableInlineExtentPt: location.availableBounds.widthPt,
      });
      if (candidate.kind === 'paragraph') {
        const spacing = paragraphGapAdjustment(
          priorParagraph,
          candidate,
          priorParagraph?.spaceAfterPt ?? 0,
          candidate.spaceBeforePt,
        );
        totalExtentPt += measured.fullExtentPt - spacing.overlap;
        priorParagraph = candidate;
      } else {
        totalExtentPt += measured.fullExtentPt;
        priorParagraph = null;
      }
    }
    if (!terminated || totalExtentPt > region.columns.length * physicalExtentPt) return null;
    return totalExtentPt / region.columns.length;
  };
  state = setBodyBalanceTarget(state, approximateBalanceTarget(0, input.initialSection));
  const commitTransition = (
    transition: ReturnType<typeof applyAuthoredBreak>,
    nextEntryIndex: number,
    nextOwner?: BodySectionLayoutInput,
  ) => {
    const previousPageIndex = state.flow.pageIndex;
    state = commitPageFlowTransition(state, transition, factory);
    if (nextOwner) {
      state = setBodyBalanceTarget(
        state,
        approximateBalanceTarget(nextEntryIndex, nextOwner),
      );
    }
    const nextLocation = acquisitionLocation(state);
    if (state.flow.pageIndex !== previousPageIndex) {
      session.resetPageAcquisition(nextLocation);
      prescanPageAnchors(state, nextEntryIndex);
    } else {
      session.moveAcquisitionCursor(nextLocation);
    }
  };
  const footnoteIdsByPage = new Map<number, Set<string>>();
  const footnoteAdmission = (
    candidate: ParagraphLayout | TableLayout,
    inlineExtentPt: number,
  ): Readonly<{ ids: readonly string[]; reservePt: number }> => {
    const retained = footnoteIdsByPage.get(state.flow.pageIndex) ?? new Set<string>();
    const ids = footnoteIdsInRetainedSlice(candidate)
      .filter((id) => !retained.has(id));
    return Object.freeze({
      ids: Object.freeze(ids),
      reservePt: session.measureFootnoteReserve({
        referenceIds: ids,
        availableInlineExtentPt: inlineExtentPt,
        firstOnPage: state.footnoteReservePt === 0,
      }),
    });
  };
  const commitFootnotes = (ids: readonly string[], reservePt: number) => {
    let retained = footnoteIdsByPage.get(state.flow.pageIndex);
    if (!retained) {
      retained = new Set<string>();
      footnoteIdsByPage.set(state.flow.pageIndex, retained);
    }
    ids.forEach((id) => retained!.add(id));
    state = addPageFootnoteReserve(state, reservePt);
  };
  let previousParagraph: BodyParagraphSourceInput | null = null;
  const activeColumnBreakIndexes = wordActiveColumnBreakIndexes(input.sequence);

  for (let entryIndex = 0; entryIndex < input.sequence.length; entryIndex += 1) {
    const entry = input.sequence[entryIndex]!;
    if (entry.kind === 'consume-source') {
      state = markBodySourceConsumed(state);
      continue;
    }
    if (entry.kind === 'authored-break') {
      previousParagraph = null;
      if (entry.break === 'column' && !activeColumnBreakIndexes.has(entryIndex)) {
        continue;
      }
      commitTransition(applyAuthoredBreak(state.flow, entry.break), entryIndex + 1);
      continue;
    }
    if (entry.kind === 'begin-section') {
      previousParagraph = null;
      const currentWritingMode = writingModeFromTextDirection(activeRegion(state).section.textDirection);
      const incomingWritingMode = writingModeFromTextDirection(
        sectionContextForPage(entry.section, state.flow.pageIndex).textDirection,
      );
      const currentPhysical = uprightPhysicalExtent({
        widthPt: activeRegion(state).section.geometry.pageWidth,
        heightPt: activeRegion(state).section.geometry.pageHeight,
      }, currentWritingMode);
      const incomingContext = sectionContextForPage(entry.section, state.flow.pageIndex);
      const incomingPhysical = uprightPhysicalExtent({
        widthPt: incomingContext.geometry.pageWidth,
        heightPt: incomingContext.geometry.pageHeight,
      }, incomingWritingMode);
      // §17.6.20 changes the logical page frame; one physical page cannot own
      // section regions whose logical axes use different writing modes.
      const effectiveStartType = entry.section.startType === 'continuous'
        && (currentWritingMode !== incomingWritingMode
          || currentPhysical.widthPt !== incomingPhysical.widthPt
          || currentPhysical.heightPt !== incomingPhysical.heightPt)
        ? 'nextPage'
        : entry.section.startType;
      commitTransition(
        beginSection(
          state.flow,
          flowSection(entry.section, state.flow.pageIndex),
          effectiveStartType,
          { hasFootnoteReferenceOnCurrentPage: state.footnoteReservePt > 0 },
        ),
        entryIndex + 1,
        entry.section,
      );
      continue;
    }
    const block = entry.kind === 'adjacent-table-group' ? entry : entry.block;
    if (block.kind === 'paragraph') {
      if (block.continuousSectionRole === 'collapse-mark') {
        state = markBodySourceConsumed(state);
        continue;
      }
      if (block.pageBreakBefore) {
        commitTransition(
          applyAuthoredBreak(state.flow, 'pageBreakBefore'),
          entryIndex,
        );
      }
      const previousAfterPt = previousParagraph?.spaceAfterPt ?? 0;
      const spacing = paragraphGapAdjustment(
        previousParagraph,
        block,
        previousAfterPt,
        block.continuousSectionRole === 'suppress-before' ? 0 : block.spaceBeforePt,
      );
      const spacingOverlap = block.continuousSectionRole === 'drop-previous-after'
        ? previousAfterPt
        : spacing.overlap;
      if (spacingOverlap > 0) {
        state = Object.freeze({
          ...state,
          flow: Object.freeze({
            ...state.flow,
            cursorBlockPt: Math.max(
              state.flow.regionStartBlockPt,
              state.flow.cursorBlockPt - spacingOverlap,
            ),
          }),
        });
      }
      let cursor: ParagraphFragmentCursor | null = Object.freeze({ boundary: null });
      while (cursor) {
        const fragmentStartKey = paragraphFragmentStartKey(cursor);
        let location = acquisitionLocation(state);
        const acquired = session.measureParagraph({
          input: block,
          location,
          availableInlineExtentPt: location.availableBounds.widthPt,
          suppressSpaceBefore: cursor.boundary !== null
            || block.continuousSectionRole === 'suppress-before'
            || spacing.suppressBefore,
          continuation: cursor,
        });
        if (acquired.placement) {
          const relocationExtentPt = acquired.relocationBlockExtentPt;
          if (
            acquired.placement.sectionFlowOwnership === 'host-flow'
            && relocationExtentPt != null
            && relocationExtentPt > location.availableBounds.heightPt
            && relocationExtentPt <= freshPageExtent(state)
            && state.flow.pageHasContent
          ) {
            commitTransition(
              advanceColumnOrPage(state.flow, 'overflow'),
              entryIndex,
            );
            continue;
          }
          state = acceptNode(
            state,
            acquired.layout,
            block.source,
            acquired.blockExtentPt,
            fragmentStartKey,
            acceptedOccurrenceIds,
            allocations,
            acquired.placement,
          );
          if (acquired.floatRegistryDelta) {
            session.commitFloatRegistryDelta(acquired.floatRegistryDelta);
          }
          cursor = null;
          session.moveAcquisitionCursor(acquisitionLocation(state));
          continue;
        }
        if (cursor.boundary === null && block.keepNext && state.flow.pageHasContent) {
          let keepSetExtentPt = acquired.blockExtentPt;
          let hasTerminalBlock = false;
          for (let nextIndex = entryIndex + 1; nextIndex < input.sequence.length; nextIndex += 1) {
            const nextEntry = input.sequence[nextIndex]!;
            if (nextEntry.kind === 'consume-source') continue;
            if (nextEntry.kind === 'authored-break' || nextEntry.kind === 'begin-section') break;
            const nextBlock = nextEntry.kind === 'adjacent-table-group'
              ? nextEntry
              : nextEntry.block;
            if (nextBlock.kind === 'paragraph' && nextBlock.pageBreakBefore) break;
            const following = session.measureFollowingBlock({
              input: nextBlock,
              location,
              availableInlineExtentPt: location.availableBounds.widthPt,
            });
            const continues = nextBlock.kind === 'paragraph' && nextBlock.keepNext;
            keepSetExtentPt += continues
              ? following.fullExtentPt
              : following.leadContentExtentPt;
            if (!continues) {
              hasTerminalBlock = true;
              break;
            }
          }
          if (
            hasTerminalBlock
            && keepSetExtentPt > location.availableBounds.heightPt
            && keepSetExtentPt <= freshPageExtent(state)
          ) {
            commitTransition(
              advanceColumnOrPage(state.flow, 'overflow'),
              entryIndex,
            );
            continue;
          }
        }
        const nextEntry = input.sequence[entryIndex + 1];
        const afterNext = input.sequence[entryIndex + 2];
        const isImmediatelyPreBreakAnchor = nextEntry?.kind === 'body-block'
          && nextEntry.block.kind === 'paragraph'
          && afterNext?.kind === 'authored-break'
          && afterNext.break === 'page';
        const hasInlineDrawingResource = wordPreBreakInlineDrawingResource(acquired.layout);
        if (
          cursor.boundary === null
          && hasInlineDrawingResource
          && isImmediatelyPreBreakAnchor
          && state.flow.pageHasContent
        ) {
          const anchorLocation = locationAfter(location, acquired.blockExtentPt);
          const anchor = session.measureParagraph({
            input: nextEntry.block,
            location: anchorLocation,
            availableInlineExtentPt: anchorLocation.availableBounds.widthPt,
            suppressSpaceBefore: false,
            continuation: Object.freeze({ boundary: null }),
          });
          session.moveAcquisitionCursor(location);
          const anchorExtentPt = wordPreBreakHostAnchorExtentPt(
            anchor.layout,
            anchorLocation.cursorPt.yPt,
          );
          if (anchorExtentPt !== null) {
            const groupExtentPt = acquired.blockExtentPt + anchorExtentPt;
            if (
              groupExtentPt > location.availableBounds.heightPt
              && groupExtentPt <= freshPageExtent(state)
            ) {
              commitTransition(
                advanceColumnOrPage(state.flow, 'overflow'),
                entryIndex,
              );
              continue;
            }
          }
        }
        const followingEntry = input.sequence[entryIndex + 1];
        const followedByHardPageBreak = followingEntry?.kind === 'authored-break'
          && followingEntry.break === 'page';
        if (
          cursor.boundary === null
          && followedByHardPageBreak
        ) {
          const neutral = wordFlowNeutralPreBreakAnchorParagraph(acquired.layout);
          if (neutral !== null) {
            state = acceptNode(
              state,
              neutral,
              block.source,
              0,
              fragmentStartKey,
              acceptedOccurrenceIds,
              allocations,
            );
            if (acquired.floatRegistryDelta) {
              session.commitFloatRegistryDelta(acquired.floatRegistryDelta);
            }
            cursor = null;
            session.moveAcquisitionCursor(acquisitionLocation(state));
            continue;
          }
        }
        const markReservePt = footnoteAdmission(
          acquired.layout,
          location.availableBounds.widthPt,
        ).reservePt;
        const pageBottomIsUnreserved = (reserves[state.flow.pageIndex]?.bottom ?? 0) === 0
          && state.footnoteReservePt === 0;
        const physicalRegionBottomIsActive = activeBlockEndPt(state) === activeRegion(state).blockEndPt;
        // Word admits an undecorated empty mark by its baseline only at the physical body edge.
        const trailingMarkAdmissionAllowancePt = cursor.boundary === null
          && block.inkless === true
          && isUndecoratedInklessMark(acquired.layout)
          && !block.keepNext
          && markReservePt === 0
          && pageBottomIsUnreserved
          && physicalRegionBottomIsActive
          && hasFollowingInkContent(input, entryIndex + 1)
          ? acquired.markBelowBaselinePt ?? 0
          : 0;
        const selected = selectParagraphFragment(
          acquired.layout,
          cursor,
          acquired.lineEndBoundaries,
          location.availableBounds.heightPt + trailingMarkAdmissionAllowancePt,
          freshPageExtent(state),
          state.flow.pageHasContent,
          { keepLines: block.keepLines, widowControl: block.widowControl },
          (fragment) => footnoteAdmission(
            fragment,
            location.availableBounds.widthPt,
          ).reservePt,
          acquired.uniformRubyAdvancePt,
        );
        if (selected.requiresFreshFlowRegion) {
          commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
          continue;
        }
        if (!selected.fragment) throw new Error('Paragraph acquisition made no progress');
        state = acceptNode(
          state,
          selected.fragment,
          block.source,
          Math.min(selected.admittedBlockExtentPt, location.availableBounds.heightPt),
          fragmentStartKey,
          acceptedOccurrenceIds,
          allocations,
          acquired.placement,
        );
        const notes = footnoteAdmission(
          selected.fragment,
          location.availableBounds.widthPt,
        );
        commitFootnotes(notes.ids, notes.reservePt);
        if (acquired.floatRegistryDelta) {
          const acceptedDelta = paragraphFloatDeltaForAcceptedFragment(
            acquired.floatRegistryDelta,
            selected.fragment,
          );
          if (acceptedDelta) session.commitFloatRegistryDelta(acceptedDelta);
        }
        cursor = selected.nextCursor;
        if (cursor) {
          commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
        }
        location = acquisitionLocation(state);
        session.moveAcquisitionCursor(location);
      }
      previousParagraph = block;
    } else {
      previousParagraph = null;
      let cursor: import('./body-layout-kernel.js').BodyTableContinuationCursor | undefined;
      let complete = false;
      while (!complete) {
        const fragmentStartKey = tableFragmentStartKey(cursor);
        const location = acquisitionLocation(state);
        const requestAt = (availableBlockExtentPt: number) => session.measureTable({
            input: block,
            location,
            availableInlineExtentPt: location.availableBounds.widthPt,
            availableBlockExtentPt,
            freshPageBlockExtentPt: freshPageExtent(state),
            ...(cursor ? { cursor } : {}),
          });
        let availableBlockExtentPt = location.availableBounds.heightPt;
        let acquired = requestAt(availableBlockExtentPt);
        if (acquired.retryAtBlockStartPt !== undefined) {
          if (!Number.isFinite(acquired.retryAtBlockStartPt)
            || acquired.retryAtBlockStartPt <= state.flow.cursorBlockPt) {
            throw new Error('Table repositioning must advance the block cursor');
          }
          state = Object.freeze({
            ...state,
            flow: Object.freeze({
              ...state.flow,
              cursorBlockPt: acquired.retryAtBlockStartPt,
            }),
          });
          session.moveAcquisitionCursor(acquisitionLocation(state));
          continue;
        }
        let notes = acquired.requiresFreshFlowRegion
          ? Object.freeze({ ids: Object.freeze([]) as readonly string[], reservePt: 0 })
          : footnoteAdmission(
              acquired.layout,
              location.availableBounds.widthPt,
            );
        const seenCandidates = new Set<string>();
        while (
          !acquired.requiresFreshFlowRegion
          && acquired.blockExtentPt + notes.reservePt > location.availableBounds.heightPt
        ) {
          const fingerprint = JSON.stringify({
            advancePt: acquired.blockExtentPt,
            nextCursor: acquired.nextCursor ?? null,
            noteIds: notes.ids,
            reservePt: notes.reservePt,
          });
          if (seenCandidates.has(fingerprint)) {
            throw new Error('Table footnote admission did not converge');
          }
          seenCandidates.add(fingerprint);
          availableBlockExtentPt = Math.max(
            0,
            location.availableBounds.heightPt - notes.reservePt,
          );
          acquired = requestAt(availableBlockExtentPt);
          notes = acquired.requiresFreshFlowRegion
            ? Object.freeze({ ids: Object.freeze([]) as readonly string[], reservePt: 0 })
            : footnoteAdmission(
                acquired.layout,
                location.availableBounds.widthPt,
              );
        }
        if (acquired.requiresFreshFlowRegion) {
          const rebasesFloatingTableOnFreshFrame = !state.flow.pageHasContent
            && acquired.nextCursor?.kind === 'table'
            && acquired.nextCursor.floatingContinuationFrame === 'fresh-text'
            && !(cursor?.kind === 'table' && cursor.floatingContinuationFrame !== undefined);
          if (acquired.nextCursor?.kind === 'table'
            && acquired.nextCursor.floatingContinuationFrame !== undefined) {
            cursor = acquired.nextCursor;
          }
          if (rebasesFloatingTableOnFreshFrame) continue;
          commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
          continue;
        }
        state = acceptNode(
          state,
          acquired.layout,
          block.source,
          acquired.blockExtentPt,
          fragmentStartKey,
          acceptedOccurrenceIds,
          allocations,
          acquired.placement,
        );
        commitFootnotes(notes.ids, notes.reservePt);
        if (acquired.floatRegistryDelta) {
          session.commitFloatRegistryDelta(bindTableFloatDeltaToAcceptedOccurrence(
            acquired.floatRegistryDelta,
            acquired.layout,
            bodyOccurrenceKey(block.source, location.flowDomainId, fragmentStartKey),
          ));
        }
        cursor = acquired.nextCursor ?? undefined;
        complete = cursor === undefined;
        if (cursor) {
          commitTransition(
            advanceColumnOrPage(state.flow, 'overflow'),
            entryIndex,
          );
        }
      }
    }
    session.moveAcquisitionCursor(acquisitionLocation(state));
  }
  return Object.freeze({
    layout: finalize(state, owners),
    session,
    allocations: Object.freeze(allocations),
  });
}

function headerFooterReserves(
  pass: Readonly<{ layout: DocumentLayout; session: BodyLayoutSession }>,
  owners: ReadonlyMap<string, BodySectionLayoutInput>,
): readonly HeaderFooterReserve[] {
  const firstSectionPage = new Map<string, number>();
  pass.layout.pages.forEach((page) => {
    page.sectionRegions.forEach((region) => {
      if (!firstSectionPage.has(region.sectionOccurrenceId)) {
        firstSectionPage.set(region.sectionOccurrenceId, page.pageIndex);
      }
    });
  });
  return Object.freeze(pass.layout.pages.map((page) => {
    if (page.parityBlank) return Object.freeze({ top: 0, bottom: 0 });
    // Vertical header/footer stories paint in physical page space; charging their
    // measured overflow to the logical body interval would create a pagination-only reserve.
    if (writingModeFromTextDirection(page.section.textDirection) !== 'horizontal-tb') {
      return Object.freeze({ top: 0, bottom: 0 });
    }
    const owner = owners.get(page.sectionOccurrenceId);
    if (!owner) throw new Error(`Unknown body section ${page.sectionOccurrenceId}`);
    const firstPageIndex = firstSectionPage.get(owner.sectionOccurrenceId) ?? page.pageIndex;
    const inlineExtentPt = Math.max(
      0,
      page.section.geometry.pageWidth
        - Math.abs(page.section.geometry.marginLeft)
        - Math.abs(page.section.geometry.marginRight),
    );
    const measure = (kind: 'header' | 'footer') => {
      const source = selectedHeaderFooterStory(
        kind === 'header' ? owner.headers : owner.footers,
        {
          titlePage: owner.titlePage,
          firstPageOfSection: page.pageIndex === firstPageIndex,
          evenAndOddHeaders: owner.evenAndOddHeaders,
          displayPageNumber: page.pageNumber.displayNumber,
        },
      );
      return source === null ? 0 : pass.session.measureStoryExtent({
        source,
        pageIndex: page.pageIndex,
        section: page.section,
        availableInlineExtentPt: inlineExtentPt,
      });
    };
    return Object.freeze({
      top: headerFooterOverflowReservePt(
        measure('header'),
        page.section.geometry.marginTop,
        page.section.geometry.headerDistance,
      ),
      bottom: headerFooterOverflowReservePt(
        measure('footer'),
        page.section.geometry.marginBottom,
        page.section.geometry.footerDistance,
      ),
    });
  }));
}

function pageAnchorDestinationPlan(layout: DocumentLayout) {
  const destinations = new Map<string, Readonly<{
    occurrenceId: string;
    paragraphSource: SourceRef;
    pageIndex: number;
    flowDomainId: string;
  }>>();
  for (const page of layout.pages) {
    for (const node of page.layers.body) {
      if (node.kind !== 'paragraph') continue;
      for (const drawing of node.drawings) {
        const anchor = drawing.anchorLayer;
        if (!anchor
          || anchor.horizontalOwnership !== 'page'
          || anchor.verticalOwnership !== 'page') continue;
        const occurrenceId = anchor.acquisitionOccurrenceId ?? anchor.occurrenceId;
        destinations.set(occurrenceId, Object.freeze({
          occurrenceId,
          paragraphSource: node.source,
          pageIndex: page.pageIndex,
          flowDomainId: node.flowDomainId,
        }));
      }
    }
  }
  return destinations;
}

function anchorPlanIdentity(plan: ReadonlyMap<string, unknown>): string {
  return JSON.stringify([...plan].sort(([left], [right]) => left.localeCompare(right)));
}

function paginateBodyWithAnchorConvergence(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
  reserves: readonly HeaderFooterReserve[],
) {
  const hasPageOwnedAnchors = input.sequence.some((entry) => (
    entry.kind === 'body-block'
    && entry.block.kind === 'paragraph'
    && (entry.block.pageOwnedAnchorOccurrenceIds?.length ?? 0) > 0
  ));
  let plan: ReturnType<typeof pageAnchorDestinationPlan> | null = null;
  const seen = new Set<string>();
  while (true) {
    const pass = paginateBodyPass(input, services, options, reserves, plan);
    const nextPlan = pageAnchorDestinationPlan(pass.layout);
    if (!hasPageOwnedAnchors) return pass;
    const nextIdentity = anchorPlanIdentity(nextPlan);
    if (plan !== null && nextIdentity === anchorPlanIdentity(plan)) return pass;
    if (seen.has(nextIdentity)) {
      throw new Error('Page-anchor destination acquisition did not converge');
    }
    seen.add(nextIdentity);
    plan = nextPlan;
  }
}

export function paginateBody(
  input: BodyLayoutInput,
  services: LayoutServices,
  options: LayoutOptions,
): DocumentLayout {
  const owners = ownerMap(input);
  const seed = paginateBodyWithAnchorConvergence(input, services, options, []);
  const converged = convergeHeaderFooterReserves({
    seed,
    measure: (pass) => headerFooterReserves(pass, owners),
    repaginate: (reserves, current) => {
      const contexts = current.layout.pages.map((page) => Object.freeze({
        pageIndex: page.pageIndex,
        displayPageNumber: page.pageNumber.displayNumber,
        pageNumberFormat: page.pageNumber.format as import('@silurus/ooxml-core').NumberFormat,
      }));
      const iterationServices = createFieldAcquisitionServicesView(services, {
        totalPages: current.layout.pages.length,
        resolveDestinationPage: (pageIndex) => contexts[pageIndex],
      });
      return paginateBodyWithAnchorConvergence(input, iterationServices, options, reserves);
    },
    identity: (pass) => pass.layout,
    requiresConvergence: seed.session.hasPaginationFields,
  }).result;
  const composed = composeCanonicalSectionFlow(
    converged.layout,
    converged.session,
    converged.allocations,
  );
  assertDocumentLayout(composed);
  return deepFreezeDocumentLayout(composed) as DocumentLayout;
}
