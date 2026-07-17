import { adjustForWidowOrphan, selectLargestFittingEnd } from '../line-fit-policy.js';
import type { LineBoundary } from '../line-layout.js';
import { sliceParagraphLayout } from './paragraph.js';
import type { ParagraphLayout } from './types.js';

export interface ParagraphFragmentCursor {
  readonly boundary: LineBoundary | null;
  readonly sourceRangeStart?: number;
  readonly uniformRubyAdvancePt?: number;
}

export interface ParagraphFragmentSelection {
  readonly fragment: ParagraphLayout | null;
  readonly nextCursor: ParagraphFragmentCursor | null;
  readonly requiresFreshFlowRegion: boolean;
  readonly additionalReservePt: number;
  /** Flow charge admitted to this region; retained line/ink geometry can be taller. */
  readonly admittedBlockExtentPt: number;
}

function compareLineBoundaries(left: LineBoundary, right: LineBoundary): number {
  return left.segIndex - right.segIndex || left.charOffset - right.charOffset;
}

export function selectParagraphFragment(
  acquired: ParagraphLayout,
  cursor: ParagraphFragmentCursor,
  lineEndBoundaries: readonly LineBoundary[],
  availableBlockExtentPt: number,
  freshFlowRegionBlockExtentPt: number,
  canRelocate: boolean,
  policy: Readonly<{ keepLines: boolean; widowControl: boolean }>,
  additionalReserveFor?: (fragment: ParagraphLayout) => number,
  uniformRubyAdvancePt?: number,
  additionalReserveFits?: (reservePt: number) => boolean,
): ParagraphFragmentSelection {
  if (![availableBlockExtentPt, freshFlowRegionBlockExtentPt].every(
    (value) => Number.isFinite(value) && value >= 0,
  )) throw new RangeError('Paragraph fragment extents must be finite and non-negative');
  if (lineEndBoundaries.length !== acquired.lines.length) {
    throw new RangeError('Paragraph source boundaries must align with retained lines');
  }
  const total = acquired.lines.length;
  const slice = (end: number) => sliceParagraphLayout(acquired, {
    lineStart: 0,
    lineEnd: end,
    continuesFromPrevious: cursor.boundary !== null,
    continuesOnNext: end < total,
  });
  const reserveFor = (fragment: ParagraphLayout): number => {
    const reserve = additionalReserveFor?.(fragment) ?? 0;
    if (!Number.isFinite(reserve) || reserve < 0) {
      throw new RangeError('Paragraph page-local reserve must be finite and non-negative');
    }
    return reserve;
  };
  const reserveFits = (reservePt: number): boolean => additionalReserveFits?.(reservePt) ?? true;
  if (total === 0) {
    const reserve = reserveFor(acquired);
    if (canRelocate && (
      acquired.advancePt + reserve > availableBlockExtentPt
      || !reserveFits(reserve)
    )
      && acquired.advancePt + reserve <= freshFlowRegionBlockExtentPt) {
      return {
        fragment: null, nextCursor: cursor,
        requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0,
      };
    }
    return {
      fragment: acquired, nextCursor: null,
      requiresFreshFlowRegion: false, additionalReservePt: reserve,
      admittedBlockExtentPt: Math.min(acquired.advancePt, availableBlockExtentPt),
    };
  }
  const completeReserve = reserveFor(acquired);
  if (cursor.boundary === null && policy.keepLines && canRelocate
    && (
      acquired.advancePt + completeReserve > availableBlockExtentPt
      || !reserveFits(completeReserve)
    )
    && acquired.advancePt + completeReserve <= freshFlowRegionBlockExtentPt) {
    return {
      fragment: null, nextCursor: cursor,
      requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0,
    };
  }
  let end = selectLargestFittingEnd(0, total, availableBlockExtentPt, (lineEnd) => (
    (() => {
      const candidate = slice(lineEnd);
      const reserve = reserveFor(candidate);
      return reserveFits(reserve)
        ? candidate.advancePt + reserve
        : availableBlockExtentPt + 1;
    })()
  )).end;
  if (end === 0) {
    if (canRelocate) return {
      fragment: null, nextCursor: cursor,
      requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0,
    };
    end = 1;
  }
  for (;;) {
    const widow = adjustForWidowOrphan({
      widowControl: policy.widowControl,
      start: 0,
      end,
      totalLines: total,
      canRelocate,
    });
    if (widow.kind === 'relocate') {
      return {
        fragment: null, nextCursor: cursor,
        requiresFreshFlowRegion: true, additionalReservePt: 0, admittedBlockExtentPt: 0,
      };
    }
    if (widow.kind !== 'dropLastLine') break;
    end -= 1;
  }
  const fragment = slice(end);
  const nextBoundary = end < total ? lineEndBoundaries[end - 1]! : null;
  if (
    nextBoundary !== null
    && cursor.boundary !== null
    && compareLineBoundaries(nextBoundary, cursor.boundary) <= 0
  ) {
    throw new Error('Paragraph continuation source boundary did not advance');
  }
  return {
    fragment,
    nextCursor: nextBoundary === null ? null : Object.freeze({
      boundary: nextBoundary,
      sourceRangeStart: fragment.lines.at(-1)!.range.end,
      ...(uniformRubyAdvancePt === undefined ? {} : { uniformRubyAdvancePt }),
    }),
    requiresFreshFlowRegion: false,
    additionalReservePt: reserveFor(fragment),
    admittedBlockExtentPt: Math.min(fragment.advancePt, availableBlockExtentPt),
  };
}
