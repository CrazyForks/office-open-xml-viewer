import type { AnchorReferenceFramesInput } from './anchor-frame.js';
import { normalizeWrapSide, type FloatRect } from './float-wrap.js';
import type {
  DrawingMLCollisionEntryPt,
  WrapExclusion,
} from './types.js';

/** Project the compatibility float registry into retained text-wrap authority.
 * A DrawingML wrapNone object remains absent here because §20.4.2.3 gives it
 * object-collision semantics without adding text exclusion geometry. */
export function paragraphWrapExclusions(
  floats: readonly FloatRect[],
  flowDomainId: string,
): readonly WrapExclusion[] {
  return floats.flatMap((float, index): WrapExclusion[] => {
    if (
      float.kind === 'shape'
      && float.anchorOccurrenceId
      && float.authoredWrap === undefined
    ) return [];
    return [{
      id: float.imageKey || `${flowDomainId}:float:${index}`,
      wrap: float.authoredWrap
        ?? (float.mode === 'topAndBottom' ? 'topAndBottom' : 'square'),
      wrapSide: normalizeWrapSide(float.side),
      bounds: {
        xPt: float.xLeft,
        yPt: float.yTop,
        widthPt: Math.max(0, float.xRight - float.xLeft),
        heightPt: Math.max(0, float.yBottom - float.yTop),
      },
      polygon: float.wrapPolygon ?? [
        { xPt: float.xLeft, yPt: float.yTop },
        { xPt: float.xRight, yPt: float.yTop },
        { xPt: float.xRight, yPt: float.yBottom },
        { xPt: float.xLeft, yPt: float.yBottom },
      ],
      ...(float.kind === 'table' && !float.anchorOccurrenceId
        ? { verticalOwnership: 'page' as const }
        : {}),
      ...(float.anchorOccurrenceId
        ? {
            anchorOccurrenceId: float.anchorOccurrenceId,
            verticalOwnership: 'page' as const,
          }
        : {}),
    }];
  });
}

/** Collision authority is intentionally independent from text-wrap authority:
 * DrawingML wrapNone objects still participate in §20.4.2.3 avoidance. */
export function paragraphAnchorCollisions(
  floats: readonly FloatRect[],
): readonly DrawingMLCollisionEntryPt[] {
  return floats.flatMap((float): DrawingMLCollisionEntryPt[] => {
    if (float.kind !== 'shape' || !float.anchorOccurrenceId) return [];
    return [{
      occurrenceId: float.anchorOccurrenceId,
      bounds: {
        xPt: float.imageX,
        yPt: float.imageY,
        widthPt: float.imageW,
        heightPt: float.imageH,
      },
      horizontalOwnership: 'page',
      verticalOwnership: 'page',
    }];
  });
}

export interface ParagraphAnchorReferenceFrameSnapshot {
  readonly pageIndex: number;
  readonly scale: number;
  readonly pageWidth: number;
  readonly pageH: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly contentX: number;
  readonly contentW: number;
}

/** Resolve the renderer's mixed legacy px/pt state at the layout boundary so
 * anchor acquisition receives one point-space reference-frame snapshot. */
export function paragraphAnchorReferenceFrames(
  snapshot: ParagraphAnchorReferenceFrameSnapshot,
): Readonly<Pick<
  AnchorReferenceFramesInput,
  'page' | 'margin' | 'column' | 'pageParity'
>> {
  const pageHeightPt = snapshot.pageH / snapshot.scale;
  const blockExtentPt = Math.max(
    0,
    pageHeightPt - snapshot.marginTop - snapshot.marginBottom,
  );
  return {
    page: {
      xPt: 0,
      yPt: 0,
      widthPt: snapshot.pageWidth,
      heightPt: pageHeightPt,
    },
    margin: {
      xPt: snapshot.marginLeft,
      yPt: snapshot.marginTop,
      widthPt: Math.max(
        0,
        snapshot.pageWidth - snapshot.marginLeft - snapshot.marginRight,
      ),
      heightPt: blockExtentPt,
    },
    column: {
      xPt: snapshot.contentX / snapshot.scale,
      yPt: snapshot.marginTop,
      widthPt: snapshot.contentW / snapshot.scale,
      heightPt: blockExtentPt,
    },
    pageParity: snapshot.pageIndex % 2 === 0 ? 'odd' : 'even',
  };
}
