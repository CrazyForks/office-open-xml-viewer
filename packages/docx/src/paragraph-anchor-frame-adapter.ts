import type { AnchorReferenceFramesInput } from './layout/anchor-frame.js';

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

/** Renderer state still mixes CSS pixels with authored points, so conversion
 * remains outside layout and only point-space frames cross that boundary. */
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
