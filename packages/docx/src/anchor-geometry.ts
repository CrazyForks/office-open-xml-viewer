// DrawingML anchor placement geometry (ECMA-376 §20.4.3.x).
//
// Pure placement math for `<wp:positionH>` / `<wp:positionV>` (relativeFrom +
// posOffset / align / pctPos): given the container indicated by `relativeFrom`
// and the explicit `AnchorGeometryContext`, it answers "where does this anchor /
// anchor-group child sit?". Extracted from renderer.ts so the resolve logic can
// be unit-reasoned in isolation (see anchor-align.test.ts).

import type { AnchorGeometryContext } from './layout/acquisition-context.js';

/** Resolve a shape's page X by combining the explicit `anchorXPt` offset with
 *  any `anchorXAlign` (ECMA-376 §20.4.3.1 wp:align). When align is set we
 *  position the shape inside the container indicated by `relativeFrom` (or
 *  `anchorXFromMargin` for the legacy two-state hint). When `pctPos` is set
 *  we ignore the explicit offset and place the shape at `pct` of the
 *  container's width / height (ECMA-376 §20.4.2.7 wp14:pctPosH/VOffset).
 *
 *  relativeFrom containers (ECMA-376 §20.4.3.4):
 *    - "page"          → full page rect
 *    - "margin"        → printable area between margins
 *    - "leftMargin"    → strip from x=0 to x=marginLeft
 *    - "rightMargin"   → strip from x=pageW-marginRight to x=pageW
 *    - "insideMargin"  → on odd pages = leftMargin, even = rightMargin
 *                        (we approximate as leftMargin)
 *    - "outsideMargin" → on odd pages = rightMargin, even = leftMargin
 *                        (we approximate as rightMargin)
 *    - "column"        → the current text column (state.contentX/contentW):
 *                        the margin band at body level, a specific column in a
 *                        multi-column section, or a table cell's inner text box
 *    - "character"     → degrade to "column" (no run-relative offset data), i.e.
 *                        the containing text column — the closest available base
 *    - "topMargin"     → strip from y=0 to y=marginTop
 *    - "bottomMargin"  → strip from y=pageH-marginBottom to y=pageH
 *    - "paragraph"/"line" → relative to paragraph top (V only) */
export function xContainer(
  relativeFrom: string | null | undefined,
  fromMarginHint: boolean,
  state: AnchorGeometryContext,
): { start: number; end: number } {
  const pageW = state.pageWidth;
  const ml = state.marginLeft;
  const mr = state.marginRight;
  const rf = relativeFrom ?? (fromMarginHint ? 'margin' : 'page');
  switch (rf) {
    case 'page':          return { start: 0, end: pageW };
    case 'leftMargin':    return { start: 0, end: ml };
    case 'rightMargin':   return { start: pageW - mr, end: pageW };
    case 'insideMargin':  return { start: 0, end: ml };
    case 'outsideMargin': return { start: pageW - mr, end: pageW };
    // ECMA-376 §20.4.3.4 ST_RelFromH: `column` is "relative to the extents of the
    // COLUMN which contains its anchor" — the current TEXT column, not the page
    // margins. `character` is "relative to the position of the anchor within its
    // run content"; with no run-relative offset data we degrade it to the same
    // containing column (the closest available base). The renderer keeps the
    // current point-space column band in state.contentX/contentW: the
    // section margin band at body level, a specific column band in a multi-column
    // section, or a table CELL's inner text box while rendering cell content
    // (renderCell). This lets a header-logo anchor authored `relativeFrom="column"`
    // inside an RTL bidi cell land in that cell's column (sample-28) instead of
    // being flattened onto the page margin band.
    case 'character':
    case 'column':        return { start: state.contentX, end: state.contentX + state.contentW };
    case 'margin':
    default:              return { start: ml, end: pageW - mr };
  }
}

export function yContainer(
  relativeFrom: string | null | undefined,
  fromParaHint: boolean,
  paragraphTopPt: number,
  state: AnchorGeometryContext,
): { start: number; end: number } {
  const mt = state.marginTop;
  const mb = state.marginBottom;
  const rf = relativeFrom ?? (fromParaHint ? 'paragraph' : 'page');
  switch (rf) {
    case 'page':         return { start: 0, end: state.pageH };
    case 'topMargin':    return { start: 0, end: mt };
    case 'bottomMargin': return { start: state.pageH - mb, end: state.pageH };
    case 'paragraph':
    case 'line':         return { start: paragraphTopPt, end: state.pageH };
    case 'margin':
    default:             return { start: mt, end: state.pageH - mb };
  }
}

/** Resolve the page X in points for an anchor or anchor-group child. `offsetPt`
 *  carries
 *  the shape's offset for explicit posOffset anchors, or the within-group child
 *  offset when `alignWidthPt` is set. For standalone `<wp:align>` anchors it is
 *  ignored: ECMA-376 §20.4.3.1 uses `<wp:align>` / `<wp:posOffset>` as a choice,
 *  and Word commonly leaves a duplicate a:xfrm/simplePos fallback that must not
 *  be added to the aligned position. `alignWidthPt` is the width used when
 *  aligning — the GROUP's width for wgp children, the shape's own width for
 *  standalone anchors. */
export function resolveAnchorX(
  align: string | null | undefined,
  fromMargin: boolean,
  offsetPt: number,
  widthPt: number,
  state: AnchorGeometryContext,
  relativeFrom?: string | null,
  pctPos?: number | null,
  alignWidthPt?: number | null,
): number {
  const c = xContainer(relativeFrom, fromMargin, state);
  if (pctPos != null) {
    return c.start + (c.end - c.start) * pctPos + offsetPt;
  }
  if (!align) {
    return c.start + offsetPt;
  }
  const containerW = c.end - c.start;
  const aw = alignWidthPt ?? widthPt;
  const alignOffsetPt = alignWidthPt != null ? offsetPt : 0;
  switch (align) {
    case 'center': return c.start + (containerW - aw) / 2 + alignOffsetPt;
    case 'right':
    case 'outside': return c.end - aw + alignOffsetPt;
    case 'inside':
    case 'left':
    default:        return c.start + alignOffsetPt;
  }
}

export function resolveAnchorY(
  align: string | null | undefined,
  fromPara: boolean,
  offsetPt: number,
  heightPt: number,
  paragraphTopPt: number,
  state: AnchorGeometryContext,
  relativeFrom?: string | null,
  pctPos?: number | null,
  alignHeightPt?: number | null,
): number {
  const c = yContainer(relativeFrom, fromPara, paragraphTopPt, state);
  if (pctPos != null) {
    return c.start + (c.end - c.start) * pctPos + offsetPt;
  }
  if (!align) {
    return c.start + offsetPt;
  }
  const containerH = c.end - c.start;
  const ah = alignHeightPt ?? heightPt;
  const alignOffsetPt = alignHeightPt != null ? offsetPt : 0;
  switch (align) {
    case 'center': return c.start + (containerH - ah) / 2 + alignOffsetPt;
    // ECMA-376 §20.4.3.1 ST_AlignV: "inside"/"outside" are page-binding-
    // relative. Mirroring resolveAnchorX (and the insideMargin/outsideMargin
    // approximation in yContainer/xContainer): on an odd page the binding edge
    // is the top, so inside→top edge and outside→bottom edge. This is an
    // odd-page approximation; the true §20.4.3.1 page-parity behavior (even
    // pages mirror the binding edge) is not implemented. Update this when that
    // approximation is removed.
    case 'bottom':
    case 'outside': return c.end - ah + alignOffsetPt;
    case 'top':
    case 'inside':
    default:        return c.start + alignOffsetPt;
  }
}
