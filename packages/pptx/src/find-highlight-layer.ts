/**
 * IX2 pptx find-highlight overlay.
 *
 * The highlight twin of {@link buildPptxTextLayer}: it draws a visible box per
 * matched run-slice, grouped into one positioned + rotated `<div>` per shape
 * frame exactly as the selection overlay groups its transparent spans, so a box
 * tracks the drawn (rotated) text. Riding the same shape-grouped DOM overlay
 * (rather than a canvas draw pass) means highlights rotate with their shape and
 * compose with the selection / hyperlink layer.
 *
 * A slice's horizontal extent within its run is the shared core
 * `sliceHorizontalExtent`, measured against the run's font; the box's vertical
 * extent is the run's line box (`h`). Boxes are placed in the shape's own
 * coordinate frame (`inShapeX`/`inShapeY`), so the shape div's `rotate()` lays
 * them along the glyphs. The active match uses a distinct emphasis colour.
 */
import { sliceHorizontalExtent, type MatchRunSlice } from '@silurus/ooxml-core';
import type { PptxTextRunInfo } from './renderer';

export interface PptxHighlightMatch {
  slices: MatchRunSlice[];
  active: boolean;
}

/** Browser find-bar palette (translucent so glyphs stay legible). */
export const DEFAULT_FIND_HIGHLIGHT = 'rgba(255, 214, 0, 0.42)';
export const DEFAULT_FIND_ACTIVE_HIGHLIGHT = 'rgba(255, 140, 0, 0.55)';

export interface PptxHighlightColors {
  match?: string;
  active?: string;
}

/**
 * Populate a highlight overlay layer with a box per matched run-slice, grouped
 * by shape frame (with the shape's rotation) so each box lands on the drawn
 * glyphs.
 *
 * @param layer     the overlay div (cleared + re-sized here).
 * @param runs      the slide's runs (same array the slide was rendered from).
 * @param matches   the slide's matches (run-slices + active flag).
 * @param cssWidth  rendered canvas CSS width (px, number).
 * @param cssHeight rendered canvas CSS height (px, number).
 * @param measureForFont returns a width-measurer primed with a run's font.
 * @param colors    optional colour overrides.
 */
export function buildPptxHighlightLayer(
  layer: HTMLDivElement,
  runs: PptxTextRunInfo[],
  matches: PptxHighlightMatch[],
  cssWidth: number,
  cssHeight: number,
  measureForFont: (font: string) => (s: string) => number,
  colors: PptxHighlightColors = {},
): void {
  layer.innerHTML = '';
  layer.style.width = `${cssWidth}px`;
  layer.style.height = `${cssHeight}px`;

  const matchColor = colors.match ?? DEFAULT_FIND_HIGHLIGHT;
  const activeColor = colors.active ?? DEFAULT_FIND_ACTIVE_HIGHLIGHT;

  // One positioned + rotated div per shape frame (keyed like the text layer), so
  // boxes inside it inherit the shape's rotation. Reused across matches/slices.
  const shapeMap = new Map<string, HTMLDivElement>();
  const shapeDiv = (run: PptxTextRunInfo): HTMLDivElement => {
    const totalRot = run.rotation + (run.textBodyRotation ?? 0);
    const key = `${run.shapeX},${run.shapeY},${run.shapeW},${run.shapeH},${totalRot}`;
    let div = shapeMap.get(key);
    if (!div) {
      div = document.createElement('div');
      div.style.cssText =
        `position:absolute;` +
        `left:${run.shapeX}px;top:${run.shapeY}px;` +
        `width:${run.shapeW}px;height:${run.shapeH}px;` +
        `pointer-events:none;overflow:hidden;`;
      if (totalRot !== 0) {
        div.style.transformOrigin = 'center center';
        div.style.transform = `rotate(${totalRot}deg)`;
      }
      shapeMap.set(key, div);
      layer.appendChild(div);
    }
    return div;
  };

  for (const match of matches) {
    const fill = match.active ? activeColor : matchColor;
    for (const slice of match.slices) {
      const run = runs[slice.runIndex];
      if (!run) continue;
      const measure = measureForFont(run.font);
      const { x, width } = sliceHorizontalExtent(run.text, slice.start, slice.end, measure);
      if (width <= 0) continue;
      const box = document.createElement('div');
      box.style.cssText =
        `position:absolute;` +
        `left:${run.inShapeX + x}px;top:${run.inShapeY}px;` +
        `width:${width}px;height:${run.h}px;` +
        `background:${fill};pointer-events:none;`;
      shapeDiv(run).appendChild(box);
    }
  }
}
