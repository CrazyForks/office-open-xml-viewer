import type { ChartRun, ImageRun, ShapeRun } from '../types.js';
import { wordPageLevelAnchorY } from './anchor-compatibility.js';
import { isWrapFloat } from './float-wrap.js';

/** Whether vertical anchor placement is independent of source-order paragraph
 * position under ECMA-376 §20.4.3.5 and the registered Word compatibility rule. */
export function isPageLevelAnchorY(
  relativeFrom: string | null | undefined,
  fromParagraph: boolean,
): boolean {
  return wordPageLevelAnchorY(relativeFrom, fromParagraph);
}

export function isPageLevelWrapFloat(run: ImageRun | ChartRun | ShapeRun): boolean {
  return isWrapFloat(run.wrapMode)
    && isPageLevelAnchorY(
      run.anchorYRelativeFrom ?? null,
      run.anchorYFromPara ?? false,
    );
}
