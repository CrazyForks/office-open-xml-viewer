import {
  loadMathJax,
  mathMLToSvg,
} from '../../../core/src/math/engine.js';

/**
 * Local Word-PDF acceptance must exercise the same opt-in math route as the
 * Storybook viewer. Leaving the engine out silently drops equations before
 * pagination and can fabricate both pixel and page-count regressions.
 */
export const math = { loadMathJax, mathMLToSvg };
