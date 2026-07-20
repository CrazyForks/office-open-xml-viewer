import { defineCompatibilityRule } from './compatibility.js';

export const WORD_TRACK_CHANGE_AUTHOR_PALETTE = defineCompatibilityRule({
  id: 'word-track-change-author-palette',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/compatibility.test.ts#pins the eight track-change author colors independently of author indexing',
  },
  description: 'Use the established eight-color revision-author palette while keeping the renderer deterministic author-index policy outside this compatibility claim.',
});

export const WORD_PARAGRAPH_SHADING_BORDER_BOX = defineCompatibilityRule({
  id: 'word-paragraph-shading-border-box',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/para-shading-border.test.ts#extends the fill by a present border edge’s space (matching drawParaBorders)',
  },
  description: 'Extend paragraph shading through each visible paragraph-border spacing interval so the fill reaches the painted border box.',
});

export const WORD_TRACK_CHANGE_AUTHOR_COLORS = Object.freeze([
  '#C00000',
  '#0070C0',
  '#00B050',
  '#7030A0',
  '#E97132',
  '#196B24',
  '#9E480E',
  '#525252',
] as const);
