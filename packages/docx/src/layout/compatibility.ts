import type { CompatibilityRule, DeepReadonly } from './types.js';

function requireText(value: string, field: string): void {
  if (value.trim() === '') throw new Error(`CompatibilityRule.${field} must not be empty`);
}

export function defineCompatibilityRule(rule: CompatibilityRule): DeepReadonly<CompatibilityRule> {
  requireText(rule.id, 'id');
  requireText(rule.description, 'description');
  if (rule.evidence.kind === 'microsoft-note'
    || rule.evidence.kind === 'regression-test') {
    requireText(rule.evidence.reference, 'evidence.reference');
  } else {
    requireText(rule.evidence.syntheticFixtureId, 'evidence.syntheticFixtureId');
    requireText(rule.evidence.application, 'evidence.application');
    requireText(rule.evidence.version, 'evidence.version');
    requireText(rule.evidence.platform, 'evidence.platform');
  }
  Object.freeze(rule.evidence);
  return Object.freeze(rule);
}

export const WORD_SECTION_BTLR_TBRL_PAGE_FRAME = defineCompatibilityRule({
  id: 'word-section-btlr-tbrl-page-frame',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s',
  },
  description: 'Issue #988 comment 4950296007 records that, unlike the normative ECMA-376 Part 4 §14.11.7 equivalence to lr, Word uses the tbRl page frame for section-level btLr; this rule covers only the page frame, while glyph orientation is paint-owned.',
});
