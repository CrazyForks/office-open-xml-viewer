import { defineCompatibilityRule } from './compatibility.js';

export const ESTABLISHED_NEXT_COLUMN_PAGE_ADVANCE = defineCompatibilityRule({
  id: 'established-next-column-page-advance',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/paginator.test.ts#advances nextColumn to the next page when the outgoing column has no same-page successor',
  },
  description: 'When §17.18.77 has no following column on the current page, continue in the incoming section on the following page, preserving the established pre-cutover non-continuous section behavior instead of rejecting the document.',
});
