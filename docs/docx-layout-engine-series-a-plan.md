# DOCX Layout Series A Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make body paragraphs, tables, and page/column flow produce one immutable `DocumentLayout` with no parser-model stamps or legacy paint route.

**Architecture:** Introduce stable result and service contracts first, then migrate paragraphs, in-flow/nested tables, floating/split tables, and finally the page state machine. Every paint module consumes self-contained geometry and resources and cannot import measurement code.

**Tech Stack:** TypeScript, Canvas 2D, Vitest, ast-grep, pnpm, Rust-generated DOCX models.

## Global Constraints

- Follow all constraints and the per-PR independent review gate in `docx-layout-engine-implementation-roadmap.md`.
- Preserve public API signatures and render behavior.
- Do not retain a migrated feature's old algorithm behind a flag or predicate.
- Layout code cannot read scale, DPR, Canvas state, or paint callbacks.
- Paint code cannot measure, shape, paginate, resolve styles, or dereference parser objects.

---

### Task A1: Establish immutable layout, diagnostics, invariants, and paint purity

**Files:**

- Create: `packages/docx/src/layout/types.ts`
- Create: `packages/docx/src/layout/diagnostics.ts`
- Create: `packages/docx/src/layout/invariants.ts`
- Create: `packages/docx/src/layout/invariants.test.ts`
- Create: `packages/docx/src/paint/paint-purity.test.ts`
- Create: `rules/no-docx-layout-in-paint.yml`
- Create: `rules/no-docx-display-scale-in-layout.yml`
- Modify: `sgconfig.yml`
- Modify: `.github/workflows/ci.yml`
- Delete: `rules/no-docx-measurement-in-fragment-paint.yml`

**Interfaces:**

- Consumes: `SectionLayoutContext` from `packages/docx/src/layout-context.ts` and parser model types from `types.ts` only at the layout boundary.
- Produces: `DocumentLayout`, `LayoutPage`, `PageLayers`, `PaintNode`, `LayoutDiagnostic`, `assertDocumentLayout`, and `layoutFingerprint`.

- [ ] **Step 1: Write failing invariant and paint-purity tests**

Add tests that construct two body nodes with overlapping page rectangles, a body
node crossing `geometry.contentBottomPt`, a NaN coordinate, and a Canvas stub
whose `measureText()` throws. Assert `assertDocumentLayout` throws diagnostic
codes `FLOW_OVERLAP`, `BOTTOM_MARGIN_INVASION`, and `INVALID_GEOMETRY`, while a
minimal paint-only page succeeds without calling `measureText`.

```ts
expect(() => assertDocumentLayout(overlappingLayout)).toThrow(/FLOW_OVERLAP/);
expect(() => assertDocumentLayout(marginLayout)).toThrow(/BOTTOM_MARGIN_INVASION/);
expect(() => assertDocumentLayout(nanLayout)).toThrow(/INVALID_GEOMETRY/);
expect(() => paintLayoutPage(paintOnlyLayout, 0, canvas, { scale: 1, dpr: 1 }))
  .not.toThrow();
```

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/invariants.test.ts packages/docx/src/paint/paint-purity.test.ts
```

Expected: FAIL because the new modules and exports do not exist.

- [ ] **Step 3: Add the minimal immutable contracts**

Define readonly point-space rectangles and diagnostics:

```ts
export interface LayoutRect { readonly xPt: number; readonly yPt: number; readonly widthPt: number; readonly heightPt: number }
export interface PageGeometry extends LayoutRect { readonly contentTopPt: number; readonly contentBottomPt: number }
export type LayoutDiagnosticCode = 'FLOW_OVERLAP' | 'BOTTOM_MARGIN_INVASION' | 'INVALID_GEOMETRY' | 'NON_CONVERGENCE' | 'UNSUPPORTED_FEATURE';
export interface LayoutDiagnostic { readonly code: LayoutDiagnosticCode; readonly severity: 'warning' | 'error'; readonly source?: SourceRef; readonly message: string }
export function assertDocumentLayout(layout: DocumentLayout): void;
export function layoutFingerprint(layout: DocumentLayout): string;
```

Implement `layoutFingerprint` by recursively normalizing finite numbers to six
decimal places and serializing pages, layers, and reading order; omit diagnostics'
free-form message but include code, severity, and source.

Replace the single-file ast-grep rule with two directory rules: reject imports
from `../layout/`, `paragraph-measure`, `layout-context`, or `renderer` and calls
to `measureText`, `layoutLines`, `computeTableLayout`, or `paginateDocument` in
`packages/docx/src/paint/**/*.ts`; reject identifiers `scale`, `dpr`,
`CanvasRenderingContext2D`, and `OffscreenCanvasRenderingContext2D` in
`packages/docx/src/layout/**/*.ts` except test files.

- [ ] **Step 4: Run focused and static checks**

Run:

```bash
pnpm vitest run packages/docx/src/layout/invariants.test.ts packages/docx/src/paint/paint-purity.test.ts
pnpm lint
pnpm typecheck
```

Expected: all commands pass; intentionally adding `ctx.measureText('x')` to a
paint fixture makes ast-grep fail, and removing it restores green.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A1**

Commit subject: `refactor(docx): establish immutable layout boundaries`.
Use the roadmap review gate and merge only with all checks and findings clear.

### Task A2: Route every body paragraph through self-contained paragraph layout

**Files:**

- Create: `packages/docx/src/layout/text.ts`
- Create: `packages/docx/src/layout/paragraph.ts`
- Create: `packages/docx/src/layout/paragraph.test.ts`
- Create: `packages/docx/src/paint/canvas-text.ts`
- Create: `packages/docx/src/paint/canvas-text.test.ts`
- Modify: `packages/docx/src/layout/types.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/layout-fragments.ts`
- Delete: `packages/docx/src/fragment-paint.ts`
- Modify: `packages/docx/src/fragment-paint.test.ts`
- Modify: `packages/docx/src/layout-lines-reuse-identity.test.ts`
- Modify: `packages/docx/src/layout-lines-scale-invariance.test.ts`
- Modify: `packages/docx/src/layout-lines-zoom-invariant.test.ts`

**Interfaces:**

- Consumes: `TextLayoutService`, `SourceRef`, and invariant contracts from A1.
- Produces: `ParagraphLayout`, `TextPlacement`, `layoutParagraph`, and `paintParagraphLayout`.

```ts
export interface TextLayoutService {
  shape(request: Readonly<TextShapeRequest>): TextShapeResult;
}
export interface ParagraphLayout {
  readonly kind: 'paragraph';
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  readonly bounds: LayoutRect;
  readonly advancePt: number;
  readonly lines: readonly LineLayout[];
  readonly borders: readonly BorderSegment[];
  readonly shading?: FillPaint;
}
export function layoutParagraph(input: ParagraphLayoutInput, services: LayoutServices): ParagraphLayout;
export function paintParagraphLayout(node: ParagraphLayout, context: CanvasPaintContext): void;
```

- [ ] **Step 1: Add failing behavior tests**

Add synthetic tests for numbered paragraphs, bidi runs, vertical text, tab
leaders, floating-wrap exclusions, continuation slices, contextual spacing,
paragraph borders, hidden paragraph marks, and page fields. For each case assert
the exact line text ranges and point bounds, `advancePt === bounds.heightPt`, and
that two paints at scale 1 and 2 do not call `measureText` or change the layout
fingerprint.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/paragraph.test.ts packages/docx/src/paint/canvas-text.test.ts packages/docx/src/layout-lines-reuse-identity.test.ts packages/docx/src/layout-lines-scale-invariance.test.ts packages/docx/src/layout-lines-zoom-invariant.test.ts
```

Expected: new contract tests fail because paint still delegates to
`renderBodyParagraphLines` and scale-2 paint remeasures.

- [ ] **Step 3: Move line acquisition and glyph geometry into layout**

Adapt existing `buildSegments`, bidi/tab resolution, `layoutLines`, numbering,
field resolution, and paragraph decoration calculations into `layout/text.ts`
and `layout/paragraph.ts`. Store resolved glyph text, font descriptor, advances,
offsets, decorations, link/bookmark metadata, and resource keys on
`TextPlacement`. `canvas-text.ts` only applies `CanvasPaintContext.transform` and
calls drawing primitives.

Delete `fitMeasureReuseEnabled`, `fragmentPaintEnabled`,
`lineReuseEnabled`, `isFragmentPaintableParagraph`, `layoutLinesInputs`, and
`stampParagraphLines`. Remove `source: DocParagraph` and `MeasuredParagraph` from
paint-facing fragments; retain only `SourceRef` and self-contained paint data.

- [ ] **Step 4: Verify Green and prove deletion**

Run:

```bash
pnpm vitest run packages/docx/src/{layout/paragraph,paint/canvas-text,fragment-paint,layout-lines-reuse-identity,layout-lines-scale-invariance,layout-lines-zoom-invariant}.test.ts
rg -n 'fitMeasureReuseEnabled|fragmentPaintEnabled|lineReuseEnabled|isFragmentPaintableParagraph|layoutLinesInputs|stampParagraphLines|renderBodyParagraphLines' packages/docx/src
pnpm typecheck
```

Expected: tests and typecheck pass; `rg` has no production matches.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A2**

Commit subject: `refactor(docx): make paragraph paint consume layout geometry`.
Use the roadmap review gate.

### Task A3: Build in-flow and nested table geometry from one measurement

**Files:**

- Create: `packages/docx/src/layout/table.ts`
- Create: `packages/docx/src/layout/table.test.ts`
- Create: `packages/docx/src/paint/canvas-table.ts`
- Create: `packages/docx/src/paint/canvas-table.test.ts`
- Modify: `packages/docx/src/layout/types.ts`
- Modify: `packages/docx/src/table-fragments.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/table-layout-reuse.test.ts`
- Modify: `packages/docx/src/cell-border-conflict-render.test.ts`
- Modify: `packages/docx/src/column-widths.test.ts`

**Interfaces:**

- Consumes: `layoutParagraph` and recursive `layoutFlowBlocks` supplied by the container coordinator.
- Produces: `TableLayout`, `TableRowLayout`, `TableCellLayout`, `ResolvedBorderSegment`, `layoutTable`, and `paintTableLayout`.

```ts
export interface TableLayout {
  readonly kind: 'table';
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  readonly bounds: LayoutRect;
  readonly advancePt: number;
  readonly columnWidthsPt: readonly number[];
  readonly rows: readonly TableRowLayout[];
  readonly borders: readonly ResolvedBorderSegment[];
}
export function layoutTable(input: TableLayoutInput, services: LayoutServices): TableLayout;
export function paintTableLayout(node: TableLayout, context: CanvasPaintContext): void;
```

- [ ] **Step 1: Write failing single-acquisition tests**

Create a counting `TextLayoutService` and synthetic fixed/auto tables containing
paragraphs, nested tables, vertical merges, row spans, exact/at-least heights,
cell margins, and conflicting borders. Assert each paragraph is shaped once per
placement, row heights equal the sum/max of retained child layouts, and paint
does not increment the counter.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/table.test.ts packages/docx/src/paint/canvas-table.test.ts packages/docx/src/table-layout-reuse.test.ts packages/docx/src/cell-border-conflict-render.test.ts packages/docx/src/column-widths.test.ts
```

Expected: counting assertions fail because `buildTableCellBlocks` performs a
second cell-content measurement and paint retains a legacy supplied-geometry
bridge.

- [ ] **Step 3: Implement one retained table acquisition**

Resolve the grid, lay out each cell's blocks once, compute intrinsic cell heights
from those retained blocks, resolve row/vMerge heights, translate child bounds to
final cell positions, and resolve shared border segments once. Recursively use
the same function for nested tables. `paintTableLayout` draws stored backgrounds,
children, clipping, and border segments only.

Remove `tableColWidthsPt`, `tableRowHeightsPt`, and `tableLayoutInputs` from
`PaginatedBodyElement`; delete their writes and reuse checks for in-flow and
nested tables. Delete the second paragraph acquisition in
`buildTableCellBlocks`; preserve a single function that converts retained child
layouts into page fragments.

- [ ] **Step 4: Verify Green and mutation safety**

Run:

```bash
pnpm vitest run packages/docx/src/layout/table.test.ts packages/docx/src/paint/canvas-table.test.ts packages/docx/src/table-layout-reuse.test.ts packages/docx/src/cell-border-conflict-render.test.ts packages/docx/src/column-widths.test.ts
rg -n 'tableColWidthsPt|tableRowHeightsPt|tableLayoutInputs' packages/docx/src --glob '!**/*.test.ts'
pnpm typecheck
```

Expected: tests pass, parser input remains deeply equal before/after layout and
paint, and `rg` has no production matches.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A3**

Commit subject: `refactor(docx): retain one table layout acquisition`.
Use the roadmap review gate.

### Task A4: Migrate floating and page-split tables without a legacy gate

**Files:**

- Modify: `packages/docx/src/layout/table.ts`
- Create: `packages/docx/src/layout/table-pagination.test.ts`
- Modify: `packages/docx/src/table-fragments.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/float-table-geometry.test.ts`
- Modify: `packages/docx/src/float-table-page-fit.test.ts`
- Modify: `packages/docx/src/float-table-width.test.ts`
- Modify: `packages/docx/src/pagination.test.ts`

**Interfaces:**

- Consumes: `TableLayout` from A3 and current float exclusion inputs.
- Produces: `TableContinuation`, `splitTableLayout`, and floating `TableLayout` placements using the same node type.

```ts
export interface TableContinuation { readonly rowStart: number; readonly rowEnd: number; readonly continuesFromPreviousPage: boolean; readonly continuesOnNextPage: boolean }
export function splitTableLayout(table: TableLayout, availableHeightPt: number): readonly TableLayout[];
```

- [ ] **Step 1: Write failing continuation and floating tests**

Cover repeated header rows, `cantSplit`, mid-cell paragraph continuation,
vertical merge continuation, nested table continuation, negative table indent,
floating table wrapping, and a float that must move to the next page. Assert row
ownership is disjoint and exhaustive, fragments reuse the same resolved columns,
and no fragment enters the bottom margin.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/table-pagination.test.ts packages/docx/src/float-table-geometry.test.ts packages/docx/src/float-table-page-fit.test.ts packages/docx/src/float-table-width.test.ts packages/docx/src/pagination.test.ts
```

Expected: legacy-gated floating cases do not produce `TableLayout` continuations.

- [ ] **Step 3: Implement splits as immutable views over retained geometry**

Split only at legal row/cell/line boundaries, repeat the resolved header layout,
and create fragment-local translated bounds without recomputing columns or text.
Represent floating placement as a `DrawingLayout`-style placement wrapper whose
child is the same `TableLayout` used for in-flow content.

Delete `tableRequiresLegacyPaint`, `isFragmentPaintableTable`,
`tableReuseEnabled`, `renderTableFragment`, and the legacy table-paint selection.
Leave one `layoutTable` and one `paintTableLayout` production route.

- [ ] **Step 4: Verify Green and prove one route**

Run:

```bash
pnpm vitest run packages/docx/src/layout/table-pagination.test.ts packages/docx/src/float-table-{geometry,page-fit,width}.test.ts packages/docx/src/pagination.test.ts
rg -n 'tableRequiresLegacyPaint|isFragmentPaintableTable|tableReuseEnabled|renderTableFragment' packages/docx/src --glob '!**/*.test.ts'
pnpm typecheck
```

Expected: all tests pass and `rg` has no production matches.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A4**

Commit subject: `refactor(docx): unify floating and split table layout`.
Use the roadmap review gate.

### Task A5: Extract the page/column state machine and establish worker parity

**Files:**

- Create: `packages/docx/src/layout/context.ts`
- Create: `packages/docx/src/layout/paginator.ts`
- Create: `packages/docx/src/layout/paginator.test.ts`
- Create: `packages/docx/src/layout/worker-parity.test.ts`
- Modify: `packages/docx/src/layout/types.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/types.ts`
- Modify: `packages/docx/src/render-worker.ts`
- Modify: `packages/docx/src/worker-protocol.ts`
- Modify: `packages/docx/src/document.ts`
- Modify: `packages/docx/src/bookmark-nav.ts`

**Interfaces:**

- Consumes: paragraph/table layout functions and A1 fingerprints.
- Produces: `PageFlowState`, `PageFlowEvent`, final `layoutDocument`, and worker-retained `DocumentLayout`.

```ts
export interface PageFlowState { readonly pageIndex: number; readonly columnIndex: number; readonly cursorYPt: number; readonly section: SectionLayoutContext }
export type PageFlowEvent =
  | Readonly<{ type: 'place'; node: PaintNode }>
  | Readonly<{ type: 'next-column' }>
  | Readonly<{ type: 'next-page'; reason: 'overflow' | 'explicit-break' | 'section-break' | 'parity' }>
  | Readonly<{ type: 'begin-section'; section: SectionLayoutContext }>;
export function paginateBody(input: BodyLayoutInput, services: LayoutServices): DocumentLayout;
```

- [ ] **Step 1: Add failing state-machine and parity tests**

Cover explicit page/column breaks, continuous and next-page sections, even/odd
parity pages, mixed page sizes, per-section vertical direction, multi-column
regions starting mid-page, keep-next, widow/orphan control, hidden paragraphs,
and bottom-margin overflow. Serialize the same synthetic document through direct
layout and the render-worker layout handler and assert identical fingerprints and
page sizes.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/paginator.test.ts packages/docx/src/layout/worker-parity.test.ts packages/docx/src/pagination.test.ts packages/docx/src/per-section-headers-footers.test.ts
```

Expected: the worker retains `PaginatedBodyElement[][]`, and section/page facts
are recovered from element stamps rather than `LayoutPage`.

- [ ] **Step 3: Implement explicit transitions and page ownership**

Make each transition return a new `PageFlowState`; store section, geometry,
columns, content origin, page numbering, direction, header/footer references,
and parity-page metadata on `LayoutPage`. Replace `computePages` closure state
with `paginateBody`. The render worker retains `DocumentLayout`; page metadata and
bookmarks derive from it. Keep worker protocol response shapes and public methods
unchanged.

Remove `sectionBreakSpacer`, `collapsedSpacer`, `leadsCollapsedRun`,
`hiddenCollapsed`, `colIndex`, `colGeom`, `colTopPt`, `sectionHF`,
`sectionGeom`, `sectionPageNumType`, and `sectionTextDirection` from
`PaginatedBodyElement`, then remove `PaginatedBodyElement` if no consumers remain.

- [ ] **Step 4: Verify Green and deletion**

Run:

```bash
pnpm vitest run packages/docx/src/layout/{paginator,worker-parity}.test.ts packages/docx/src/pagination.test.ts packages/docx/src/per-section-headers-footers.test.ts packages/docx/src/document-destroy.test.ts
rg -n 'sectionBreakSpacer|collapsedSpacer|leadsCollapsedRun|hiddenCollapsed|colGeom|colTopPt|sectionHF|sectionGeom|sectionPageNumType|sectionTextDirection|PaginatedBodyElement' packages/docx/src --glob '!**/*.test.ts'
pnpm typecheck
```

Expected: tests pass; no runtime stamp or `PaginatedBodyElement` production match
remains; normalized main and worker fingerprints are equal.

- [ ] **Step 5: Commit, independently review, fix, and merge PR A5**

Commit subject: `refactor(docx): make page flow an immutable state machine`.
Use the roadmap review gate.
