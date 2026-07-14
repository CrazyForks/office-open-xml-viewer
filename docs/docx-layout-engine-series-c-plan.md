# DOCX Layout Series C Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Finish compatibility, font, diagnostics, conformance, and cleanup infrastructure so the DOCX renderer has one auditable production algorithm.

**Architecture:** Floats and page-dependent content converge through explicit fingerprints; font resolution/shaping is an injected service shared by main and worker; unsupported facts become diagnostics; a generated public corpus proves geometry. The final PR makes `renderer.ts` a thin compatibility adapter and statically proves no legacy path remains.

**Tech Stack:** TypeScript, Rust, Canvas 2D, FontFace/worker FontFaceSet, Vitest, Playwright, ast-grep, pnpm, GitHub Actions.

## Global Constraints

- Follow all constraints and the per-PR independent review gate in `docx-layout-engine-implementation-roadmap.md`.
- Compatibility rules require a normative citation, Microsoft implementation note, or documented synthetic Office observation.
- Non-convergence is an error diagnostic; it never returns stale or overlapping geometry.
- Generated fixtures must be redistributable and contain no private content.

---

### Task C1: Express float placement as explicit constraints and convergence

**Files:**

- Create: `packages/docx/src/layout/floats.ts`
- Create: `packages/docx/src/layout/floats.test.ts`
- Create: `packages/docx/src/layout/compatibility.ts`
- Create: `packages/docx/src/layout/compatibility.test.ts`
- Modify: `packages/docx/src/layout/paginator.ts`
- Modify: `packages/docx/src/layout/diagnostics.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/float-line-start-one-inch.test.ts`
- Modify: `packages/docx/src/float-table-geometry.test.ts`

**Interfaces:**

- Consumes: placed nodes, page/column bounds, wrap geometry, and A1 diagnostics/fingerprints.
- Produces: `FloatConstraint`, `FloatPlacement`, `solveFloatPlacement`, and `convergeLayout`.

```ts
export interface FloatConstraint { readonly anchor: SourceRef; readonly horizontal: AxisConstraint; readonly vertical: AxisConstraint; readonly wrap: WrapConstraint; readonly allowOverlap: boolean; readonly layoutInCell: boolean }
export interface FloatPlacement { readonly bounds: LayoutRect; readonly exclusion: readonly PointPt[]; readonly pageIndex: number; readonly columnIndex: number }
export function solveFloatPlacement(input: FloatSolveInput): FloatPlacement;
export function convergeLayout(seed: LayoutIteration, step: (iteration: LayoutIteration) => LayoutIteration, limit: number): LayoutIteration;
```

- [ ] **Step 1: Add failing constraint and convergence tests**

Cover page/margin/column/character anchors, align versus offset precedence,
square/tight/through/top-bottom wrap, `allowOverlap`, `layoutInCell`, negative
offsets, multiple interacting floats, and page-dependent fields. Assert stable
placement fingerprints, repeated-fingerprint cycle detection, and
`NON_CONVERGENCE` when `limit` is reached.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/floats.test.ts packages/docx/src/layout/compatibility.test.ts packages/docx/src/float-line-start-one-inch.test.ts packages/docx/src/float-table-geometry.test.ts
```

Expected: float placement mutates renderer state and lacks explicit cycle/error
contracts.

- [ ] **Step 3: Implement the pure solver and isolate compatibility**

Translate OOXML anchor facts to axis/wrap constraints, solve against immutable
page/container exclusions, and return placement plus exclusion polygon. Iterate
page fields, note reserve, and interacting floats until the normalized relevant
geometry fingerprint is stable. Track seen fingerprints in a `Set<string>` and
throw a diagnostic error on a cycle or safety limit.

Move only evidenced Office-specific behavior into `compatibility.ts`, with a
named function, evidence comment, and synthetic test per rule. Delete mutable
float placement/retry logic from `renderer.ts`.

- [ ] **Step 4: Verify Green and deterministic failure**

Run the command from Step 2 plus:

```bash
pnpm vitest run packages/docx/src/layout/paginator.test.ts packages/docx/src/layout/invariants.test.ts
pnpm typecheck
```

Expected: tests pass; identical input produces identical float fingerprints; an
oscillating fixture fails with `NON_CONVERGENCE` rather than stale geometry.

- [ ] **Step 5: Commit, independently review, fix, and merge PR C1**

Commit subject: `refactor(docx): solve floating layout as constraints`.
Use the roadmap review gate.

### Task C2: Unify main and worker font resolution and shaping services

**Files:**

- Modify: `packages/docx/src/layout/text.ts`
- Create: `packages/docx/src/layout/font-service.ts`
- Create: `packages/docx/src/layout/font-service.test.ts`
- Modify: `packages/docx/src/local-font-metrics.ts`
- Modify: `packages/docx/src/embedded-fonts.ts`
- Modify: `packages/docx/src/google-fonts.ts`
- Modify: `packages/docx/src/document.ts`
- Modify: `packages/docx/src/render-worker.ts`
- Modify: `packages/docx/src/renderer.ts`
- Inspect and reuse where coherent: `packages/core/src` font and text helpers

**Interfaces:**

- Consumes: authored run fonts, theme/script resolution, embedded/local/substitute resources, and environment FontFaceSet.
- Produces: `FontResolver`, `FontResolution`, and `createTextLayoutService` shared by main and worker.

```ts
export interface FontResolution { readonly requestedFamily: string; readonly resolvedFamily: string; readonly source: 'embedded' | 'local' | 'google' | 'substitute' | 'generic'; readonly weight: number; readonly style: 'normal' | 'italic'; readonly diagnostics: readonly LayoutDiagnostic[] }
export interface FontResolver { resolve(request: Readonly<FontRequest>): FontResolution }
export function createTextLayoutService(resolver: FontResolver, measurer: GlyphMeasurer): TextLayoutService;
```

- [ ] **Step 1: Add failing resolution and parity tests**

Test ASCII, East Asian, complex-script, theme, embedded, local, Google, missing,
bold, and italic choices. Supply identical fake font inventories to main and
worker service factories and assert identical resolution records, shaped glyph
advances, paragraph line ranges, and layout fingerprints.

- [ ] **Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/font-service.test.ts packages/docx/src/layout/worker-parity.test.ts packages/docx/src/layout-lines-scale-invariance.test.ts
```

Expected: font preload and resolved-local-metric state are wired separately and
paint-scale paths can still influence glyph metrics.

- [ ] **Step 3: Resolve once before layout and inject one text service**

Register font resources, snapshot available families into an immutable resolver,
and create the same `TextLayoutService` in main and worker. Store `FontResolution`
and shaped glyph geometry on text placements. Adapt reusable core font inventory
and Canvas primitives only where their semantics are format-neutral; keep DOCX
theme/script fallback in DOCX.

Delete paint-time font selection, `rescaleLayoutLines`, and module-global
document-specific resolved-font state. Painting may set the already resolved CSS
font descriptor but cannot choose a substitute or measure it.

- [ ] **Step 4: Verify Green under scale and worker changes**

Run the command from Step 2 plus:

```bash
rg -n 'rescaleLayoutLines|setResolvedLocalFonts|clearResolvedLocalFonts' packages/docx/src --glob '!**/*.test.ts'
pnpm typecheck
```

Expected: tests pass and production matches are absent or replaced by the new
instance-scoped service names; scale/DPR do not alter layout fingerprints.

- [ ] **Step 5: Commit, independently review, fix, and merge PR C2**

Commit subject: `refactor(docx): share font shaping across render modes`.
Use the roadmap review gate.

### Task C3: Propagate diagnostics and add a synthetic conformance corpus

**Files:**

- Modify: `packages/docx/parser/src/types.rs`
- Modify: `packages/docx/parser/src/parser.rs`
- Create: `packages/docx/parser/tests/diagnostics.rs`
- Modify: `packages/docx/src/types.ts`
- Modify: `packages/docx/src/layout/diagnostics.ts`
- Create: `packages/docx/src/layout/diagnostics.test.ts`
- Create: `packages/docx/tests/conformance/generate.ts`
- Create: `packages/docx/tests/conformance/cases.ts`
- Create: `packages/docx/tests/conformance/layout.spec.ts`
- Create: `packages/docx/tests/conformance/browser.spec.ts`
- Modify: `packages/docx/playwright.config.ts`
- Modify: `.github/workflows/ci.yml`

**Interfaces:**

- Consumes: parser-preserved unsupported/invalid facts and final layout fingerprints.
- Produces: serialized `ParseDiagnostic`, mapped `LayoutDiagnostic`, generated minimal DOCX fixtures, and browser geometry assertions.

```rust
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParseDiagnostic {
    pub code: String,
    pub severity: DiagnosticSeverity,
    pub part: String,
    pub path: Vec<usize>,
}
```

- [ ] **Step 1: Add failing parser-to-layout diagnostic tests**

Build minimal OOXML for a recognized unsupported decoration, invalid geometry,
unknown enum value, and a supported control case. Assert stable codes and source
paths; assert layout maps recoverable cases to warnings, fatal geometry to an
error, and the supported case to no diagnostic.

- [ ] **Step 2: Add failing generated-corpus geometry tests**

Generate redistributable pairwise cases spanning story, container, paragraph,
table, nested table, inline/floating object, direction, spacing, style source,
font source, and anchor reference. Assert page count, line ranges, non-overlap,
bottom-margin clearance, stable fingerprints, and main/worker parity in Vitest;
assert the same normalized geometry in Chromium, Firefox, and WebKit.

- [ ] **Step 3: Run tests to verify Red**

Run:

```bash
cargo test -p docx-parser diagnostics -- --nocapture
pnpm vitest run packages/docx/src/layout/diagnostics.test.ts packages/docx/tests/conformance/layout.spec.ts
pnpm playwright test --config packages/docx/playwright.config.ts conformance/browser.spec.ts --project=chrome
```

Expected: diagnostic fields and generated fixture modules do not exist.

- [ ] **Step 4: Preserve diagnostics and generate fixtures deterministically**

Record stable parser codes without document text, map them at the parser/layout
boundary, and include them in `DocumentLayout.diagnostics`. Implement a deterministic
ZIP/XML generator using repository dependencies, with fixed timestamps and IDs,
so generated bytes and expected fingerprints are stable. CI runs node geometry on
every change and all three browser projects on the existing browser-test cadence.

- [ ] **Step 5: Rebuild and verify Green**

Run:

```bash
pnpm build:wasm
cargo test -p docx-parser
pnpm vitest run packages/docx/src/layout/diagnostics.test.ts packages/docx/tests/conformance/layout.spec.ts
pnpm playwright test --config packages/docx/playwright.config.ts conformance/browser.spec.ts
pnpm typecheck
```

Expected: all checks pass and two consecutive corpus generations have identical
hashes and normalized fingerprints.

- [ ] **Step 6: Commit, independently review, fix, and merge PR C3**

Commit subject: `test(docx): add layout diagnostics and conformance corpus`.
Use the roadmap review gate.

### Task C4: Reduce renderer to an adapter and prove architectural completion

**Files:**

- Create: `packages/docx/src/paint/canvas-page.test.ts`
- Modify: `packages/docx/src/paint/canvas-page.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/document.ts`
- Delete when unreferenced: `packages/docx/src/layout-context.ts`
- Delete when unreferenced: `packages/docx/src/paragraph-measure.ts`
- Delete when unreferenced: `packages/docx/src/layout-fragments.ts`
- Delete when unreferenced: `packages/docx/src/table-fragments.ts`
- Create: `rules/no-docx-runtime-layout-stamps.yml`
- Create: `rules/no-docx-migration-flags.yml`
- Create: `packages/docx/src/layout/architecture.test.ts`
- Modify: `sgconfig.yml`
- Modify: `.github/workflows/ci.yml`
- Modify: `docs/docx-layout-engine-redesign.md`

**Interfaces:**

- Consumes: final `layoutDocument`, `paintLayoutPage`, and layout-derived metadata.
- Produces: existing `paginateDocument` and `renderDocumentToCanvas` compatibility entry points as thin adapters with unchanged public callers.

```ts
export function paginateDocument(document: Readonly<DocxDocumentModel>): DocumentLayout;
export async function renderDocumentToCanvas(document: Readonly<DocxDocumentModel>, target: HTMLCanvasElement | OffscreenCanvas, pageIndex: number, options: RenderPageOptions): Promise<void>;
```

- [ ] **Step 1: Write failing architecture tests and static rules**

Assert deep-frozen parser and layout inputs survive layout, repeated paint, failed
image paint, and search projection unchanged. Assert two layout calls with the
same services have identical fingerprints. Add static rules rejecting runtime
stamp property names, `*ReuseEnabled`, `*PaintEnabled`, `RequiresLegacy`,
`legacy` layout branches, `dryRun` layout, and layout/measurement declarations in
`renderer.ts`.

- [ ] **Step 2: Run tests and static scan to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/architecture.test.ts packages/docx/src/paint/canvas-page.test.ts
pnpm lint
```

Expected: current renderer still contains mutable render/layout state and legacy
layout declarations, so the new assertions/rules fail.

- [ ] **Step 3: Complete adapter extraction and delete transitional modules**

Keep only resource acquisition, public-option normalization, service creation,
layout invocation/cache ownership, canvas sizing, and paint invocation in
`renderer.ts`. Move no algorithm into another compatibility wrapper. Delete each
transitional module only after `rg` proves no import remains. Update the design
status to implemented and record the final module dependency direction.

- [ ] **Step 4: Run focused architectural proof**

Run:

```bash
pnpm vitest run packages/docx/src/layout/architecture.test.ts packages/docx/src/paint/canvas-page.test.ts packages/docx/src/layout/invariants.test.ts packages/docx/src/layout/worker-parity.test.ts
pnpm lint
rg -n 'fitMeasureReuseEnabled|fragmentPaintEnabled|lineReuseEnabled|tableReuseEnabled|RequiresLegacy|requiresLegacy|dryRun|PaginatedBodyElement|tableColWidthsPt|tableRowHeightsPt|layoutLinesInputs|deferFront' packages/docx/src --glob '!**/*.test.ts'
```

Expected: tests and static scan pass and `rg` has no production matches.

- [ ] **Step 5: Run the final broad verification and API diff**

Run:

```bash
pnpm build:wasm
pnpm test
pnpm typecheck
pnpm build-storybook
pnpm playwright test --config packages/docx/playwright.config.ts conformance/browser.spec.ts
cargo test -p docx-parser
git diff main...HEAD -- packages/docx/package.json packages/docx/src/index.ts packages/docx/src/document.ts packages/docx/src/types.ts
git diff --check
```

Expected: all verification passes; the public diff contains no removed or
incompatibly changed exports/signatures; no private artifact is tracked.

- [ ] **Step 6: Obtain an independent final architecture audit**

Use the roadmap review brief and additionally require the reviewer to prove each
release-gate claim with a command or exact code reference. Expected: one production
algorithm per feature class, no paint measurement, no parser stamps, no migration
flags/fallback, all stories in layout, main/worker parity, and compatible public APIs.

- [ ] **Step 7: Fix findings, reverify, commit, and merge PR C4**

Commit subject: `refactor(docx): complete immutable layout pipeline`.
Repeat Steps 4–6 after material fixes, then merge with `gh pr merge <number> --merge`.
Close Issue #1037 only after GitHub shows PR C4 merged. Do not create a release or tag.
