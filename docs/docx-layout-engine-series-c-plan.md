# DOCX Layout Series C Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Finish compatibility, diagnostics, conformance, and cleanup infrastructure so the DOCX renderer has one auditable production algorithm.

**Architecture:** Floats use explicit constraints and the stable Series A convergence/service foundation; unsupported facts become diagnostics; a generated public corpus proves geometry and invariants. The final PR makes `renderer.ts` a thin compatibility adapter and statically proves no legacy path remains.

**Tech Stack:** TypeScript, Rust, Canvas 2D, FontFace/worker FontFaceSet, Vitest, Playwright, ast-grep, pnpm, GitHub Actions.

## Global Constraints

- Follow all constraints and the per-PR independent review gate in `docx-layout-engine-implementation-roadmap.md`.
- Compatibility rules require a normative citation, Microsoft implementation note, or documented synthetic Office observation.
- Non-convergence is an error diagnostic; it never returns stale or overlapping geometry.
- Generated fixtures must be redistributable and contain no private content.

---

### Task C1: Express float placement as explicit constraints with isolated compatibility

> Design correction (2026-07-20): Series A/B already made
> `layout/anchor-frame.ts` the immutable authority for anchor axes, size, wrap
> geometry, and required CT_Anchor behavior. C1 must not add a second
> `AxisConstraint`/`WrapConstraint` dialect or claim page/column admission.
> C1 is split into C1a/C1b so displacement policy and real fixed points can be
> reviewed independently from deletion of the remaining legacy adapters.

**Files:**

- Create: `packages/docx/src/layout/floats.ts`
- Create: `packages/docx/src/layout/floats.test.ts`
- Modify: `packages/docx/src/layout/compatibility.ts`
- Modify: `packages/docx/src/layout/compatibility.test.ts`
- Modify: `packages/docx/src/layout/convergence.ts`
- Modify: `packages/docx/src/layout/line-wrap-convergence.ts`
- Modify: `packages/docx/src/layout/body-paginator.ts`
- Modify: `packages/docx/src/layout/paragraph.ts`
- Modify: `packages/docx/src/layout/floating-table-transaction.ts`
- Modify: `packages/docx/src/layout/table-pagination.ts`
- Modify: `packages/docx/src/layout/float-wrap.ts`
- Delete: `packages/docx/src/layout/repeated-state.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/float-line-start-one-inch.test.ts`
- Modify: `packages/docx/src/float-table-geometry.test.ts`
- Modify: `scripts/docx-layout-boundary-baseline.json`
- Modify: `scripts/check-docx-layout-boundaries.mjs`
- Modify: `scripts/check-docx-layout-boundaries.test.mjs`

**Interfaces:**

- Consumes: `anchor-frame` object/wrap geometry, immutable collision/float
  registries, containment boundaries, and the A2 convergence module.
- Produces: a pure displacement-policy result (`bounds`, `exclusionBounds`,
  `displacement`, and applied compatibility rule IDs). It does not resolve
  anchor axes, line-wrap windows, page/column admission, or paint.
- Extends the A2 convergence module with one generic exact-state primitive.
  Adjacent equality is a fixed point, non-adjacent equality is a cycle, and a
  hard pass budget is an explicit fail-closed resource guard rather than a
  geometry heuristic.

**Specification evidence:** ECMA-376 §20.4.2.3 (`wp:anchor`),
§20.4.2.10/§20.4.2.11 positioning, §20.4.2.15–§20.4.2.20 wrap
geometry, `allowOverlap`, and `layoutInCell` define constraints.
§17.4.56 defines floating-table overlap against table extents; §17.4.57
`*FromText` distances remain text-exclusion geometry. `[MS-OE376]` §2.1.474
defines the narrow Office `shapeLayoutLikeWW8` negative line-relative-offset
pagination behavior; its parser/internal-wire implementation belongs with C2
diagnostics and is not generalized into float geometry. Each rule in
`compatibility.ts` names its Microsoft note or documented regression evidence;
generic geometry contains no observation-derived constant.

```ts
export type FloatAvoidance =
  | { readonly kind: 'drawingml-normative' }
  | { readonly kind: 'word-different-paragraph'; readonly paragraphId: number }
  | { readonly kind: 'none' };
export type FloatPlacementParticipant =
  | (FloatParticipantCore & {
      readonly kind: 'table';
      readonly tableOverlap: 'never' | 'overlap';
    })
  | (FloatParticipantCore & { readonly kind: 'drawingml' })
  | (FloatParticipantCore & { readonly kind: 'frame' });
export interface FloatPlacement {
  readonly bounds: LayoutRect;
  readonly exclusionBounds: LayoutRect;
  readonly displacement: Readonly<{ xPt: number; yPt: number }>;
  readonly appliedCompatibilityRuleIds: readonly string[];
}
export function resolveFloatPlacement(input: ResolveFloatPlacementInput): FloatPlacement;
```

- [x] **C1a Step 1: Add failing policy and real-convergence tests**

Do not duplicate the existing `anchor-frame.test.ts` axis/wrap matrix. Cover
DrawingML-only normative blockers, table-only §17.4.56 blockers using raw table
extents, cell/page right boundaries, same- versus different-paragraph
compatibility, stable blocker-order results, and preservation of text-exclusion
padding. Add exact-state tests for adjacent fixed points, non-adjacent cycles,
and all-distinct resource-budget exhaustion. Pin the actual paragraph-anchor,
line-wrap, and floating-table final-frame fixed points.

- [x] **C1a Step 2: Run tests to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/floats.test.ts packages/docx/src/layout/compatibility.test.ts packages/docx/src/float-line-start-one-inch.test.ts packages/docx/src/float-table-geometry.test.ts
```

Expected: displacement policy is duplicated, `repeated-state.ts` is a second
unbounded convergence mechanism, all-distinct line/anchor states can run
without a typed limit failure, and table collision includes text-only padding.

- [x] **C1a Step 3: Implement policy extraction and bounded exact convergence**

Keep axis/wrap resolution in `anchor-frame.ts`. Select normative versus
compatibility blocker geometry in `floats.ts`, delegate only direction/math to
`axis-aligned-overlap.ts`, and return both object and exclusion translations.
Extend `convergence.ts` with the generic exact-state driver and migrate paragraph
anchor reflow, line-wrap measure/resolve, and floating-table final-frame reflow.
Delete `repeated-state.ts`.

Move only evidenced Office-specific behavior into `compatibility.ts`, with a
named record and test per rule. The issue #676 one-inch square threshold and
different-paragraph displacement are compatibility rules. Normative
`allowOverlap=false`, `tblOverlap=never`, and `layoutInCell` remain typed
constraints rather than compatibility records.

- [x] **C1a Step 4: Verify Green and deterministic failure**

Run the command from Step 2 plus:

```bash
pnpm vitest run packages/docx/src/layout/paginator.test.ts packages/docx/src/layout/invariants.test.ts
pnpm typecheck
```

Expected: tests pass; identical input produces identical placement and layout
fingerprints; a cycle and an all-distinct orbit both fail with
`NON_CONVERGENCE`; no last candidate is returned.

- [x] **C1a Step 5: Commit, independently review, fix, and merge**

Commit subject: `refactor(docx): centralize float placement policy`.
Use the roadmap review gate.

- [x] **C1b Step 1: Consolidate remaining registry callers and pairwise facts**

Add the blocker-side `tblOverlap=never` fact required by §17.4.56's pairwise
"either table" rule as a required discriminated-union field. Route
renderer/table/frame adapters through the same typed policy and retain the
established display/point numerical policies explicitly.

Extract ordinary-table retry as block-flow admission against §17.4.57 exclusion
geometry. It is a finite monotone sweep, not float-to-float placement and not a
non-monotone convergence problem. Migrate the two concrete remaining fixed
points — page-owned-anchor destination planning and floating-table parent/child
final-frame reflow — to the shared exact-state convergence primitive. The old
"selected-page ownership" wording was stale; selected-page ownership is now
pure and loop-free.

- [x] **C1b Step 2: Delete duplicated policy and ratchet boundaries**

Delete the scalar overlap adapter and inline blocker filters. Keep the mutable
registry only as a transport bridge tracked for C3; it must no longer choose
blocker classes, geometry, direction, tolerance, or retries. Isolate the
existing absolute floating-table page-deferral observation behind a named
compatibility record. Add static/mutation checks that float-specific
observation-derived rules are declared in `compatibility.ts`, generic numerical
policies retain their exact audited owners/values, and the displacement kernel
cannot be reached through named/namespace/default imports, re-exports, dynamic
imports, CommonJS, or string-key laundering.

- [x] **C1b Step 3: Review, broad-verify, and merge**

Commit subject: `refactor(docx): unify float placement callers`.
Use the roadmap review gate.

### Task C2a: Isolate Office compatibility evidence

> Design correction (2026-07-20): issue #1037 tracks compatibility isolation,
> parser/layout diagnostics, and the conformance corpus as three independently
> reviewable outcomes. C1 isolated float-specific rules but did not inventory
> the remaining layout and paint observations. Complete C2a before diagnostics
> so C2b can distinguish an unsupported fact from a named compatibility rule
> and C2c can cite the correct rule identity.

**Files:**

- Modify: `packages/docx/src/layout/compatibility.ts`
- Modify: `packages/docx/src/layout/compatibility.test.ts`
- Create: `packages/docx/src/layout/anchor-compatibility.ts`
- Create: `packages/docx/src/layout/section-compatibility.ts`
- Create: `packages/docx/src/layout/table-compatibility.ts`
- Modify: owning layout and paint modules only to delegate existing decisions
- Create: `scripts/check-docx-compatibility-evidence.mjs`
- Create: `scripts/check-docx-compatibility-evidence.test.mjs`
- Create: `scripts/docx-compatibility-microsoft-evidence.json`
- Create: `scripts/docx-compatibility-observation-baseline.json`
- Modify: `.github/workflows/ci.yml`

**Classification boundary:** A compatibility rule requires a Microsoft
implementation note, a live regression-test reference, or a generated Office
observation fixture. Normative OOXML behavior, deterministic solver direction,
numeric tolerances, convergence/resource limits, and public API compatibility
are not compatibility rules. Missing evidence remains unsupported/unresolved
and belongs to C2b diagnostics; it is never promoted to an Office claim.

- [x] **C2a-1: Isolate layout-side compatibility decisions**

Add one generic immutable factory, module-local rule records, stable unique IDs,
and pure decision helpers. Mechanically verify that every regression reference
resolves to a live test title, every Microsoft-note section exists in the
reviewed evidence catalog, rule IDs are globally unique, and rule declarations
occur only in the reviewed compatibility modules. Reject aliased, namespace,
re-exported, dynamic, CommonJS, and indirect-binding access to the factory, and
scan every production-importable TypeScript/JavaScript module shape. Add a
shrinking observation-comment baseline so new inline Office claims fail CI.
The lexical observation scan is a one-way migration ratchet, not a semantic
substitute for evidence review. Preserve every existing value and branch result.

Do not register the retained `nextColumn` no-successor transition: absent an
Office observation it is deterministic solver policy. Likewise, retained
section-band column-separator geometry remains a layout ownership invariant,
not an Office claim. Non-zero table-cell spacing remains normative OOXML;
only the `[MS-OI29500]` inside-border conflict deviation is compatibility-owned.
The Microsoft catalog is pinned to the reviewed published revision and section
titles; the specification files themselves remain local and uncommitted.

Runtime applied-rule tracing is not a C2a-1 acceptance condition. If C2b needs
it, propagate immutable rule IDs on retained result values and aggregate only
from the accepted final layout tree. Never introduce a mutable collector whose
contents can retain rejected probes or convergence candidates.

- [ ] **C2a-2: Isolate paint-side and supported model-boundary compatibility decisions**

Name and evidence the remaining observations in `renderer.ts`, its legacy paint
helpers, and the supported table-model boundary in `parser-model.ts` / `types.ts`
without performing the C3 extraction or C2b unsupported-content diagnostics.
Move only decision data and pure predicates; geometry, paint ownership, and the
public model remain unchanged. Reconcile legacy comments that claim Office
behavior with retained owners that deliberately treat the same policy as
implementation-defined. Expand the observation scanner to cover the remaining
Office-claim verbs, require its baseline to match exactly, and shrink that
baseline in the same PR.

Private-sample observations use public behavior-pinning regression tests as
evidence; private filenames and content never enter committed rule records.
Normative OOXML/UAX behavior, deterministic renderer policy, and deliberately
unsupported Office behavior are classified explicitly instead of being promoted
to compatibility rules. In particular, keep the draw-only #990 baseline rule
distinct from the #981 trailing-mark pagination-admission rule.

The exact transitional observation baseline may retain the encryption note and
the pre-existing nested-table marker comment inside
`measureCellContentHeightPx`. The latter declaration is byte-frozen by the
layout-boundary baseline, so C2a records the behavior as
`word-trailing-structural-cell-marker` without rewriting that one legacy
comment. C3 removes the frozen declaration and this final transitional entry;
adding any other inline observation remains a CI failure.

- [ ] **C2a-3: Review, broad-verify, and merge**

Use independent specification/evidence review for each C2a PR. Mark issue
#1037's compatibility-isolation item complete only after both layout and paint
inventories merge.

### Task C2b: Propagate parser diagnostics

**Files:**

- Modify: `packages/docx/parser/src/types.rs`
- Modify: `packages/docx/parser/src/parser.rs`
- Create: `packages/docx/parser/tests/diagnostics.rs`
- Modify: `packages/docx/src/parser-model.ts`
- Modify: `packages/docx/src/layout/diagnostics.ts`
- Create: `packages/docx/src/layout/diagnostics.test.ts`

**Interfaces:**

- Consumes: parser-preserved unsupported/invalid facts and final layout fingerprints.
- Produces: serialized `ParseDiagnostic` and mapped `LayoutDiagnostic`.

**Specification evidence:** Parser diagnostics distinguish schema-recognized
unsupported content, invalid values, and compatibility observations. A
diagnostic never includes document text or private source content and never
changes the exported TypeScript model.

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

Rust adds `diagnostics: Vec<ParseDiagnostic>` to its serialized document model.
The non-exported TypeScript parser boundary reads it through
`InternalDocxDocumentModel = DocxDocumentModel & { diagnostics?:
ParseDiagnosticWire[] }`; `packages/docx/src/types.ts` and the A1 public
declaration baseline remain unchanged.

- [ ] **C2b Step 1: Add failing parser-to-layout diagnostic tests**

Build minimal OOXML for a recognized unsupported decoration, invalid geometry,
unknown enum value, and a supported control case. Assert stable codes and source
paths; assert layout maps recoverable cases to warnings, fatal geometry to an
error, and the supported case to no diagnostic.

- [ ] **C2b Step 2: Run tests to verify Red**

Run the Rust parser and TypeScript diagnostic tests. Expected: diagnostic fields
and private wire mapping do not exist.

- [ ] **C2b Step 3: Preserve diagnostics through the private wire boundary**

Record stable parser codes without document text, map them at the parser/layout
boundary, and include them in `DocumentLayout.diagnostics`. Rebuild WASM, run
the parser and layout suites, typecheck, public API comparison, independent
review, and merge.

### Task C2c: Add a synthetic conformance corpus

**Files:**

- Create: `packages/docx/src/conformance/generate.ts`
- Create: `packages/docx/src/conformance/cases.ts`
- Create: `packages/docx/src/conformance/layout.test.ts`
- Create: `packages/docx/tests/visual/conformance.spec.ts`
- Modify: `packages/docx/playwright.config.ts`
- Modify: `.github/workflows/ci.yml`

- [ ] **C2c Step 1: Add failing generated-corpus geometry tests**

Generate redistributable pairwise cases spanning story, container, paragraph,
table, nested table, inline/floating object, direction, spacing, style source,
font source, and anchor reference. Assert page count, line ranges, non-overlap,
bottom-margin clearance, exact deterministic service fingerprints, and
main/worker parity in Vitest. In Chromium, Firefox, and WebKit assert the same
semantic invariants and exact main/worker parity within that browser; compare
cross-browser coordinates with explicit per-primitive tolerances because native
Canvas shaping may differ. Exact cross-browser fingerprints are required only if
the project later supplies one deterministic shaping engine and bundled font
corpus.

Use these explicit comparison rules: authored page boxes, fixed table grids,
fixed border endpoints, and drawing extents are normalized to six decimal places
and compare exactly because they do not depend on shaping. Text coordinates,
line partitions, and page partitions are not compared numerically across browser
engines when native shaping is used. Every browser must instead pass finite
geometry, flow ownership, bottom-margin, structured-clone, text/source-range
coverage, and exact same-browser main/worker fingerprint invariants. No empirical
cross-browser text tolerance is permitted.

- [ ] **C2c Step 2: Run tests to verify Red**

Run:

```bash
cargo test -p docx-parser --test diagnostics -- --nocapture
pnpm vitest run packages/docx/src/layout/diagnostics.test.ts packages/docx/src/conformance/layout.test.ts
pnpm playwright test --config packages/docx/playwright.config.ts conformance.spec.ts --project=chrome
```

Expected: generated fixture modules do not exist.

- [ ] **C2c Step 3: Generate fixtures deterministically**

Implement a deterministic ZIP/XML generator using repository dependencies, with fixed timestamps and IDs,
so generated bytes and deterministic-service expected fingerprints are stable.
Broaden the DOCX Playwright config to include the committed visual/conformance
test directory and add explicit Chrome, Firefox, and WebKit projects. CI runs
node geometry on every change and all three browser projects on the existing
browser-test cadence.

- [ ] **C2c Step 4: Rebuild and verify Green**

Run:

```bash
pnpm build:wasm
cargo test -p docx-parser
pnpm vitest run packages/docx/src/layout/diagnostics.test.ts packages/docx/src/conformance/layout.test.ts
pnpm playwright test --config packages/docx/playwright.config.ts conformance.spec.ts
pnpm typecheck
```

Expected: all checks pass and two consecutive corpus generations have identical
hashes and deterministic-service fingerprints; browser tolerance and parity
assertions pass in all configured projects.

- [ ] **C2c Step 5: Commit, independently review, fix, and merge**

Commit subject: `test(docx): add synthetic conformance corpus`.
Use the roadmap review gate.

### Task C3: Reduce renderer to an adapter and prove architectural completion

**Files:**

- Create: `packages/docx/src/paint/canvas-page.test.ts`
- Modify: `packages/docx/src/paint/canvas-page.ts`
- Modify: `packages/docx/src/renderer.ts`
- Modify: `packages/docx/src/document.ts`
- Delete: `packages/docx/src/layout-context.ts`
- Delete: `packages/docx/src/paragraph-measure.ts`
- Delete: `packages/docx/src/layout-fragments.ts`
- Delete: `packages/docx/src/table-fragments.ts`
- Create: `rules/no-docx-runtime-layout-stamps.yml`
- Create: `rules/no-docx-migration-flags.yml`
- Create: `rule-tests/no-docx-runtime-layout-stamps-test.yml`
- Create: `rule-tests/no-docx-migration-flags-test.yml`
- Create: `packages/docx/src/layout/architecture.test.ts`
- Modify: `scripts/check-docx-public-api.mjs`
- Modify: `scripts/check-docx-layout-boundaries.mjs`
- Modify: `scripts/check-docx-layout-boundaries.test.mjs`
- Delete: `scripts/docx-layout-boundary-baseline.json`
- Create: `.agents/skills/docx-architecture-audit/SKILL.md`
- Modify: `sgconfig.yml`
- Modify: `.github/workflows/ci.yml`
- Modify: `docs/docx-layout-engine-redesign.md`
- Delete: `docs/docx-layout-engine-implementation-roadmap.md`
- Delete: `docs/docx-layout-engine-series-a-plan.md`
- Delete: `docs/docx-layout-engine-series-b-plan.md`
- Delete: `docs/docx-layout-engine-series-c-plan.md`

**Interfaces:**

- Consumes: final `layoutDocument`, `paintLayoutPage`, and layout-derived metadata.
- Produces: existing `paginateDocument` and `renderDocumentToCanvas` compatibility entry points as thin adapters with unchanged public callers.

**Specification evidence:** This PR introduces no OOXML behavior. It proves that
all normative and compatibility decisions are owned by the final layout modules,
that paint dependencies are measurement-free transitively, and that the public
declaration surface is byte-equivalent to the A1 baseline.

```ts
export function paginateDocument(document: Readonly<DocxDocumentModel>): DocumentLayout;
export async function renderDocumentToCanvas(document: Readonly<DocxDocumentModel>, target: HTMLCanvasElement | OffscreenCanvas, pageIndex: number, options: RenderPageOptions): Promise<void>;
```

- [ ] **Step 1: Write failing architecture tests and static rules**

Assert deep-frozen parser and layout inputs survive layout, repeated paint, failed
image paint, and search projection unchanged. Assert two layout calls with the
same services/options have identical fingerprints. Add tested static rules
rejecting runtime stamp properties, `*ReuseEnabled`, `*PaintEnabled`,
`RequiresLegacy`, legacy layout branches, dry-run layout, and layout/measurement
declarations. Extend the A1 import-graph checker so every paint entry's transitive
dependencies are free of measurement, shaping, style merge, pagination, and
parser-object access, and so every pagination/layout entry is on an explicit
allowlist outside `renderer.ts`.

Build the published root `dist/types/docx.d.ts`, normalize source-map paths/comments,
and compare it with `packages/docx/api/public-api-baseline.d.ts`. The comparison
must fail on any added, removed, or changed exported declaration; it replaces a
manual four-file diff.

- [ ] **Step 2: Run tests and static scan to verify Red**

Run:

```bash
pnpm vitest run packages/docx/src/layout/architecture.test.ts packages/docx/src/paint/canvas-page.test.ts
pnpm lint
pnpm lint:test
node scripts/check-docx-layout-boundaries.mjs --final
pnpm build
node scripts/check-docx-public-api.mjs
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
pnpm lint:test
node scripts/check-docx-layout-boundaries.mjs --final
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
pnpm playwright test --config packages/docx/playwright.config.ts conformance.spec.ts
cargo test -p docx-parser
pnpm build
node scripts/check-docx-public-api.mjs
node scripts/check-docx-layout-boundaries.mjs --final
git diff --check
```

Expected: all verification passes; generated declarations exactly match the A1
baseline; every paint dependency is layout/measurement-free; all layout and
pagination declarations are allowlisted in focused modules; no private artifact
is tracked.

- [ ] **Step 6: Obtain an independent final architecture audit**

Create and use the repository-local DOCX architecture-audit skill. It owns the
semantic checks that cannot be proved reliably by identifier matching: spec-first
ownership, one production algorithm per feature class, renderer adapter thinness,
SRP, duplication, and cross-package consistency. Require evidence by command or
exact code reference. Keep only deterministic, inexpensive checks in CI: public
API compatibility, import ownership, migration-flag/stamp absence, plain-data
worker contracts, and retained-reference integrity.

Expected: no paint measurement, parser stamps, or migration flags/fallback; all
stories are laid out through the immutable pipeline; main/worker behavior and
public APIs remain compatible. Run this skill for C3, major DOCX architecture
changes, and periodic audits rather than pretending semantic review is fully
enforceable on every CI run.

- [ ] **Step 7: Clean up, fix findings, reverify the final tree, commit, and merge PR C3**

Commit subject: `refactor(docx): complete immutable layout pipeline`.
Before the final verification, delete the migration roadmap and Series A/B/C execution plans;
their durable architectural decisions remain in `docx-layout-engine-redesign.md`,
while Issue #1037 retains the implementation history. Remove the transitional
boundary baseline, which puts `check-docx-layout-boundaries.mjs` into permanent
final mode; retain that checker and its test for transitive import ownership.
Retain only deterministic static/API/contract gates in CI, and retain the
semantic architecture audit as a repository skill. Repeat Steps 4–6 against the
exact post-cleanup tree after every material fix, then merge with
`gh pr merge <number> --merge`. Close Issue #1037 only after GitHub shows PR C3
merged. Do not create a release or tag.
