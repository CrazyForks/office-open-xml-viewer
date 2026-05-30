# VRT Baseline Infrastructure — Design

- **Date**: 2026-05-30
- **Status**: Draft (autonomous night work; awaiting user review)
- **Branch**: `feature/vrt-baseline-infrastructure`
- **Sub-project**: 1st of the stable-release roadmap (A-track → **VRT** → D → B → C+E)

## Context & Goal

Toward a stable (1.0) release we will refactor large files (D), remove
spec-violating heuristics (B), and implement missing features + perf (C+E).
Every one of those changes risks visual regressions.

The user maintains a large local corpus of gitignored Word/Excel/PowerPoint
files (`packages/*/public/{demo,private}/`: pptx 8, docx 6, **xlsx 28** = 42
files) that drove the current rendering quality, with which the user is
satisfied. We need a regression-detection foundation that exercises this full
corpus before any renderer/parser change.

This sub-project hardens the existing Playwright VRT into that foundation. It
changes **only test-infrastructure code**. It does **not** touch reference
images (user-updated only) and does **not** touch any renderer/parser.

## Current baseline health (measured 2026-05-30)

`pnpm build:wasm && pnpm vrt` (WASM rebuilt from 0.40.0 sources, observation only):

| Package | References | Result | Notes |
|---|---|---|---|
| **docx** | complete (demo/sample-1 + private/1–5, 19 imgs) | **19/19 pass** | per-page diff 0.7–12% |
| **pptx** | complete (demo/sample-1 + private/1–6, 69 imgs) | **69/69 pass** | per-slide diff 4–8%; `private/sample-7` exists but is in neither the test list nor references (uncovered) |
| **xlsx** | **demo/sample-1 only (5 imgs)** | **4 pass / 61 fail** | see below |

xlsx failure breakdown (not a rendering bug):
- `demo/sample-1` sheet 1,2 **stale** (diff 24.8% / 20.1%). Known per the 0.39.0
  CHANGELOG: the 0.37.0 pt→px row-height change and 0.39.0 `<row ht>` gate fix
  were never adopted into references. Sheets 3–5 pass.
- `private/sample-1..27` have **no reference images** (ENOENT) — 27 samples, the
  bulk of the failures.

**Implication**: docx/pptx baselines are clean enough; **xlsx is the big gap**
(private refs ungenerated, 2 demo sheets stale). Generating/refreshing those is
`UPDATE_REFS` = **user-only** (requires Excel visual confirmation).

## Key design insight: two independent axes

The 4–12% residual diff on *passing* docx/pptx pages shows that reference
comparison conflates two different questions:

1. **Fidelity** — "does our render match Excel/PowerPoint output?" Compare
   against `references/` (Office exports). Tolerant threshold (20%) because
   Canvas vs Office font rendering differs. References change only when the user
   confirms a render is correct.
2. **Regression** — "did this change alter our render at all?" Compare against a
   **pre-change snapshot of our own output** — same environment / WASM / fonts —
   so the only variable is our code. Strict threshold (≈0%).

Refactors (D) and heuristic removal (B) need **axis 2**. Stale fidelity refs
(xlsx) must not block regression testing. The harness must support both.

## Scope (this PR)

**In:**
- Extract the three near-identical `visual.spec.ts` into one shared harness;
  each package's spec becomes a thin descriptor.
- **Missing-reference handling**: a unit (sheet/slide/page) with no reference is
  reported as **skipped**, not failed — kills the xlsx ENOENT noise and surfaces
  coverage gaps explicitly.
- **Regression mode**: snapshot current output to `baseline/`, then compare
  against it at a strict threshold.
- Add pptx `private/sample-7` to coverage (reports as skipped until refs exist).
- A run summary: pass / fail(diff>threshold) / skip(no ref) / size-mismatch,
  plus a stale-candidate list (passing but diff above a warn line).

**Out:**
- Generating or refreshing any reference image (user-only).
- The D/B/C+E work itself.
- Fixture-based auto-detection of unit counts (follow-up; this PR centralizes the
  explicit per-sample counts but does not eliminate them).
- CI wiring (VRT stays local per repo policy; CI is a separate A-track item).

## Architecture

```
tests/visual-harness/
  runner.ts        # shared: render → capture → compare → report; fidelity + regression; skip-on-missing-ref
  types.ts         # SuiteConfig, SampleSpec, RunMode
packages/<pkg>/tests/visual/
  visual.spec.ts   # thin: imports runner; supplies pkg name + URL builder + sample list
  fixture.html     # unchanged
  references/      # unchanged (user-owned)
  baseline/        # NEW, gitignored — regression snapshots
  screenshots/, diffs/   # unchanged (gitignored artifacts)
```

### Shared runner API

```ts
type RunMode = 'fidelity' | 'regression';
interface SampleSpec { name: string; units: number; width?: number; }
interface SuiteConfig {
  pkg: 'pptx' | 'docx' | 'xlsx';
  unitNoun: 'slide' | 'page' | 'sheet';
  samples: SampleSpec[];
  url: (sample: string, unitIndex: number, spec: SampleSpec) => string;
  failAbovePct: number;   // fidelity threshold (20)
  warnAbovePct: number;   // stale-candidate warn line (e.g. 8)
  regressionPct: number;  // regression threshold (e.g. 0.5)
}
runVisualSuite(config): void   // registers Playwright tests
```

Per-unit behavior:
1. Load `config.url(...)`, wait for fixture `ready`/`error`.
2. Capture canvas → `screenshots/<name>/<noun>-<n>.png`.
3. `UPDATE_REFS=1` → write into `references/` and return. *(user-only)*
4. `VRT_SNAPSHOT=1` → write into `baseline/` and return.
5. Target = regression mode ? `baseline/...` : `references/...`.
6. Target missing → **skip** (annotate + record in summary), not fail.
7. pixelmatch → write diff → record pct → fail only if over the mode threshold.

Env selection: `VRT_MODE=regression` (default `fidelity`); `VRT_SNAPSHOT=1`
captures baseline.

### Regression workflow (used by D / B / C+E)

```bash
# before change (clean tree):
VRT_SNAPSHOT=1 pnpm vrt          # capture current output → baseline/
# ... make the change ...
VRT_MODE=regression pnpm vrt     # compare to baseline; expect ~0%
```

Baseline captured in the same environment, so font/AA noise cancels; only real
behavior changes surface. Immune to which WASM build produced the references.

## Verification (of this PR itself)

This PR must not change VRT outcomes:
- **docx 19 pass, pptx 69 pass** must hold via the shared harness.
- **xlsx**: 27 ENOENT failures become **skips**; demo sheet 1,2 stay
  non-passing (stale, surfaced clearly); sheets 3–5 pass. Net: xlsx goes from a
  hard "61 fail" to "≈2 stale / ≈59 skip / 4 pass".
- `VRT_SNAPSHOT=1 pnpm vrt` then `VRT_MODE=regression pnpm vrt` on an unchanged
  tree → all compared units ≈0%.
- **No tracked reference image changes** (`git status` clean for
  `references/demo/**`; private refs are gitignored anyway).

## Handoff to user (reference work — needs your eyes)

1. **xlsx demo/sample-1 sheet 1,2**: stale since 0.37/0.39 row-height fixes.
   Confirm current render vs Excel, then
   `UPDATE_REFS=1 pnpm --filter @silurus/ooxml-xlsx vrt`.
2. **xlsx private/sample-1..27**: no references. Confirm each vs Excel, then
   UPDATE_REFS. (27 samples — the biggest task.)
3. **pptx private/sample-7**: now in coverage but no refs; generate after visual
   confirmation.
4. Review whether this harness direction (case B) is right before we build on it.

> Note: WASM was rebuilt from current 0.40.0 sources during measurement. If your
> local WASM was older, day-to-day diffs may have been smaller; the regression
> workflow above is immune to this (same-environment snapshot).

## Out of scope / follow-ups

- Fixture-based auto-detection of unit counts.
- CI integration (separate A sub-project: typecheck + smoke on PR; VRT stays local).
- D / B / C+E sub-projects.
