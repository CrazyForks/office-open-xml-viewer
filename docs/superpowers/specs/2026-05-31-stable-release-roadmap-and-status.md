# Stable Release Roadmap and Status

- Date: 2026-05-31
- Open PR: #284 (feature/vrt-baseline-infrastructure)
- Purpose: single handoff doc. Where we are, what was found, what is next.

## Summary

A full 5-domain audit was completed. Feature coverage is already mature; the
real gaps for a stable (1.0) release are engineering foundations, CLAUDE.md
principle violations, and large-file maintainability. Sub-project 1 (VRT
regression infrastructure) is implemented and in PR #284. Sub-projects D, B,
C+E are not started.

## Agreed roadmap and order

User chose: harden VRT first, then implement. Order:

1. A-track / VRT regression infra  (DONE, PR #284)
2. D  large-file splits            (verified behavior-invariant via VRT regression mode)
3. B  remove heuristic / sample-fitting code
4. C+E  missing features and performance

Each sub-project is its own spec, plan, implementation, PR cycle.

## Sub-project 1 status (PR #284)

Test-infra only. No renderer/parser/reference changes.

Done:
- skip-on-missing-ref: units with no comparison target report skip, not fail.
- regression mode (VRT_MODE=regression): compare vs baseline/ at strict 0.5 pct.
- snapshot capture (VRT_SNAPSHOT=1): write current output into baseline/.
- baseline/ added to .gitignore in pptx (root), docx, xlsx.

Verified on this machine:
- docx 19/19 pass, pptx 69/69 pass (unchanged = behavior preserved).
- xlsx went from 61 fail (ENOENT crash) to 4 pass / 2 fail / 59 skip.

Deferred (honest): shared-runner DRY extraction. The skip/regression logic is
duplicated across the 3 visual.spec.ts. Long file writes kept failing this
session, so it was done as 3 small inline edits instead. Make it a follow-up PR.

## Maintainer TODO (user-only: reference generation needs Excel eyes)

1. xlsx demo/sample-1 sheet 1,2 are stale (0.37/0.39 row-height fixes). Confirm
   vs Excel, then: UPDATE_REFS=1 pnpm --filter @silurus/ooxml-xlsx vrt
2. xlsx private/sample-1..27: no references yet. Generate after visual confirm
   (27 samples, biggest task). This makes the regression baseline clean.
3. pptx private/sample-7: file exists but uncovered. Needs slide count + refs.

## Full audit findings (the backlog)

### A. Stable-release blockers (engineering foundation)
- No CI validates PR/push. All 4 workflows trigger on v* tag / release.
  publish.yml publishes to npm with no build/typecheck/test gate.
- CLI and Action are non-functional as distributed: ooxml-md and
  ooxml-thumbnail import raw .ts (fail under plain node); the markdown GitHub
  Action has a .ts-launch failure, a Windows path bug, and a missing git push.
- typecheck gap: pnpm typecheck excludes markdown/node/vscode-extension; dts
  build uses skipDiagnostics:true.
- Version drift: markdown/node at 0.36.0, everything else 0.40.0.
- No lint / ast-grep, although CLAUDE.md requires static checks.
- README cites a nonexistent @silurus/ooxml-diff package.

### B. CLAUDE.md principle violations (fidelity over heuristics)
- core chart niceAxisMax 0.9-bump, sample-2 slide-16 overlap fit,
  packages/core/src/chart/renderer.ts:43, affects all 4 chart families.
- core chart waterfall padT = h*0.12, back-derived from sample-2 slide-8
  callout tip, packages/core/src/chart/renderer.ts:2062.
- xlsx MDW_TABLE meiryo magic constant validated against sample-10,
  packages/xlsx/src/renderer.ts:56.
- pptx normAutofit runtime binary search, packages/pptx/src/renderer.ts:1206;
  parser drops fontScale/lnSpcReduction, packages/pptx/parser/src/lib.rs:2711.
- docx lastRenderedPageBreak ruby-only gate, packages/docx/parser/src/parser.rs:447
  (CLAUDE.md explicitly forbids this; self-acknowledged TODO in code).
- minor: xlsx indent 0.5 factor where spec wants 3-char MDW, renderer.ts:2561.

### C. Missing features (parser usually already has the data)
- pptx reflection / innerShdw / softEdge rendering (parsed, renderer ignores).
- docx evenAndOddHeaders: settings.xml is never read; even-page headers lost.
- xlsx number-format color [Red] and conditional sections (e.g. bracket gte 90)
  are dropped; negative-red and fraction formats missing.
- core chart logarithmic axis and secondary axis.
- docx keepLines / widowControl are dead data; footnote / comment body render.
- xlsx row/col outline grouping; diagonal cell borders.

### D. Refactor (maintainability)
- Split giant single files (docx parser is the model, already split):
  xlsx/parser/lib.rs 5804, pptx/parser/lib.rs 5784, xlsx/renderer.ts 4719,
  core/chart/renderer.ts 2318, core/shape/preset.ts 2141.
- xlsx/renderer.ts embeds a ~600-line formula engine for CF at :1602; extract.
- chart renderer duplicates title/legend layout math across 7 functions.
- Rust and TS types are hand-kept in 3 places with no sync; adopt ts-rs or a
  CI drift check.

### E. Performance (large files)
- xlsx rebuilds a full-sheet cell Map and recompiles conditional formatting on
  every scroll frame, renderer.ts:3364. Cache per sheet. Biggest win.
- xlsx cumulative row/col offsets are O(n); getCellAt scans up to ~1M rows per
  click, viewer.ts:247. Use prefix-sum plus binary search.
- xlsx re-extracts the ZIP and re-parses sharedStrings on every sheet switch,
  and clones shared strings into every cell.
- pptx inlines base64 images in JSON and re-decodes them every render,
  renderer.ts:1521. Add an ImageBitmap cache.
- docx recomputes paragraph layout 3-4x per document; no measureText memo.

## Next-session checklist

1. Review and merge PR #284.
2. Generate the missing xlsx references (TODO above) so the regression baseline
   is clean before any refactor.
3. Optional follow-up PR: shared-runner DRY extraction.
4. Start D: split xlsx/parser/lib.rs and xlsx/renderer.ts. Capture baseline with
   VRT_SNAPSHOT=1, refactor, then verify with VRT_MODE=regression (expect ~0 pct).

## Tooling note for the next session (important)

This session repeatedly emitted malformed tool calls whenever prose and a tool
call were combined in one reply: a stray prefix corrupted the tool block, which
then showed as raw XML to the user and never executed, burning usage. The
reliable workaround: emit tool calls with NO surrounding prose, and pass any
content containing angle brackets through a body-file rather than inline.
