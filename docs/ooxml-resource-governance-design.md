# OOXML resource governance and bounded processing

Status: accepted direction for implementation on Draft PR #1120. Values marked
as candidates remain subject to corpus calibration before the PR is made ready.

## Context

OOXML packages are ZIP archives. Compressed input size is not a useful memory
bound: a small package can contain a much larger XML part, and parsing that part
can create an XML arena, a format model, serialized output, and renderer state
at the same time. Issue #1102 demonstrates this amplification with one large
SpreadsheetML worksheet.

The existing 512 MiB `maxZipEntryBytes` guard only caps one inflated ZIP entry.
It does not bound aggregate inflation, parser work, model growth, serialization,
layout state, canvas allocation, or the copies made at worker boundaries.

## Goals

- Keep the existing DOCX, XLSX, and PPTX Viewer constructors, `load(source)`,
  navigation methods, callbacks, and load/render error behavior compatible.
- Reject resource-policy violations deterministically before an uncontrolled
  browser- or worker-level failure whenever the library can measure the risk.
- Use one resource-policy vocabulary and one error/report shape across all three
  formats and across main and worker rendering modes.
- Reduce avoidable peak memory by consuming inflated structured parts through
  bounded, back-pressured stages.
- Keep OOXML interpretation specification-first. Streaming must preserve the
  processed infoset, relationships, inheritance, and ordering required by the
  relevant ECMA-376 parts; it must not introduce a reduced parser dialect.
- Preserve structured diagnostic information across worker boundaries and make
  debug usage visible from the main context.

## Non-goals

- This design does not promise constant total memory for every document.
- It does not infer final inflated size from compressed size.
- It does not guarantee that every engine-level OOM, GPU allocation failure, or
  terminated browser process can be converted to a JavaScript error.
- Worker mode is containment and lifecycle isolation, not a hard browser-process
  memory sandbox.
- Streaming does not mean that DOCX can paginate an arbitrary page without
  processing preceding layout state.
- Existing `ArrayBuffer` input necessarily keeps a complete compressed package
  in memory. Bounded inflation and model processing do not change that fact.

## Delivery milestones

Each milestone has an observable exit gate. Later milestones may start only
after their shared contracts are stable; format implementation work may proceed
independently once M2 is complete.

### M0 — Baseline and contract freeze

- Rebase the Draft PR on current `main` and record the behavior of Issues #1102
  and #1088 and Draft PRs #1119 and #1124.
- Freeze the Viewer compatibility boundary, failure routing, public policy
  vocabulary, accounting semantics, and the format-owned responsibilities.
- Preserve baseline public API snapshots and focused resource-limit tests.
- Close the rejected-load worker ownership gap for all three formats.

Exit: this document is internally consistent, cleanup regression tests pass for
DOCX/XLSX/PPTX, existing public API checks pass, and no implementation question
can silently change the public contract.

### M1 — Shared policy and governor

- Replace format-local option interpretation with one immutable normalizer and
  one per-session `ResourceGovernor` in core/common layers.
- Implement the two public limits, deprecated alias reconciliation, practical
  defaults, hidden hard quotas, discriminated errors, usage snapshots, and
  poison semantics.
- Enforce archive entry count plus declared and actual per-entry/aggregate
  inflation without trusting ZIP declarations.

Exit: adversarial Rust and TypeScript tests cover forged sizes, actual overrun,
repeat reads, conflict validation, worker error transport, and poisoned sessions.

### M2 — Package session and bounded pull control plane

- Introduce the random-access `PackageSession`, bounded entry reader, correlated
  pull/ack/release protocol, credits, leases, cancellation, and idempotent close.
- Ensure logical-operation counters survive multiple pulls and that stale or
  unknown transferred resources are disposed.
- Keep complete buffered package input for the existing Viewer source contract;
  do not claim URL range loading in this milestone.

Exit: protocol tests prove bounded in-flight data, backpressure, timeout/abort
convergence, late-response disposal, and main/worker result and error parity.

### M3 — XLSX bounded worksheet pipeline

- Preserve SpreadsheetML dependency resolution and Part 3 MCE semantics while
  moving worksheet XML through complete-row batches.
- Remove the avoidable full worksheet string plus full serialized model overlap
  from the Viewer path; keep `getWorksheet()` as an explicitly materializing
  compatibility adapter.

Exit: synthetic large worksheets cross multiple pulls, render identically, stop
at deterministic limits, and show a measured reduction in transient peak usage.

### M4 — PPTX slide-granular pipeline

- Keep presentation/theme/master/layout dependencies shared while parsing and
  retaining slides and their relationships as format-owned units.
- Make navigation, media leases, cleanup, and resource failures consistent with
  the common session contract.

Exit: multi-slide tests prove on-demand unit ownership, navigation compatibility,
bounded transient retention, and cleanup on rejection, reload, and destroy.

### M5 — DOCX sequential layout pipeline

- Feed normalized body blocks into the single immutable acquisition ->
  normalization -> layout -> paint pipeline without creating a legacy layout
  path or promising random page access before pagination.
- Align recoverable parse/layout containment with Issue #1088 and retain the
  state required for page count, fields, bookmarks, and stable Viewer readiness.

Exit: pagination and visual behavior remain stable, recoverable failures yield
the defined partial result, resource failures are deterministic, and the DOCX
architecture audit passes.

### M6 — Containment and Ratatui-inspired diagnostics

- Normalize options before worker creation, terminate monolithic timed-out work,
  and converge all failure paths on idempotent disposal.
- Implement `debug: true` checkpoints and one polished final console card using
  a pure shared presentation model with browser CSS, Node ANSI, and plain output.
- Document Worker mode as stronger lifecycle containment, not a memory sandbox
  and not the default.

Exit: no rejected load leaves an owned worker or transfer alive; debug output is
snapshot-tested, data-safe, and equivalent across format and execution mode.

### M7 — Calibration and release-quality verification

- Measure public and synthetic documents across formats and realistic browser
  environments, then adopt or revise the candidate defaults.
- Run Rust, rebuilt-WASM, focused/full TypeScript, typecheck, build, public API,
  visual, and high-water verification appropriate to the touched surfaces.
- Publish the limits as admission policy, not as a promise of exact memory use.

Exit: defaults have recorded evidence and acceptable false-rejection behavior;
all mandatory checks pass or any environment-only limitation is documented with
a reproducible local command.

### M8 — Independent critical review and Draft handoff

- Obtain independent Fable and GPT-5.6 Sol reviews of OOXML specification
  fidelity, responsibility boundaries, duplicate logic, API consistency, error
  semantics, and the claims made about bounded processing.
- Fix every accepted finding and re-run the affected gates.
- Push the reviewed branch and update Draft PR #1120 without merging it.

Exit: review findings and dispositions are recorded, the branch is clean and
reproducible, and the user can validate the Draft locally before any merge
decision.

## Compatibility boundary

The public Viewer surface is the compatibility boundary:

```ts
new DocxViewer(canvas, options).load(source)
new XlsxViewer(container, options).load(source)
new PptxViewer(canvas, options).load(source)
```

The existing `string | ArrayBuffer` sources remain accepted. Existing options
remain accepted. New resource and debug options are additive. A successful
`load()` resolves at the same user-visible readiness point as before, and the
existing `onError` versus rejection behavior remains unchanged.

The following are internal implementation contracts and may change:

- worker request/response envelopes;
- wasm-bindgen parser exports;
- parser-to-renderer model representation;
- JSON versus binary/chunk transport;
- archive-handle lifetime and caching;
- parsing and rendering batch sizes.

Lower-level `DocxDocument`, `XlsxWorkbook`, and `PptxPresentation` APIs should
remain source-compatible where a facade can preserve them without defeating the
resource design. They are not allowed to dictate the worker wire representation.
When a lower-level compatibility method explicitly returns or exposes a complete
model, its adapter may necessarily materialize that complete model; bounded
retention must not be claimed for that path.

New failures follow the existing ownership boundary:

| Failure | Direct factory | Viewer without `onError` | Viewer with `onError` |
| --- | --- | --- | --- |
| invalid resource option | `load()` rejects before worker creation | `load()` rejects | callback receives the error and `load()` resolves |
| detected resource limit | `load()` rejects with typed error | `load()` rejects | callback receives the error and `load()` resolves |
| worker crash during load | `load()` rejects and disposes the partial engine | `load()` rejects | callback receives the error and `load()` resolves |
| recoverable DOCX content/layout failure | partial-result contract defined by Issue #1088 | same partial result | same partial result plus its structured diagnostic |

Viewer construction remains non-throwing for stored load options. Validation is
performed when `load()` starts, before a worker or WASM archive is created.

## Public policy

The long-term public option is a plain object. Callers do not need a policy
factory.

```ts
export type OoxmlResourceLimit = number | null;

export interface OoxmlResourceLimits {
  /** Actual inflated bytes for any one archive entry, including media. */
  maxArchiveEntryBytes?: OoxmlResourceLimit;

  /** Sum of actual inflated bytes across distinct entries in one session. */
  maxTotalInflatedBytes?: OoxmlResourceLimit;
}

export interface LoadOptions {
  resourceLimits?: OoxmlResourceLimits;
  debug?: boolean;

  /** @deprecated Use resourceLimits.maxArchiveEntryBytes. */
  maxZipEntryBytes?: number;
}
```

Semantics:

- `undefined` selects the library's standard default.
- A positive safe integer overrides that default in bytes.
- `null` disables that configurable policy limit only. It does not disable
  non-configurable hard safety quotas.
- Invalid values in `resourceLimits` reject before a worker is created.
- The deprecated `maxZipEntryBytes` adapter preserves its historical input
  behavior and remains an all-entry limit, including media. If callers supply it
  together with `maxArchiveEntryBytes` and the two supplied values disagree,
  option normalization rejects instead of applying hidden precedence.
- Standard defaults are compatibility policy and may be revised deliberately
  between releases. Hard quotas protect implementation invariants and are not a
  supported tuning surface.

`maxTotalInflatedBytes` is stable under lazy loading and internal batch changes:
the session records the greatest actual inflated size observed for each distinct
entry and sums those maxima. Re-reading an entry does not consume this public
budget again. Repeated-inflation work and per-structured-part amplification are
measured separately and protected by internal operation/unit quotas.

Initial calibration candidates are 128 MiB per archive entry and 256 MiB total.
They are not memory guarantees and must not be documented as final until the M7
corpus and browser measurements support them. Adopting defaults below the old
512 MiB per-entry default intentionally narrows the set of documents that load
without overrides: source compatibility is preserved, but behavioral
compatibility for documents above the new defaults is not.

Archive entry count, XML nesting, relationships, model complexity, serialized
bytes, image dimensions, canvas pixels, and timeouts remain internal quotas or
separate existing options unless evidence shows that users can tune them
meaningfully. A resource error reports whether the violated limit was
configurable.

## Accounting model

Resource accounting belongs to a per-document package session, not to parser
call sites or renderer-specific helpers.

The session records distinct concepts instead of forcing them into one counter:

- compressed input bytes observed by the loader;
- central-directory entry count and declared expanded sizes;
- actual inflated bytes per archive entry;
- distinct-entry inflated total used by the public session limit;
- actual inflated bytes per structured part and indivisible parser unit;
- actual bytes delivered during each operation, counting repeated reads again;
- high-water serialized/model bytes where they can be measured explicitly;
- observed WASM linear-memory pages;
- image and canvas dimensions/pixels before allocation.

Declared ZIP sizes are attacker-controlled. A declaration already above a
limit is sufficient for early rejection, but a declaration below a limit is not
proof of safety. Actual output is checked while it is inflated, using
`limit + 1` semantics where necessary to distinguish exact completion from
truncation.

After a proven package-policy violation, that package session is poisoned: no
later operation may continue reading from a partially trusted archive.

A logical operation may span many protocol pulls. Its internal work counters
live in the session ledger until that operation completes or aborts; they must
not reset at each WASM export or each chunk request.

## Internal responsibilities

### `ResourcePolicy`

An immutable normalized policy created before worker construction. It merges
standard defaults, the deprecated compatibility adapter, and caller overrides.
It contains no counters and performs no I/O.

### `ResourceGovernor`

The sole owner of counters, limit checks, high-water marks, the first violation,
and the final usage report. Parser and renderer code report observations to it;
they do not independently interpret public options.

### `PackageSession`

Owns the source bytes or random-access source, ZIP archive/index, governor, and
part handles. It provides bounded entry and structured-part readers, prevents
reads after abort/poison/close, and releases all retained archive resources
idempotently.

The input abstraction is random-access rather than a linear `ReadableStream`,
because ZIP discovery requires the end-of-central-directory record and seeks.
The initial implementation retains the complete compressed source, matching the
existing API. Range-backed URL input is a later independent optimization and
requires response-version pinning plus an owned entry reader; it is not a
prerequisite for bounded inflated-part processing. Agile-encrypted CFB input
remains on the buffered path until the crypto/container layer is redesigned.

### Pull protocol

Consumers request work instead of receiving an unbounded push:

```text
open source -> manifest/session
pull resource chunk with credit -> data + usage checkpoint + leases
consume/acknowledge -> release transient chunk and leases
pull next chunk
complete or abort -> final report + close
```

Credit bounds the amount of work and wire data produced for one pull; it is
backpressure, not a security policy. Resource limits remain independently
enforced. One indivisible format unit may exceed ordinary credit only when it is
below its hard unit limit; otherwise it fails deterministically instead of
deadlocking the stream.

At most the internally allowed number of chunks may be in flight. Abort,
timeout, resource violation, parser trap, and explicit destruction all converge
on the same idempotent session close path. Worker-retained payloads carry leases
that must be released. Transferred objects in stale or unknown responses, such
as `ImageBitmap`, are disposed rather than merely ignored.

The protocol uses structured messages internally; it is not exported as a
Viewer API. Large payloads use transferable buffers where crossing a worker
boundary is unavoidable.

The outer protocol standardizes correlation, sequencing, cancellation, usage,
leases, and close behavior. Format payloads and stream names remain opaque to
it. There is intentionally no common `ModelChunk<Row | Slide | Page>`.

## Format-specific processing

### XLSX

Workbook metadata, relationships, styles, theme, and shared strings are resolved
before dependent sheet rows. `sheetData` is consumed as complete row batches
while preserving ECMA-376 Part 3 Markup Compatibility processing. The renderer
or compatibility facade acknowledges each batch before another is produced.

The current bounded row-arena implementation is useful but incomplete: it still
starts from a fully inflated worksheet string and eventually constructs and
serializes the complete worksheet model. The target pipeline removes those
avoidable simultaneous representations. A selected sheet's retained cell model
may still grow with the number of cells.

Worker Viewer paths can retain bounded row tiles and compact sheet metadata.
The lower-level `getWorksheet()` compatibility adapter may drain those streams
and assemble a complete `Worksheet`, with correspondingly unbounded model size.

### PPTX

Presentation metadata, themes, slide masters, and layouts form the shared
dependency set. Slides and their relationships/media are independent processing
units after those dependencies are available. A slide may be parsed and cached
on demand without changing Viewer navigation semantics.

### DOCX

Styles, numbering, theme, settings, relationships, headers/footers, notes, and
other referenced stories must be available when their body content is
normalized. Body blocks can flow sequentially into the single immutable
acquisition -> normalization -> layout -> paint pipeline, releasing raw XML and
temporary parser arenas behind the layout frontier.

Pagination remains sequential because preceding layout determines later page
positions and total page count. Completed page/layout state, cross-references,
fields requiring total pages, and content needed by public document operations
may remain retained. Recoverable object/block failures and later layout failures
must follow Issue #1088's explicit containment boundaries rather than being
silently defaulted or routed through a legacy layout path.

The current Viewer readiness contract requires pagination to finish so page
count, per-page sizes, and bookmark destinations are stable. Streaming reduces
transient parser/model overlap but does not remove that readiness barrier.

## Abstraction boundary

The following are real shared abstractions:

- correlated pull-session lifecycle;
- resource ledger, reservations, and leases;
- random-access package source and bounded ZIP entry reader;
- timeout, cancel, close, error wire, and late-transfer disposal;
- common debug report/view rendering.

The following remain format-owned:

- semantic chunk types and batch boundaries;
- dependency graphs and readiness barriers;
- retained model/layout stores and eviction policy;
- DOCX pages, PPTX slides, and XLSX row tiles.

Sharing the control plane must not create a second generic OOXML semantic model.

## Failure contract

A detected policy violation throws one `OoxmlResourceLimitError` family across
all formats and modes. Its details are a discriminated resource record so an
archive-level entry-count or declared-total violation does not need a dummy part
name. Shared fields include:

- stable code and stage;
- resource/metric identifier;
- limit and observed value;
- safe OOXML part address only when applicable;
- operation/format identifier;
- whether the limit is configurable;
- the last complete usage snapshot.

Worker serialization reconstructs the real error subclass on the main side.
Messages and debug reports never include document text, passwords, URLs, sheet
names, or local filesystem paths.

A residual `WebAssembly.RuntimeError` or allocation-shaped trap remains
`parser-crashed` unless the governor had already proved a specific limit
violation. The implementation must not promise a reliable `parser-oom`
classification from `unreachable`, because Rust panic, stack overflow, and
allocation failure can converge on the same trap surface.

Canceling or timing out a legacy monolithic synchronous WASM call cannot stop it
cooperatively: the worker cannot process another message until the call returns.
During migration that path terminates the worker. Cooperative cancellation is
only claimed once the operation is divided into bounded pull exports.

Viewer reload intentionally keeps the old engine alive until the new load
succeeds. That compatibility behavior creates a temporary two-session memory
window. A per-session limit does not claim to bound this combined process peak.

## Debug reporting

`debug: true` enables measurement reporting but does not change limits or parser
behavior. Workers send structured usage checkpoints to the main context. The
main context retains checkpoints quietly and emits one final success/failure
card, or the last checkpoint if the worker dies.

The presentation is Ratatui-inspired: bordered blocks, a compact table, gauges,
and semantic status colors. A shared styled-token view is rendered as browser
console CSS, Node ANSI when attached to a TTY, or deterministic plain Unicode.
The pure, color-free representation is snapshot-tested. Console rendering and
resource measurement remain separate responsibilities.

## Verification gates

- Public Viewer type and behavioral compatibility tests for all three formats.
- Declared-size forgery, actual overrun, entry-count, aggregate/operation, repeat
  read, poison, abort, timeout, and cleanup tests.
- Main/worker parity for successful results, typed errors, partial DOCX results,
  and final metrics.
- Specification-focused tests for streamed MCE and format dependency behavior.
- Synthetic large documents whose memory amplification can be varied without
  private fixtures.
- Peak/high-water benchmarks that distinguish input, inflate, parser/model,
  serialization, worker transfer, layout, and rendering stages.
- Rust tests and clippy, rebuilt WASM, focused and full TypeScript tests,
  typecheck, build, public API checks, and the DOCX architecture audit.
- Independent Fable and GPT-5.6 Sol reviews followed by fixes and re-verification.

The Draft PR is not merged until these gates are satisfied and the user has
reviewed the result locally.
