# Read-only selection context

Selection context is the Viewer boundary for integrations that need to act on
what a user is currently looking at, especially AI/MCP assistants. It is not an
editing model. The Viewer never mutates the source file, returns save handles,
or sends context anywhere; the host application decides whether and where to
forward each snapshot.

## One query, format-specific focus kinds

Every Viewer uses `getSelectionContext()`. Results are detached,
JSON-serializable values with a stable two-part discriminator:

| Format | `kind` | Current focus |
| --- | --- | --- |
| DOCX | `text` | Browser-native text selection, with page/paragraph/run locators |
| XLSX | `range` | Canonical worksheet selection, populated cells, formulas and display text |
| PPTX | `text` | Browser-native text selection, with slide/shape/run locators |
| PPTX | `element` | Topmost rendered element clicked or hit-tested on a slide |

Hosts should switch on both `format` and `kind` and keep a default branch. New
read-only focus kinds can then extend the same transport envelope; they do not
require parallel “AI”, “MCP”, “shape click”, or editor-oriented APIs.

DOCX and PPTX expose `onSelectionContextChange` because browser selection and
PPTX element focus are DOM interactions. XLSX already exposes
`onSelectionStateChange`; that notification is the trigger for querying the
same `getSelectionContext()` API. Callbacks are conveniences, while the getter
is the state authority.

## Resource and privacy boundary

- Native selected text is capped at 65,536 UTF-16 code units and never splits a
  surrogate pair.
- At most 1,024 intersected rendered-run locators are retained.
- XLSX separately bounds populated cells and cumulative text; it never expands a
  whole-row, whole-column, whole-sheet, or sparse selection into an unbounded
  rectangular array.
- Every context reports `truncated` and `truncationReasons`.
- A native range is accepted only when all endpoints belong to a tagged Viewer
  text-selection surface. Browser chrome or adjacent page content is not folded
  into the context.
- Returned objects contain no live parser/renderer objects, DOM events, archive
  paths, binary media, or mutation methods.
- Calling a Viewer context getter after `destroy()` throws, matching the XLSX
  content-access contract and preventing stale document data from being read.

## PPTX element focus

`enableElementSelection` makes canvas clicks establish element context. Text
selection takes precedence while it exists; selecting text clears a prior
element focus rather than allowing it to reappear when the browser selection
collapses. Empty-space clicks clear element focus. Async hit tests are generation
gated so a slow earlier click cannot overwrite a later one.

`PptxPresentation.getElementContextAt(slideIndex, point, options)` exposes the
same compact query to custom slide surfaces. Coordinates and tolerance are in
slide EMU. It works in both main and worker modes, walks reverse paint order,
tests transformed element frames (or line segments with tolerance), accounts
for rotation and flips, and returns only bounded descriptive data. Shape/table
text and chart labels/values are streamed into the text budget; picture/media
contexts retain bounds and MIME type so a host may choose to crop its own canvas
or request multimodal analysis without exposing archive paths. It is frame hit
testing, not pixel-alpha or arbitrary custom-path containment.

The parser retains one provenance entry per rendered element:
`master`, `layout`, or `slide`. This is useful for explaining inherited content
without exposing a writable slide tree. Native slide-tree indexes are omitted:
the composite render list can contain inherited or synthesized elements and is
not a lossless round-trip model. `elementIndex` is explicitly the paint-order
index in the current rendered snapshot.

## Extension policy

The public model describes user focus, not an operation to perform. Future
capabilities should add a `kind` to the existing selection-context family only
when the Viewer gains a genuinely new focus target. Details specific to that
kind belong inside its discriminated snapshot. Commands such as edit, replace,
delete, save, or round-trip are outside this API and outside project scope.

This separation lets integrations add operations such as “explain”, “compare”,
“find related information”, “summarize”, or “evaluate” entirely in application
code. Those operations all consume the same bounded selection context and do not
require the Viewer to maintain one interface per assistant action.
