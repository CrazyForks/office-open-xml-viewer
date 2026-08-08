# Migrating XLSX selection APIs for 0.77

Version 0.77 removes the endpoint-based XLSX selection compatibility API. Move
to the canonical selection state before upgrading:

| Deprecated compatibility API | Replacement |
| --- | --- |
| `viewer.select(ref)` | `viewer.setSelection(ref)` |
| `viewer.selection` | `viewer.selectionState` |
| `onSelectionChange(selection)` | `onSelectionStateChange(selection)` |
| `CellRange`, `SelectionMode` | `XlsxSelectionState`, `XlsxSelectionArea` |

The old API modeled a selection as `{ anchor, active, mode }`. That shape could
not represent Excel correctly: the selected area, its ActiveCell, and the cell
from which Shift/drag extends are independent. SpreadsheetML makes the same
distinction with `sqref`, `activeCell`, and `activeCellId` (ECMA-376
§18.3.1.78).

## Simple contiguous selections

For a cell or rectangular range, only the method name changes:

```ts
viewer.setSelection('B2:D5');
```

An A1 string describes geometry, not selection direction. `B2:D5` and `D5:B2`
therefore produce the same normalized state, with the upper-left cell used as
the default ActiveCell and extension anchor.

Use Excel's unbounded syntax for whole rows or columns:

```ts
viewer.setSelection('2:4'); // rows 2 through 4
viewer.setSelection('B:D'); // columns B through D
```

`A2:XFD4` remains an explicit bounded cell rectangle. It is not silently
converted into a whole-row selection.

## ActiveCell, extension direction, and multiple areas

Pass structured state when the defaults are insufficient:

```ts
viewer.setSelection({
  areas: [
    { kind: 'cells', top: 2, left: 2, bottom: 5, right: 4 },
    { kind: 'rows', firstRow: 8, lastRow: 9 },
  ],
  activeAreaIndex: 0,
  activeCell: { row: 3, col: 3 },
  extensionAnchor: { row: 2, col: 2 },
});
```

`activeAreaIndex` is the zero-based counterpart of SpreadsheetML's
`activeCellId`. `activeCell` and `extensionAnchor` must be inside that area.
The viewer normalizes reversed bounds, rejects coordinates outside the XLSX
grid, and limits one state to 1,024 areas.

## Selection events

```ts
const viewer = new XlsxViewer(container, {
  onSelectionStateChange(selection) {
    if (!selection) return;
    console.log(selection.activeCell, selection.areas);
  },
});
```

The callback fires only for semantic changes. Sheet changes clear the selection
and report `null`. The deprecated callback receives only a lossy projection of
the active area and cannot expose independent ActiveCell or multiple areas.

## Clipboard behavior

For read-only AI/MCP integrations, use `getSelectionContext()` instead of
round-tripping through the system clipboard:

```ts
const context = viewer.getSelectionContext({ maxCells: 1_000 });
sendToAssistant(context);
```

The context contains canonical selection geometry, sheet identity, formulas,
scalar values, and Viewer-formatted display text. It returns populated cells
only, is detached from workbook internals, and reports truncation. `maxCells`
is hard-capped at 10,000 so an untrusted or accidental full-sheet selection
cannot create an unbounded prompt or allocation.

For an actual clipboard operation, `copySelection()` returns a discriminated
result instead of hiding failures:

```ts
const result = await viewer.copySelection();
if (result.status !== 'copied') {
  console.warn(`Selection was not copied: ${result.status}`);
}
```

Resource checks depend only on the selected geometry and generated text, not on
whether selection came from a pointer or the API. Whole-row, whole-column, and
whole-sheet selections are narrowed to used cells for serialization. Multiple
areas currently return `unsupported-multiple-areas` because flattening disjoint
areas into one TSV has no lossless representation.

## Removal schedule

`select()`, `selection`, `onSelectionChange`, `CellRange`, and `SelectionMode`
remain as migration-only compatibility surfaces on the development line and
are scheduled for removal in 0.77.0. Do not add new usage of them.
