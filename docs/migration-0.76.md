# Migrating from 0.75 to 0.76

Version 0.76 is a breaking release for the Node parser helpers. It removes the
synchronous compatibility parsers in favor of one canonical asynchronous,
owned-session pipeline. Browser `Document`, `Presentation`, `Workbook`, and
Viewer load flows remain source-compatible.

## Node parser replacements

| Removed 0.75 export | 0.76 replacement |
| --- | --- |
| `parseDocx()` | `await materializeDocxDocument()` |
| `parsePptx()` | `await materializePptxPresentation()` |
| `parseXlsx()` | `await materializeXlsxWorkbookIndex()` or `session.workbookIndex` |
| `parseXlsxSheet()` | `await materializeXlsxWorksheet()` |
| `parseXlsxAllSheets()` | `await materializeXlsxWorkbook()` |
| `extractPptxImage()` | `await session.getImage()` on `openPptxPresentation()` |
| `extractPptxMedia()` | `await session.getMedia()` on `openPptxPresentation()` |

Use a materializer when the application needs a complete caller-owned model:

```ts
import { materializePptxPresentation } from '@silurus/ooxml/node';

const presentation = await materializePptxPresentation(bytes);
```

Use an owned session for bounded sequential work:

```ts
import { openPptxPresentation } from '@silurus/ooxml/node';

const presentation = await openPptxPresentation(bytes);
try {
  for await (const slide of presentation.slides()) {
    // Consume one caller-owned slide.
  }
} finally {
  await presentation.close();
}
```

`close()` is idempotent. DOCX `pages()` and PPTX `slides()` are terminal,
one-pass iterators and close their session on completion, break, or error. XLSX
`worksheetRows()` closes only the active worksheet operation, leaving the
workbook session available for another sheet. Immutable counts, dimensions,
names, workbook index, and the last resource-usage snapshot remain readable
after close; new rendering, extraction, or streaming work rejects.

There is intentionally no synchronous wrapper around these asynchronous APIs.

## XLSX canvas-mounted sheet viewer

`XlsxSheetViewer` adds the same unit-viewer mounting shape used by
`DocxViewer` and `PptxViewer`: it renders an active worksheet into a
caller-owned `HTMLCanvasElement`. It includes selection, search, zoom, and
logical viewport APIs, but no sheet tabs, footer controls, zoom chrome, or
native scrollbar. Use the existing container-mounted `XlsxViewer` when those
workbook controls are wanted.

```ts
import { XlsxSheetViewer } from '@silurus/ooxml/xlsx';

const viewer = new XlsxSheetViewer(canvas);
await viewer.load(fileBuffer);
await viewer.goToSheet(1);
await viewer.setViewportOffset({ x: 120, y: 80 });

viewer.destroy();
```

## Archive entry-count limit

All Browser and Node load/session options now accept
`resourceLimits.maxArchiveEntries`. Omission uses the calibrated default,
`null` disables the configurable limit, and the internal hard ceiling remains
enforced.
