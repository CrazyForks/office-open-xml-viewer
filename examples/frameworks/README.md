# Framework examples

Portable TypeScript integrations for `@silurus/ooxml`. This directory is an
independent pnpm workspace, so installing the main repository does not install
React, Vue, Svelte, or Solid.

Each example combines two layers:

1. `shared/src/office-viewer.ts` dynamically imports the selected DOCX, XLSX,
   or PPTX viewer and guarantees cleanup when loading fails.
2. The framework-specific lifecycle module owns asynchronous mounting,
   replacement, and teardown and exposes the common zoom controls.

Run an example from this directory:

```sh
pnpm install
pnpm dev:react
```

Replace `react` with `vue`, `svelte`, or `solid`. To move an integration into an
application, copy the shared module and the framework module, change its local
import path, and install `@silurus/ooxml` in that application.
