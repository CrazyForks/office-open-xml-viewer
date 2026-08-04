<script lang="ts">
  import type { OfficeViewerOptionsByFormat } from '@ooxml-framework-examples/shared';
  import { createOfficeViewer } from './createOfficeViewer';

  const SOURCE = 'https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/packages/docx/public/demo/sample-1.docx';
  const viewerOptions: OfficeViewerOptionsByFormat['docx'] = {
    width: 820,
    enableTextSelection: true,
    useGoogleFonts: true,
  };
  const viewer = createOfficeViewer<'docx'>();
  const { status, error } = viewer;
  const config = { format: 'docx' as const, source: SOURCE, options: viewerOptions };
</script>

<main class="app">
  <div class="toolbar">
    <button aria-label="Zoom out" onclick={viewer.zoomOut}>−</button>
    <button onclick={viewer.fitWidth}>Fit width</button>
    <button aria-label="Zoom in" onclick={viewer.zoomIn}>+</button>
    <button onclick={viewer.reload}>Reload</button>
    <span class:error={$error} class="status">
      {$error?.message ?? $status}
    </span>
  </div>
  <div class="stage"><canvas use:viewer.action={config}></canvas></div>
</main>
