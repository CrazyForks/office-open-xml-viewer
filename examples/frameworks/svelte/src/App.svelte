<script lang="ts">
  import { writable } from 'svelte/store';
  import { createOfficeViewer, type OfficeSource, type OfficeViewerHandle } from './createOfficeViewer';

  type OfficeFormat = 'docx' | 'xlsx' | 'pptx';
  type SelectedDocument = { format: OfficeFormat; source: OfficeSource; name: string };
  const SAMPLE = 'https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/packages/docx/public/demo/sample-1.docx';
  const selected = writable<SelectedDocument>({ format: 'docx', source: SAMPLE, name: 'sample-1.docx' });
  const viewer = createOfficeViewer<OfficeViewerHandle>();
  const { status, error } = viewer;
  const config = $derived({
    source: $selected.source,
    createViewer: async (container: HTMLElement): Promise<OfficeViewerHandle> => {
      if ($selected.format === 'xlsx') {
        const { XlsxViewer } = await import('@silurus/ooxml/xlsx');
        return new XlsxViewer(container, { showZoomSlider: true });
      }
      if ($selected.format === 'pptx') {
        const { PptxScrollViewer } = await import('@silurus/ooxml/pptx');
        return new PptxScrollViewer(container, { background: '#53606d' });
      }
      const { DocxScrollViewer } = await import('@silurus/ooxml/docx');
      return new DocxScrollViewer(container, { enableTextSelection: true, background: '#53606d' });
    },
  });
  const chooseFile = async (event: Event) => {
    const file = (event.currentTarget as HTMLInputElement).files?.[0];
    const format = file?.name.split('.').pop()?.toLowerCase();
    if (!file || (format !== 'docx' && format !== 'xlsx' && format !== 'pptx')) return;
    selected.set({ format, source: await file.arrayBuffer(), name: file.name });
  };
</script>

<main class="app">
  <div class="toolbar">
    <button aria-label="Zoom out" onclick={viewer.zoomOut}>−</button>
    <button onclick={viewer.fitWidth}>Fit width</button>
    <button aria-label="Zoom in" onclick={viewer.zoomIn}>+</button>
    <button onclick={viewer.reload}>Reload</button>
    <input class="file-input" type="file" accept=".docx,.xlsx,.pptx" onchange={chooseFile} />
    <span class:error={$error} class="status">
      {$error?.message ?? `${$selected.name} · ${$status}`}
    </span>
  </div>
  <div class="stage" use:viewer.action={config}></div>
</main>
