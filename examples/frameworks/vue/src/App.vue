<script setup lang="ts">
import { computed, shallowRef, useTemplateRef } from 'vue';
import { useOfficeViewer, type OfficeSource, type OfficeViewerHandle } from './useOfficeViewer';

type OfficeFormat = 'docx' | 'xlsx' | 'pptx';
const SAMPLE = 'https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/packages/docx/public/demo/sample-1.docx';
const document = shallowRef<{ format: OfficeFormat; source: OfficeSource; name: string }>({
  format: 'docx',
  source: SAMPLE,
  name: 'sample-1.docx',
});
const target = useTemplateRef<HTMLDivElement>('target');
const createViewer = computed(() => async (container: HTMLElement): Promise<OfficeViewerHandle> => {
  if (document.value.format === 'xlsx') {
    const { XlsxViewer } = await import('@silurus/ooxml/xlsx');
    return new XlsxViewer(container, { showZoomSlider: true });
  }
  if (document.value.format === 'pptx') {
    const { PptxScrollViewer } = await import('@silurus/ooxml/pptx');
    return new PptxScrollViewer(container, { background: '#53606d' });
  }
  const { DocxScrollViewer } = await import('@silurus/ooxml/docx');
  return new DocxScrollViewer(container, { enableTextSelection: true, background: '#53606d' });
});
const viewer = useOfficeViewer({
  target,
  source: computed(() => document.value.source),
  createViewer: (container) => createViewer.value(container),
});
const chooseFile = async (event: Event) => {
  const file = (event.currentTarget as HTMLInputElement).files?.[0];
  const format = file?.name.split('.').pop()?.toLowerCase();
  if (!file || (format !== 'docx' && format !== 'xlsx' && format !== 'pptx')) return;
  document.value = { format, source: await file.arrayBuffer(), name: file.name };
};
</script>

<template>
  <main class="app">
    <div class="toolbar">
      <button aria-label="Zoom out" @click="viewer.zoomOut">−</button>
      <button @click="viewer.fitWidth">Fit width</button>
      <button aria-label="Zoom in" @click="viewer.zoomIn">+</button>
      <button @click="viewer.reload">Reload</button>
      <input class="file-input" type="file" accept=".docx,.xlsx,.pptx" @change="chooseFile" />
      <span :class="['status', { error: viewer.error.value }]">
        {{ viewer.error.value?.message ?? `${document.name} · ${viewer.status.value}` }}
      </span>
    </div>
    <div ref="target" class="stage"></div>
  </main>
</template>
