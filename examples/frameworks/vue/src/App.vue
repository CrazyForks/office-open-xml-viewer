<script setup lang="ts">
import { useTemplateRef } from 'vue';
import type { OfficeViewerOptionsByFormat } from '@ooxml-framework-examples/shared';
import { useOfficeViewer } from './useOfficeViewer';

const SOURCE = 'https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/packages/docx/public/demo/sample-1.docx';
const target = useTemplateRef<HTMLCanvasElement>('target');
const viewerOptions: OfficeViewerOptionsByFormat['docx'] = {
  width: 820,
  enableTextSelection: true,
  useGoogleFonts: true,
};
const viewer = useOfficeViewer({ target, format: 'docx', source: SOURCE, viewerOptions });
</script>

<template>
  <main class="app">
    <div class="toolbar">
      <button aria-label="Zoom out" @click="viewer.zoomOut">−</button>
      <button @click="viewer.fitWidth">Fit width</button>
      <button aria-label="Zoom in" @click="viewer.zoomIn">+</button>
      <button @click="viewer.reload">Reload</button>
      <span :class="['status', { error: viewer.error.value }]">
        {{ viewer.error.value?.message ?? viewer.status.value }}
      </span>
    </div>
    <div class="stage"><canvas ref="target" /></div>
  </main>
</template>
