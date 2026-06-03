// Framework integration snippets shown in the Code section. PptxViewer is the
// representative viewer (canvas-based, has destroy()); DocxViewer/XlsxViewer
// follow the same shape.

export const vanillaSnippet = `import { PptxViewer } from '@silurus/ooxml/pptx';

const canvas = document.getElementById('deck') as HTMLCanvasElement;
const viewer = new PptxViewer(canvas, { width: 960 });

await viewer.load('/quarterly-review.pptx');
viewer.nextSlide();`;

export const reactSnippet = `import { useEffect, useRef } from 'react';
import { PptxViewer } from '@silurus/ooxml/pptx';

export function Deck({ src }: { src: string }) {
  const ref = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    const canvas = ref.current;
    if (!canvas) return;
    const viewer = new PptxViewer(canvas, { width: 960 });
    void viewer.load(src);
    return () => viewer.destroy();
  }, [src]);

  return <canvas ref={ref} />;
}`;

export const vueSnippet = `<script setup lang="ts">
import { onMounted, onBeforeUnmount, ref } from 'vue';
import { PptxViewer } from '@silurus/ooxml/pptx';

const props = defineProps<{ src: string }>();
const canvas = ref<HTMLCanvasElement>();
let viewer: PptxViewer | undefined;

onMounted(() => {
  viewer = new PptxViewer(canvas.value as HTMLCanvasElement, { width: 960 });
  void viewer.load(props.src);
});
onBeforeUnmount(() => viewer?.destroy());
</script>

<template>
  <canvas ref="canvas" />
</template>`;

export const svelteSnippet = `<script lang="ts">
  import { onMount } from 'svelte';
  import { PptxViewer } from '@silurus/ooxml/pptx';

  export let src: string;
  let canvas: HTMLCanvasElement;

  onMount(() => {
    const viewer = new PptxViewer(canvas, { width: 960 });
    void viewer.load(src);
    return () => viewer.destroy();
  });
</script>

<canvas bind:this={canvas}></canvas>`;

export const headlessSnippet = `import { PptxPresentation } from '@silurus/ooxml/pptx';

// Headless engine — render any slide into a canvas you control. Build your
// own thumbnail grid, scroll view, or master–detail pane around it.
const deck = await PptxPresentation.load('/quarterly-review.pptx');

for (let i = 0; i < deck.slideCount; i++) {
  const canvas = document.createElement('canvas');
  thumbnails.appendChild(canvas);
  await deck.renderSlide(canvas, i, { width: 240 });
}`;
