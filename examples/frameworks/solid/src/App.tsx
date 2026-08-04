import { createSignal } from 'solid-js';
import type { OfficeViewerOptionsByFormat } from '@ooxml-framework-examples/shared';
import { createOfficeViewer } from './createOfficeViewer';

const SOURCE = 'https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/packages/docx/public/demo/sample-1.docx';
const viewerOptions: OfficeViewerOptionsByFormat['docx'] = {
  width: 820,
  enableTextSelection: true,
  useGoogleFonts: true,
};

export function App() {
  const [target, setTarget] = createSignal<HTMLCanvasElement | null>(null);
  const viewer = createOfficeViewer({
    target,
    format: 'docx',
    source: SOURCE,
    viewerOptions,
  });

  return (
    <main class="app">
      <div class="toolbar">
        <button onClick={viewer.zoomOut} aria-label="Zoom out">−</button>
        <button onClick={viewer.fitWidth}>Fit width</button>
        <button onClick={viewer.zoomIn} aria-label="Zoom in">+</button>
        <button onClick={viewer.reload}>Reload</button>
        <span classList={{ status: true, error: Boolean(viewer.error()) }}>
          {viewer.error()?.message ?? viewer.status()}
        </span>
      </div>
      <div class="stage"><canvas ref={setTarget} /></div>
    </main>
  );
}
