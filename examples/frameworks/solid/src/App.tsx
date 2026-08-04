import { createSignal } from 'solid-js';
import { createOfficeViewer, type OfficeSource, type OfficeViewerHandle } from './createOfficeViewer';

type OfficeFormat = 'docx' | 'xlsx' | 'pptx';
interface SelectedDocument { format: OfficeFormat; source: OfficeSource; name: string }

const SAMPLE = 'https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/packages/docx/public/demo/sample-1.docx';

export function App() {
  const [target, setTarget] = createSignal<HTMLElement | null>(null);
  const [selected, setSelected] = createSignal<SelectedDocument>({
    format: 'docx',
    source: SAMPLE,
    name: 'sample-1.docx',
  });
  const createViewer = () => async (container: HTMLElement): Promise<OfficeViewerHandle> => {
    if (selected().format === 'xlsx') {
      const { XlsxViewer } = await import('@silurus/ooxml/xlsx');
      return new XlsxViewer(container, { showZoomSlider: true });
    }
    if (selected().format === 'pptx') {
      const { PptxScrollViewer } = await import('@silurus/ooxml/pptx');
      return new PptxScrollViewer(container, { background: '#53606d' });
    }
    const { DocxScrollViewer } = await import('@silurus/ooxml/docx');
    return new DocxScrollViewer(container, { enableTextSelection: true, background: '#53606d' });
  };
  const viewer = createOfficeViewer({
    target,
    source: () => selected().source,
    createViewer,
  });
  const chooseFile = async (event: Event) => {
    const file = (event.currentTarget as HTMLInputElement).files?.[0];
    const format = file?.name.split('.').pop()?.toLowerCase();
    if (!file || (format !== 'docx' && format !== 'xlsx' && format !== 'pptx')) return;
    setSelected({ format, source: await file.arrayBuffer(), name: file.name });
  };

  return (
    <main class="app">
      <div class="toolbar">
        <button onClick={viewer.zoomOut} aria-label="Zoom out">−</button>
        <button onClick={viewer.fitWidth}>Fit width</button>
        <button onClick={viewer.zoomIn} aria-label="Zoom in">+</button>
        <button onClick={viewer.reload}>Reload</button>
        <input class="file-input" type="file" accept=".docx,.xlsx,.pptx" onChange={chooseFile} />
        <span classList={{ status: true, error: Boolean(viewer.error()) }}>
          {viewer.error()?.message ?? `${selected().name} · ${viewer.status()}`}
        </span>
      </div>
      <div class="stage" ref={setTarget} />
    </main>
  );
}
