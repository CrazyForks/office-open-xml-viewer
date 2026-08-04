import { useCallback, useState, type ChangeEvent } from 'react';
import { useOfficeViewer, type OfficeSource, type OfficeViewerHandle } from './useOfficeViewer';

type OfficeFormat = 'docx' | 'xlsx' | 'pptx';
interface SelectedDocument { format: OfficeFormat; source: OfficeSource; name: string }

const SAMPLE = 'https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/packages/docx/public/demo/sample-1.docx';

export function App() {
  const [selected, setSelected] = useState<SelectedDocument>({
    format: 'docx',
    source: SAMPLE,
    name: 'sample-1.docx',
  });
  const createViewer = useCallback(async (container: HTMLElement): Promise<OfficeViewerHandle> => {
    if (selected.format === 'xlsx') {
      const { XlsxViewer } = await import('@silurus/ooxml/xlsx');
      return new XlsxViewer(container, { showZoomSlider: true });
    }
    if (selected.format === 'pptx') {
      const { PptxScrollViewer } = await import('@silurus/ooxml/pptx');
      return new PptxScrollViewer(container, { background: '#53606d' });
    }
    const { DocxScrollViewer } = await import('@silurus/ooxml/docx');
    return new DocxScrollViewer(container, { enableTextSelection: true, background: '#53606d' });
  }, [selected.format]);
  const viewer = useOfficeViewer({ source: selected.source, createViewer });
  const chooseFile = useCallback(async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    const format = file?.name.split('.').pop()?.toLowerCase();
    if (!file || (format !== 'docx' && format !== 'xlsx' && format !== 'pptx')) return;
    setSelected({ format, source: await file.arrayBuffer(), name: file.name });
  }, []);

  return (
    <main className="app">
      <div className="toolbar">
        <button onClick={viewer.zoomOut} aria-label="Zoom out">−</button>
        <button onClick={viewer.fitWidth}>Fit width</button>
        <button onClick={viewer.zoomIn} aria-label="Zoom in">+</button>
        <button onClick={viewer.reload}>Reload</button>
        <input className="file-input" type="file" accept=".docx,.xlsx,.pptx" onChange={chooseFile} />
        <span className={viewer.error ? 'status error' : 'status'}>
          {viewer.error?.message ?? `${selected.name} · ${viewer.status}`}
        </span>
      </div>
      {viewer.renderOfficeViewer({ className: 'stage' })}
    </main>
  );
}
