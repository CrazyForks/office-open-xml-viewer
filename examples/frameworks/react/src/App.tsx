import { useMemo } from 'react';
import type { OfficeViewerOptionsByFormat } from '@ooxml-framework-examples/shared';
import { useOfficeViewer } from './useOfficeViewer';

const SOURCE = 'https://raw.githubusercontent.com/yukiyokotani/office-open-xml-viewer/main/packages/docx/public/demo/sample-1.docx';

export function App() {
  const viewerOptions = useMemo<OfficeViewerOptionsByFormat['docx']>(() => ({
    width: 820,
    enableTextSelection: true,
    useGoogleFonts: true,
  }), []);
  const viewer = useOfficeViewer({
    format: 'docx',
    source: SOURCE,
    viewerOptions,
  });

  return (
    <main className="app">
      <div className="toolbar">
        <button onClick={viewer.zoomOut} aria-label="Zoom out">−</button>
        <button onClick={viewer.fitWidth}>Fit width</button>
        <button onClick={viewer.zoomIn} aria-label="Zoom in">+</button>
        <button onClick={viewer.reload}>Reload</button>
        <span className={viewer.error ? 'status error' : 'status'}>
          {viewer.error?.message ?? viewer.status}
        </span>
      </div>
      <div className="stage">{viewer.renderOfficeViewer()}</div>
    </main>
  );
}
