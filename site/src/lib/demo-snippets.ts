// Implementation snippets shown beside each per-format live demo. Generated
// from a small config so pptx/docx stay in sync; all checked against the API.

interface Cfg {
  Viewer: string;
  Doc: string;
  sub: 'pptx' | 'docx';
  count: string;
  render: string;
  next: string;
  prev: string;
  go: string;
}

const pptx: Cfg = {
  Viewer: 'PptxViewer', Doc: 'PptxPresentation', sub: 'pptx',
  count: 'slideCount', render: 'renderSlide', next: 'nextSlide', prev: 'prevSlide', go: 'goToSlide',
};
const docx: Cfg = {
  Viewer: 'DocxViewer', Doc: 'DocxDocument', sub: 'docx',
  count: 'pageCount', render: 'renderPage', next: 'nextPage', prev: 'prevPage', go: 'goToPage',
};

export interface DemoSnippets {
  demo: string;
  scroll: string;
  thumbnails: string;
  masterdetail: string;
}

function build(c: Cfg): DemoSnippets {
  return {
    demo: `import { ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

// The built-in viewer tracks the current ${c.sub === 'pptx' ? 'slide' : 'page'} for you.
const viewer = new ${c.Viewer}(canvas, { width: 960, useGoogleFonts: true });
await viewer.load('/sample.${c.sub}');

nextBtn.addEventListener('click', () => viewer.${c.next}());
prevBtn.addEventListener('click', () => viewer.${c.prev}());`,

    scroll: `import { ${c.Doc} } from '@silurus/ooxml/${c.sub}';

// Headless engine — render every ${c.sub === 'pptx' ? 'slide' : 'page'} into a canvas you control.
const doc = await ${c.Doc}.load('/sample.${c.sub}');

for (let i = 0; i < doc.${c.count}; i++) {
  const canvas = document.createElement('canvas');
  scroller.appendChild(canvas);
  await doc.${c.render}(canvas, i, { width: 1100 });
}`,

    thumbnails: `import { ${c.Doc} } from '@silurus/ooxml/${c.sub}';

// Render each ${c.sub === 'pptx' ? 'slide' : 'page'} small, wire up navigation.
const doc = await ${c.Doc}.load('/sample.${c.sub}');

for (let i = 0; i < doc.${c.count}; i++) {
  const thumb = document.createElement('canvas');
  thumb.addEventListener('click', () => open(i));
  grid.appendChild(thumb);
  await doc.${c.render}(thumb, i, { width: 320 });
}`,

    masterdetail: `import { ${c.Doc}, ${c.Viewer} } from '@silurus/ooxml/${c.sub}';

// A large preview viewer on the right…
const viewer = new ${c.Viewer}(detailCanvas, { width: 820, enableTextSelection: true });

// …and a thumbnail rail on the left, sharing the same file.
const [doc] = await Promise.all([
  ${c.Doc}.load('/sample.${c.sub}'),
  viewer.load('/sample.${c.sub}'),
]);

for (let i = 0; i < doc.${c.count}; i++) {
  const thumb = document.createElement('canvas');
  thumb.addEventListener('click', () => viewer.${c.go}(i));  // jump the preview
  rail.appendChild(thumb);
  await doc.${c.render}(thumb, i, { width: 200 });
}`,
  };
}

export const pptxSnippets = build(pptx);
export const docxSnippets = build(docx);

export const xlsxSheetSnippet = `import { XlsxViewer } from '@silurus/ooxml/xlsx';

// XlsxViewer owns its canvas, sheet-tab bar and zoom slider — hand it a
// container element (not a canvas). Click-drag selects a range; Ctrl/Cmd+C
// copies it as TSV.
const container = document.getElementById('sheet') as HTMLElement;
const viewer = new XlsxViewer(container, { showZoomSlider: true });

await viewer.load('/sample.xlsx');`;
