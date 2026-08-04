import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const siteRoot = new URL('.', import.meta.url);
const repositoryRoot = new URL('../../', siteRoot);

describe('framework integration guides', () => {
  it.each(['react', 'vue', 'svelte', 'solid'])('publishes an independent SEO page for %s', (framework) => {
    const pageUrl = new URL(`pages/frameworks/${framework}.astro`, siteRoot);
    expect(existsSync(pageUrl)).toBe(true);

    const page = readFileSync(pageUrl, 'utf8');
    expect(page).toContain('FrameworkGuide');
    expect(page).toContain(`framework="${framework}"`);
  });

  it('keeps Angular out of the supported framework registry', async () => {
    const { frameworkGuides } = await import('./lib/framework-guides');
    expect(frameworkGuides.map(({ id }) => id)).toEqual(['react', 'vue', 'svelte', 'solid']);
    expect(frameworkGuides.some(({ id }) => (id as string) === 'angular')).toBe(false);
  });

  it('keeps framework dependencies outside the root pnpm workspace', () => {
    const rootWorkspace = readFileSync(new URL('pnpm-workspace.yaml', repositoryRoot), 'utf8');
    const exampleWorkspace = readFileSync(new URL('examples/frameworks/pnpm-workspace.yaml', repositoryRoot), 'utf8');

    expect(rootWorkspace).not.toContain('examples/frameworks');
    expect(exampleWorkspace).toContain('- react');
    expect(exampleWorkspace).toContain('- vue');
    expect(exampleWorkspace).toContain('- svelte');
    expect(exampleWorkspace).toContain('- solid');
  });

  it('uses search-oriented titles without making them identical', async () => {
    const { frameworkGuides } = await import('./lib/framework-guides');
    const titles = frameworkGuides.map(({ title }) => title);

    expect(new Set(titles).size).toBe(4);
    for (const title of titles) {
      expect(title).toMatch(/^How to render Office files in the browser with /);
      expect(title).toContain('DOCX, XLSX, and PPTX');
    }
  });

  it('keeps every integration module portable outside the examples workspace', async () => {
    const { frameworkGuides } = await import('./lib/framework-guides');
    for (const guide of frameworkGuides) {
      expect(guide.adapterCode).not.toContain('@ooxml-framework-examples');
      expect(guide.adapterCode).not.toContain('../shared');
      expect(guide.adapterCode).toContain("from '@silurus/ooxml/docx'");
      expect(guide.adapterCode).toContain("from '@silurus/ooxml/xlsx'");
      expect(guide.adapterCode).toContain("from '@silurus/ooxml/pptx'");
      expect(guide.adapterCode).not.toContain("await import('@silurus/ooxml");
      expect(guide.adapterCode).toContain('destroy');
      expect(guide.adapterCode).not.toMatch(/\blet\s+/);
      expect(guide.adapterCode).not.toMatch(/\bvoid\s+[A-Za-z_(]/);
    }
  });

  it('implements the React integration as a render hook with an internal ref and cleanup', () => {
    const hook = readFileSync(new URL('examples/frameworks/react/src/useOfficeViewer.tsx', repositoryRoot), 'utf8');
    expect(hook).toContain('const mountRef = useRef<HTMLDivElement>(null);');
    expect(hook).toContain('const renderOfficeViewer = useCallback(');
    expect(hook).toContain('<div {...props} ref={mountRef} />');
    expect(hook).toContain('useEffect(() =>');
    expect(hook).toContain('const controller = new AbortController();');
    expect(hook).toContain('viewer.destroy();');
    expect(hook).not.toMatch(/\blet\s+/);
    expect(hook).not.toMatch(/\bvoid\s+mountOfficeViewer/);
    expect(hook).not.toContain('targetRef:');
  });

  it('keeps viewer construction out of components and accepts a replaceable local file source', async () => {
    const { frameworkGuides } = await import('./lib/framework-guides');
    for (const guide of frameworkGuides) {
      expect(guide.appCode).not.toContain("import('@silurus/ooxml");
      expect(guide.appCode).not.toContain('DocxScrollViewer');
      expect(guide.appCode).not.toContain('PptxScrollViewer');
      expect(guide.appCode).not.toContain('XlsxViewer');
      expect(guide.appCode).toContain('file.arrayBuffer()');
      expect(guide.appCode).toContain('.docx,.xlsx,.pptx');
      expect(guide.appCode).toContain("'Choose an Office file'");
      expect(guide.appCode).not.toContain('raw.githubusercontent.com');
      expect(guide.appCode).not.toContain('<canvas');
    }
  });

  it('embeds a local-file try surface without an initial sample', () => {
    const guide = readFileSync(new URL('components/FrameworkGuide.astro', siteRoot), 'utf8');
    const tryYours = readFileSync(new URL('components/FrameworkTryYours.astro', siteRoot), 'utf8');

    expect(guide).toContain('FrameworkTryYours');
    expect(guide).not.toContain('LiveShowcase');
    expect(tryYours).toContain('accept=".docx,.xlsx,.pptx"');
    expect(tryYours).toContain('Your file stays in this browser.');
    expect(tryYours).not.toContain('sample-1');
  });

  it('documents and performs DOCX teardown in the existing snippets', () => {
    const snippets = readFileSync(new URL('lib/demo-snippets.ts', siteRoot), 'utf8');
    expect(snippets).toContain("const fwDocx: FwCfg = { Viewer: 'DocxViewer'");
    expect(snippets).toContain('return () => viewer.destroy();');
    expect(snippets).toContain('onBeforeUnmount(() => viewer?.destroy());');
    expect(snippets).not.toMatch(/\bvoid\s+viewer\.load/);
    expect(snippets).not.toContain('destroy: false');
    expect(snippets).not.toContain('docx renders into the canvas you own and needs no teardown');
  });
});
