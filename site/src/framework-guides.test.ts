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

  it('keeps one format-aware mount implementation shared by every adapter', () => {
    const common = readFileSync(new URL('examples/frameworks/shared/src/office-viewer.ts', repositoryRoot), 'utf8');
    expect(common).toContain("import('@silurus/ooxml/docx')");
    expect(common).toContain("import('@silurus/ooxml/xlsx')");
    expect(common).toContain("import('@silurus/ooxml/pptx')");
    expect(common).toContain('viewer.destroy();');
  });

  it('implements the React integration as a render hook with an internal ref and cleanup', () => {
    const hook = readFileSync(new URL('examples/frameworks/react/src/useOfficeViewer.tsx', repositoryRoot), 'utf8');
    expect(hook).toContain('const mountRef = useRef<HTMLDivElement>(null);');
    expect(hook).toContain('const renderOfficeViewer = useCallback(');
    expect(hook).toContain('<div {...props} ref={mountRef} />');
    expect(hook).toContain('useEffect(() =>');
    expect(hook).toContain('mountedViewer?.destroy();');
    expect(hook).not.toContain('targetRef:');
  });

  it('documents and performs DOCX teardown in the existing snippets', () => {
    const snippets = readFileSync(new URL('lib/demo-snippets.ts', siteRoot), 'utf8');
    expect(snippets).toContain("const fwDocx: FwCfg = { Viewer: 'DocxViewer'");
    expect(snippets).toContain('return () => viewer.destroy();');
    expect(snippets).toContain('onBeforeUnmount(() => viewer?.destroy());');
    expect(snippets).not.toContain('destroy: false');
    expect(snippets).not.toContain('docx renders into the canvas you own and needs no teardown');
  });
});
