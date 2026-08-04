import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

const siteRoot = new URL('.', import.meta.url);
const repositoryRoot = new URL('../../', siteRoot);
const exampleFiles = {
  react: { integration: 'react/src/useOfficeViewer.tsx', app: 'react/src/App.tsx' },
  vue: { integration: 'vue/src/useOfficeViewer.ts', app: 'vue/src/App.vue' },
  svelte: { integration: 'svelte/src/createOfficeViewer.ts', app: 'svelte/src/App.svelte' },
  solid: { integration: 'solid/src/createOfficeViewer.ts', app: 'solid/src/App.tsx' },
} as const;
const exampleSource = (path: string) => readFileSync(
  new URL(`examples/frameworks/${path}`, repositoryRoot),
  'utf8',
);

describe('framework integration guides', () => {
  it.each(['react', 'vue', 'svelte', 'solid'])('publishes an independent SEO page for %s', (framework) => {
    const pageUrl = new URL(`pages/frameworks/${framework}.astro`, siteRoot);
    expect(existsSync(pageUrl)).toBe(true);

    const page = readFileSync(pageUrl, 'utf8');
    expect(page).toContain('FrameworkGuide');
    expect(page).toContain(`framework="${framework}"`);
  });

  it('uses a framework chooser as the navigation destination', () => {
    const nav = readFileSync(new URL('components/Nav.astro', siteRoot), 'utf8');
    const index = readFileSync(new URL('pages/frameworks/index.astro', siteRoot), 'utf8');

    expect(nav).toContain('href="/frameworks"');
    expect(nav).not.toContain('href="/frameworks/react"');
    expect(nav.indexOf('href="/try"')).toBeLessThan(nav.indexOf('href="/frameworks"'));
    expect(index).toContain('frameworkGuides.map');
    expect(index).toContain('React, Vue, Svelte, and Solid');
  });

  it('keeps Storybook as a development tool rather than a public-site destination', () => {
    const publicNavigation = [
      'components/Nav.astro',
      'layouts/FormatPage.astro',
      'pages/index.astro',
      'pages/try.astro',
    ].map((path) => readFileSync(new URL(path, siteRoot), 'utf8')).join('\n');

    expect(publicNavigation).not.toContain('/storybook/');
    expect(publicNavigation).not.toContain('>Storybook<');
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
      const integration = exampleSource(exampleFiles[guide.id].integration);
      expect(integration).not.toContain('@ooxml-framework-examples');
      expect(integration).not.toContain('../shared');
      expect(integration).toContain("from '@silurus/ooxml/docx'");
      expect(integration).toContain("from '@silurus/ooxml/xlsx'");
      expect(integration).toContain("from '@silurus/ooxml/pptx'");
      expect(integration).not.toContain("await import('@silurus/ooxml");
      expect(integration).toContain('destroy');
      expect(integration).not.toMatch(/\blet\s+/);
      expect(integration).not.toMatch(/\bvoid\s+[A-Za-z_(]/);
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
      const app = exampleSource(exampleFiles[guide.id].app);
      expect(app).not.toContain("import('@silurus/ooxml");
      expect(app).not.toContain("from '@silurus/ooxml");
      expect(app).not.toContain('DocxScrollViewer');
      expect(app).not.toContain('PptxScrollViewer');
      expect(app).not.toContain('XlsxViewer');
      expect(app).toContain('file.arrayBuffer()');
      expect(app).toContain('.docx,.xlsx,.pptx');
      expect(app).toContain("'Choose an Office file'");
      expect(app).not.toContain('raw.githubusercontent.com');
      expect(app).not.toContain('<canvas');
    }
  });

  it('embeds each runnable project with StackBlitz', async () => {
    const guide = readFileSync(new URL('components/FrameworkGuide.astro', siteRoot), 'utf8');
    const stackBlitz = readFileSync(new URL('components/FrameworkStackBlitz.astro', siteRoot), 'utf8');
    const { frameworkGuides } = await import('./lib/framework-guides');

    expect(guide).toContain('FrameworkStackBlitz');
    expect(guide).not.toContain('LiveShowcase');
    expect(guide).not.toContain('CodeTabs');
    expect(guide).not.toContain('Integration module');
    expect(guide).not.toContain('<h2>Component</h2>');
    expect(guide).not.toContain('Step 1');
    expect(guide).not.toContain('Step 2');
    expect(guide).not.toContain('Open live');
    expect(stackBlitz).toContain('<iframe');
    for (const framework of frameworkGuides) {
      expect(framework.stackBlitzEmbedUrl).toContain('stackblitz.com/github/');
      expect(framework.stackBlitzEmbedUrl).toContain('embed=1');
      expect(framework.stackBlitzEmbedUrl).toContain(`startScript=dev%3A${framework.id}`);
    }
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
