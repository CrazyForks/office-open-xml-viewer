import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

export type FrameworkId = 'react' | 'vue' | 'svelte' | 'solid';

export interface FrameworkGuide {
  id: FrameworkId;
  name: string;
  title: string;
  description: string;
  integrationName: string;
  integrationSummary: string;
  lifecycle: string;
  adapterFilename: string;
  adapterLanguage: 'ts' | 'tsx';
  adapterCode: string;
  appFilename: string;
  appLanguage: 'vue' | 'svelte' | 'tsx';
  appCode: string;
  stackBlitzUrl: string;
}

const examplesRoot = [
  resolve(process.cwd(), 'examples/frameworks'),
  resolve(process.cwd(), '../examples/frameworks'),
].find(existsSync);
if (!examplesRoot) throw new Error('Cannot locate examples/frameworks from the current build directory.');
const source = (path: string): string => readFileSync(resolve(examplesRoot, path), 'utf8').trim();
const repository = 'yukiyokotani/office-open-xml-viewer';

function stackBlitz(framework: FrameworkId, file: string): string {
  const params = new URLSearchParams({ file, startScript: `dev:${framework}` });
  return `https://stackblitz.com/github/${repository}/tree/main/examples/frameworks?${params}`;
}

export const frameworkGuides: FrameworkGuide[] = [
  {
    id: 'react',
    name: 'React',
    title: 'How to render Office files in the browser with React | DOCX, XLSX, and PPTX',
    description: 'Render DOCX, XLSX, and PPTX Office files in a React and TypeScript application with a reusable custom hook, Canvas viewer controls, and correct effect cleanup.',
    integrationName: 'useOfficeViewer',
    integrationSummary: 'A custom render hook owns its container ref and viewer instance. The component supplies only a replaceable format and source, and receives stable zoom, fit, reload, status, and error controls.',
    lifecycle: 'useEffect mirrors viewer setup with destroy() cleanup, including React Strict Mode’s development setup/cleanup cycle.',
    adapterFilename: 'useOfficeViewer.tsx',
    adapterLanguage: 'tsx',
    adapterCode: source('react/src/useOfficeViewer.tsx'),
    appFilename: 'App.tsx',
    appLanguage: 'tsx',
    appCode: source('react/src/App.tsx'),
    stackBlitzUrl: stackBlitz('react', 'react/src/useOfficeViewer.tsx'),
  },
  {
    id: 'vue',
    name: 'Vue',
    title: 'How to render Office files in the browser with Vue | DOCX, XLSX, and PPTX',
    description: 'Render DOCX, XLSX, and PPTX Office files in Vue 3 with TypeScript, a reusable Composition API composable, template refs, viewer controls, and automatic cleanup.',
    integrationName: 'useOfficeViewer',
    integrationSummary: 'A Composition API composable watches the template ref, format, and replaceable source while exposing readonly status and common viewer controls.',
    lifecycle: 'watchEffect invalidation destroys the previous viewer before a new source or format mounts, and also runs when the component scope is disposed.',
    adapterFilename: 'useOfficeViewer.ts',
    adapterLanguage: 'ts',
    adapterCode: source('vue/src/useOfficeViewer.ts'),
    appFilename: 'App.vue',
    appLanguage: 'vue',
    appCode: source('vue/src/App.vue'),
    stackBlitzUrl: stackBlitz('vue', 'vue/src/useOfficeViewer.ts'),
  },
  {
    id: 'svelte',
    name: 'Svelte',
    title: 'How to render Office files in the browser with Svelte | DOCX, XLSX, and PPTX',
    description: 'Render DOCX, XLSX, and PPTX Office files in Svelte with TypeScript, a reusable action, readable status stores, viewer controls, and reliable teardown.',
    integrationName: 'createOfficeViewer',
    integrationSummary: 'A self-contained Svelte action mounts on one div, reacts to a replaceable format and source, and returns readable status stores plus imperative controls.',
    lifecycle: 'The action update and destroy callbacks replace or release the viewer, while a generation token rejects stale asynchronous mounts.',
    adapterFilename: 'createOfficeViewer.ts',
    adapterLanguage: 'ts',
    adapterCode: source('svelte/src/createOfficeViewer.ts'),
    appFilename: 'App.svelte',
    appLanguage: 'svelte',
    appCode: source('svelte/src/App.svelte'),
    stackBlitzUrl: stackBlitz('svelte', 'svelte/src/createOfficeViewer.ts'),
  },
  {
    id: 'solid',
    name: 'Solid',
    title: 'How to render Office files in the browser with Solid | DOCX, XLSX, and PPTX',
    description: 'Render DOCX, XLSX, and PPTX Office files in Solid with TypeScript, a reusable reactive primitive, signal-based status, viewer controls, and cleanup.',
    integrationName: 'createOfficeViewer',
    integrationSummary: 'A self-contained Solid primitive reads target, format, and source accessors, exposes status and error signals, and returns common viewer controls.',
    lifecycle: 'createEffect tracks reactive inputs and onCleanup releases each viewer before the effect reruns or the owner is disposed.',
    adapterFilename: 'createOfficeViewer.ts',
    adapterLanguage: 'ts',
    adapterCode: source('solid/src/createOfficeViewer.ts'),
    appFilename: 'App.tsx',
    appLanguage: 'tsx',
    appCode: source('solid/src/App.tsx'),
    stackBlitzUrl: stackBlitz('solid', 'solid/src/createOfficeViewer.ts'),
  },
];

export function getFrameworkGuide(id: FrameworkId): FrameworkGuide {
  const guide = frameworkGuides.find((candidate) => candidate.id === id);
  if (!guide) throw new Error(`Unknown framework guide: ${id}`);
  return guide;
}
