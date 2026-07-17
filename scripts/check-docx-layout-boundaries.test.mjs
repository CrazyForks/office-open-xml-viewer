import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const checker = resolve(import.meta.dirname, 'check-docx-layout-boundaries.mjs');

function write(root, path, contents) {
  const absolute = join(root, path);
  mkdirSync(resolve(absolute, '..'), { recursive: true });
  writeFileSync(absolute, contents);
}

function command(root, executable, args) {
  const result = spawnSync(executable, args, { cwd: root, encoding: 'utf8' });
  return {
    status: result.status,
    output: `${result.stdout ?? ''}${result.stderr ?? ''}`,
  };
}

function runChecker(root, ...args) {
  return command(root, process.execPath, [checker, '--root', root, ...args]);
}

function expectDiagnostic(root, diagnostic, message, ...args) {
  const result = runChecker(root, ...args);
  assert.notEqual(result.status, 0, message ?? result.output);
  assert.match(result.output, new RegExp(`^${diagnostic}:`, 'm'), message ?? result.output);
  return result;
}

function git(root, ...args) {
  const result = command(root, 'git', args);
  assert.equal(result.status, 0, result.output);
}

function validEmptyBaseline() {
  return {
    version: 2,
    legacySymbolCounts: {},
    migrationIdentifierCounts: {},
    nonLayoutDeclarationKeys: [],
    legacyDeclarationHashes: {},
    rendererImportEdges: [],
  };
}

const canonicalRenderer = `
import { layoutDocument } from './layout/document.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
import { paintLayoutPage } from './paint/canvas-page.js';
function createConcreteBodyLayoutKernel(doc, ctx, localMetrics) {
  return { doc, ctx, localMetrics };
}
function createLayoutServices(doc, ctx, localMetrics) {
  const services = {};
  attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx, localMetrics));
  return services;
}
export function renderDocumentToCanvas(services, input, pageIndex) {
  const selection = selectDocumentLayoutPage(services, input, pageIndex);
  layoutDocument();
  return paintLayoutPage(selection.page);
}
`;

function initializeCanonicalFixture(prefix = 'docx-layout-boundary-canonical-') {
  const root = mkdtempSync(join(tmpdir(), prefix));
  write(root, 'packages/docx/src/layout/plain-data.ts',
    'export function snapshotPlainData(value) { return value; }\n');
  write(root, 'packages/docx/src/layout/retained-geometry-translation.ts',
    "import type { PointPt } from './types.js';\nexport const translate = (point: PointPt) => point;\n");
  write(root, 'packages/docx/src/layout/occurrence-projection.ts',
    "import { snapshotPlainData } from './plain-data.js';\nexport const project = snapshotPlainData;\n");
  write(root, 'packages/docx/src/layout/types.ts',
    'export interface PointPt { xPt: number; yPt: number }\n');
  write(root, 'packages/docx/src/layout/coordinate-space.ts',
    "import type { PointPt } from './types.js';\nexport const coordinate = (point: PointPt) => point;\n");
  write(root, 'packages/docx/src/layout/page-graph.ts', 'export const PAGE_LAYER_IDS = [];\n');
  write(root, 'packages/docx/src/layout/page-factory.ts',
    "import { coordinate } from './coordinate-space.js';\n"
      + "import { PAGE_LAYER_IDS } from './page-graph.js';\n"
      + "import type { BodyOccurrenceDestination } from './occurrence-projection.js';\n"
      + 'export const pageFactory = [coordinate, PAGE_LAYER_IDS] satisfies unknown;\n'
      + 'export type Destination = BodyOccurrenceDestination;\n');
  write(root, 'packages/docx/src/layout/invariants.ts',
    'export function assertDocumentLayout(value) { return value; }\n'
      + 'export function deepFreezeDocumentLayout(value) { return Object.freeze(value); }\n');
  write(root, 'packages/docx/src/layout/body-paginator.ts',
    "import { assertDocumentLayout, deepFreezeDocumentLayout } from './invariants.js';\n"
      + 'export function paginateBody(input, services, options) {\n'
      + '  const layout = { pages: [], diagnostics: [], input, services, options };\n'
      + '  assertDocumentLayout(layout);\n'
      + '  return deepFreezeDocumentLayout(layout);\n'
      + '}\n');
  write(root, 'packages/docx/src/layout/document.ts',
    "import { paginateBody } from './body-paginator.js';\n"
      + 'export function layoutDocument(input, services, options) {\n'
      + '  return paginateBody(input, services, options);\n'
      + '}\n');
  write(root, 'packages/docx/src/layout/document-layout-variants.ts',
    'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
      + '  const store = layoutVariantStoreOf(services);\n'
      + "  if (!store) throw new Error('Document layout variant store is not attached');\n"
      + '  return store.selectPage(layoutOptionsForRender(input), pageIndex);\n'
      + '}\n');
  write(root, 'packages/docx/src/paint/canvas-page.ts',
    'export function paintLayoutPage(page) { return page; }\n');
  write(root, 'packages/docx/src/render-worker.ts',
    "import { renderDocumentToCanvas, layoutDocument } from './renderer.js';\n"
      + "import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';\n"
      + 'export function initializeWorker(model, layoutServices) {\n'
      + '  return attachDocumentLayoutVariants({ model, services: layoutServices,\n'
      + '    buildLayout: (options) => layoutDocument(model, layoutServices, options) });\n'
      + '}\n'
      + 'export function renderWorkerPage(model, canvas, pageIndex, options) {\n'
      + '  return renderDocumentToCanvas(model, canvas, pageIndex, { ...options, layoutServices });\n'
      + '}\n'
      + 'export function collectWorkerRuns(model, canvas, pageIndex, options) {\n'
      + '  return renderDocumentToCanvas(model, canvas, pageIndex, { ...options, layoutServices });\n'
      + '}\n');
  write(root, 'packages/docx/src/renderer.ts', canonicalRenderer);
  return root;
}

function initializeLegacyStoryPaintFixture() {
  const root = initializeCanonicalFixture('docx-layout-boundary-legacy-story-');
  write(root, 'packages/docx/src/line-layout.ts', 'export function layoutLines() {}\n');
  write(root, 'packages/docx/src/paint/canvas-text.ts',
    'export function paintPlacedParagraphLayout() {}\n');
  write(root, 'packages/docx/src/renderer.ts', `
import { paintPlacedParagraphLayout } from './paint/canvas-text.js';
import { layoutLines } from './line-layout.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
function createConcreteBodyLayoutKernel(doc, ctx, localMetrics) { return { doc, ctx, localMetrics }; }
function createLayoutServices(doc, ctx, localMetrics) {
  const services = {};
  attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx, localMetrics));
  return services;
}
function renderBodyElements(elements) {
  for (const element of elements) {
    if (element.type === 'paragraph') paintPlacedParagraphLayout(element);
  }
}
function renderDocumentToCanvas(services, input, pageIndex) {
  return selectDocumentLayoutPage(services, input, pageIndex);
}
function renderHeaderFooterStory(story) { return layoutLines(story); }
`);
  command(root, 'git', ['init', '-b', 'main']);
  command(root, 'git', ['config', 'user.email', 'boundary-test@example.invalid']);
  command(root, 'git', ['config', 'user.name', 'Boundary Test']);
  command(root, 'git', ['add', '.']);
  command(root, 'git', ['commit', '-m', 'canonical fixture']);
  const branch = command(root, 'git', ['switch', '-c', 'boundary']);
  assert.equal(branch.status, 0, branch.output);
  const baseline = runChecker(root, '--write-transitional-baseline', '--base-ref', 'main');
  assert.equal(baseline.status, 0, baseline.output);
  return root;
}

const canonicalBodyLayoutAdapter =
  "import { bodyLayoutAcquisitionInput } from './parser-model.js';\n"
  + "import type { DocxDocumentModel } from './types.js';\n"
  + "import { projectBodyLayoutInput, type BodyLayoutInput } from './layout/body-layout-input.js';\n"
  + 'export function createBodyLayoutInput(document: DocxDocumentModel): BodyLayoutInput {\n'
  + '  return projectBodyLayoutInput(bodyLayoutAcquisitionInput(document));\n'
  + '}\n';

function initializeBodyLayoutAdapterFixture(prefix = 'docx-layout-boundary-input-adapter-') {
  const root = initializeCanonicalFixture(prefix);
  write(root, 'packages/docx/src/parser-model.ts',
    'export function bodyLayoutAcquisitionInput(document) { return document; }\n');
  write(root, 'packages/docx/src/types.ts', 'export interface DocxDocumentModel {}\n');
  write(root, 'packages/docx/src/layout/body-layout-input.ts',
    'export interface BodyLayoutInput {}\n'
      + 'export function projectBodyLayoutInput(value) { return value; }\n');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter);
  return root;
}

function initializeTransitionalRepository() {
  const root = initializeCanonicalFixture('docx-layout-boundary-transitional-');
  write(root, 'packages/docx/src/renderer.ts',
    canonicalRenderer.replace(
      '\nexport function renderDocumentToCanvas',
      '\nfunction renderShapeText() { return 1; }\nexport function renderDocumentToCanvas',
    ));
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'canonical transition base');
  git(root, 'switch', '-c', 'boundary');
  const baseline = runChecker(root, '--write-transitional-baseline', '--base-ref', 'main');
  assert.equal(baseline.status, 0, baseline.output);
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'establish transition baseline');
  git(root, 'switch', 'main');
  git(root, 'merge', '--ff-only', 'boundary');
  git(root, 'switch', '-c', 'next');
  return root;
}

test('accepts the final canonical producer, retained model, selected variant, and worker route', () => {
  const root = initializeCanonicalFixture();
  const result = runChecker(root, '--final');
  assert.equal(result.status, 0, result.output);
});

test('requires one private concrete body-kernel owner with exact loud attachment', () => {
  for (const [name, source] of [
    ['missing implementation', canonicalRenderer.replace(
      /function createConcreteBodyLayoutKernel[\s\S]*?\n}\n/,
      '',
    )],
    ['missing attachment', canonicalRenderer.replace(
      'attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx, localMetrics));',
      'createConcreteBodyLayoutKernel(doc, ctx, localMetrics);',
    )],
    ['wrong owner arguments', canonicalRenderer.replace(
      'attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx, localMetrics));',
      'attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx));',
    )],
    ['duplicate owner', canonicalRenderer.replace(
      'return services;',
      'attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx, localMetrics));\n  return services;',
    )],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-owner-${name}-`);
    write(root, 'packages/docx/src/renderer.ts', source);
    expectDiagnostic(root, 'BODY_KERNEL_SERVICE_OWNER', name, '--final');
  }
});

test('rejects every deleted body producer, raw page route, and test adapter symbol', () => {
  const deleted = [
    'computePages',
    'paginateDocument',
    'paginateWithHeaderFooterReserve',
    'physicalPageSizeForPage',
    'PaginatedBodyElement',
    'prebuiltPages',
    'retainedLayout',
    'bodyFragmentFor',
    'sectionBreakSpacer',
    'collapsedSpacer',
    'leadsCollapsedRun',
    'hiddenCollapsed',
  ];
  for (const name of deleted) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-deleted-${name}-`);
    write(root, 'packages/docx/src/layout/deleted-path.ts', `export const ${name} = true;\n`);
    expectDiagnostic(root, 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER', name, '--final');
  }
});

test('rejects migration flags and silent alternate layout fallbacks', () => {
  for (const [name, diagnostic] of [
    ['useOldLayoutEngine', 'FINAL_LEGACY_BOUNDARY'],
    ['preferAlternateLayoutPath', 'FINAL_LEGACY_BOUNDARY'],
    ['bodyLayoutFallback', 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-migration-${name}-`);
    write(root, 'packages/docx/src/layout/path.ts', `export const ${name} = true;\n`);
    expectDiagnostic(root, diagnostic, name, '--final');
  }
});

test('coordinate-space and page-factory accept only their explicit dependencies', () => {
  assert.equal(runChecker(initializeCanonicalFixture(), '--final').status, 0);
  for (const [name, path, source] of [
    ['renderer', 'coordinate-space.ts', "import { value } from '../renderer.js';\nexport const coordinate = value;\n"],
    ['parser', 'coordinate-space.ts', "import { value } from '../parser-model.js';\nexport const coordinate = value;\n"],
    ['paint', 'coordinate-space.ts', "import { value } from '../paint/canvas-page.js';\nexport const coordinate = value;\n"],
    ['dynamic', 'coordinate-space.ts', "export const coordinate = import('./types.js');\n"],
    ['dynamic nonliteral', 'coordinate-space.ts', "const moduleName = './types.js';\nexport const coordinate = import(moduleName);\n"],
    ['decorated', 'coordinate-space.ts', "import type { PointPt } from './types.js?raw';\nexport const coordinate = (point: PointPt) => point;\n"],
    ['worker', 'coordinate-space.ts', "import { value } from '../render-worker.js';\nexport const coordinate = value;\n"],
    ['shaping', 'coordinate-space.ts', "import { value } from './text.js';\nexport const coordinate = value;\n"],
    ['package', 'coordinate-space.ts', "import value from 'canvas';\nexport const coordinate = value;\n"],
    ['projection runtime', 'page-factory.ts', "import { project } from './occurrence-projection.js';\nexport const pageFactory = project;\n"],
    ['page factory decorated type', 'page-factory.ts', "import type { PointPt } from './types.js?raw';\nexport type Point = PointPt;\n"],
    ['page factory package runtime', 'page-factory.ts', "import value from 'canvas';\nexport const pageFactory = value;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-coordinate-${name}-`);
    write(root, `packages/docx/src/layout/${path}`, source);
    expectDiagnostic(root, 'COORDINATE_SPACE_RUNTIME_DEPENDENCY', name, '--final');
  }
});

test('page-factory may retain section decoration geometry through its dedicated layout helper', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-page-decoration-');
  write(root, 'packages/docx/src/layout/column-separators.ts',
    'export const columnSeparatorSegments = () => [];\n');
  write(root, 'packages/docx/src/layout/page-factory.ts',
    "import { coordinate } from './coordinate-space.js';\n"
      + "import { columnSeparatorSegments } from './column-separators.js';\n"
      + "import { PAGE_LAYER_IDS } from './page-graph.js';\n"
      + 'export const pageFactory = [coordinate, columnSeparatorSegments, PAGE_LAYER_IDS] satisfies unknown;\n');

  assert.equal(runChecker(root, '--final').status, 0);
});

test('occurrence projection accepts only translation and plain-data runtime dependencies', () => {
  for (const [name, path, source] of [
    ['external layout', 'occurrence-projection.ts', "import { value } from '../paragraph-measure.js';\nexport const project = value;\n"],
    ['reverse edge', 'retained-geometry-translation.ts', "import { project } from './occurrence-projection.js';\nexport const translate = project;\n"],
    ['parser edge', 'plain-data.ts', "import { value } from '../parser-model.js';\nexport const snapshotPlainData = value;\n"],
    ['dynamic edge', 'occurrence-projection.ts', "export const project = import('./plain-data.js');\n"],
    ['dynamic nonliteral edge', 'occurrence-projection.ts', "const moduleName = './plain-data.js';\nexport const project = import(moduleName);\n"],
    ['decorated edge', 'occurrence-projection.ts', "import { snapshotPlainData } from './plain-data.js?raw';\nexport const project = snapshotPlainData;\n"],
    ['package edge', 'occurrence-projection.ts', "import value from 'measurement-service';\nexport const project = value;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-occurrence-${name}-`);
    write(root, `packages/docx/src/layout/${path}`, source);
    expectDiagnostic(root, 'OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', name, '--final');
  }
});

test('coordinate-space and page-factory seams are mandatory', () => {
  for (const name of ['coordinate-space.ts', 'page-factory.ts']) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-coordinate-missing-${name}-`);
    rmSync(join(root, 'packages/docx/src/layout', name));
    expectDiagnostic(root, 'COORDINATE_SPACE_RUNTIME_DEPENDENCY', name, '--final');
  }
});

test('every occurrence-projection seam is mandatory', () => {
  for (const name of [
    'occurrence-projection.ts',
    'retained-geometry-translation.ts',
    'plain-data.ts',
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-occurrence-missing-${name}-`);
    rmSync(join(root, 'packages/docx/src/layout', name));
    expectDiagnostic(root, 'OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', name, '--final');
  }
});

test('keeps layout display-independent and paint measurement-free', () => {
  for (const [name, path, source, diagnostic] of [
    ['paint measurement', 'paint/canvas-page.ts', 'export function paint(ctx) { return ctx.measureText("x"); }\n', 'PAINT_CAPABILITY'],
    ['paint style cascade', 'paint/canvas-page.ts', 'export function resolveRunStyle() {}\n', 'PAINT_CAPABILITY'],
    ['layout display scale', 'layout/document.ts', 'export const displayScale = 2;\n', 'LAYOUT_DISPLAY_CAPABILITY'],
    ['layout canvas context', 'layout/document.ts', 'export let CanvasRenderingContext2D;\n', 'LAYOUT_DISPLAY_CAPABILITY'],
    ['layout style cascade', 'layout/document.ts', 'export function mergeRunProperties() {}\n', 'LAYOUT_STYLE_CAPABILITY'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-capability-${name}-`);
    write(root, `packages/docx/src/${path}`, source);
    expectDiagnostic(root, diagnostic, name, '--final');
  }
});

test('paint imports only retained layout contracts and reviewed atomic core painters', () => {
  const allowed = initializeCanonicalFixture('docx-layout-boundary-paint-allowed-');
  write(allowed, 'packages/docx/src/paint/helper.ts',
    "import type { PointPt } from '../layout/types.js';\n"
      + "import { crispOffset } from '@silurus/ooxml-core';\n"
      + 'export const helper = [crispOffset] satisfies unknown as PointPt[];\n');
  assert.equal(runChecker(allowed, '--final').status, 0);

  for (const [name, source] of [
    ['layout implementation', "import { paginateBody } from '../layout/body-paginator.js';\nexport const helper = paginateBody;\n"],
    ['measurement module', "import { measure } from '../paragraph-measure.js';\nexport const helper = measure;\n"],
    ['forbidden core API', "import { measureText } from '@silurus/ooxml-core';\nexport const helper = measureText;\n"],
    ['aliased core API', "import { crispOffset as crisp } from '@silurus/ooxml-core';\nexport const helper = crisp;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-paint-${name}-`);
    write(root, 'packages/docx/src/paint/helper.ts', source);
    expectDiagnostic(root, 'FORBIDDEN_PAINT_EDGE', name, '--final');
  }
});

test('every reviewed atomic core paint binding remains explicitly allowed', () => {
  const valueBindings = [
    'autoContrastColor',
    'canvasFontString',
    'crispOffset',
    'drawImageCropped',
    'doubleRailGeometry',
    'fillDoubleBorder',
    'paintDrawingMLShape',
    'resolveFill',
    'renderChart',
  ];
  for (const name of valueBindings) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-core-${name}-`);
    write(root, 'packages/docx/src/paint/core-binding.ts',
      `import { ${name} } from '@silurus/ooxml-core';\nvoid ${name};\n`);
    const result = runChecker(root, '--final');
    assert.equal(result.status, 0, `${name}: ${result.output}`);
  }
  const typeRoot = initializeCanonicalFixture('docx-layout-boundary-core-hyperlink-');
  write(typeRoot, 'packages/docx/src/paint/core-binding.ts',
    "import type { HyperlinkTarget } from '@silurus/ooxml-core';\nexport type Target = HyperlinkTarget;\n");
  const result = runChecker(typeRoot, '--final');
  assert.equal(result.status, 0, result.output);
});

test('paint boundary follows transitive edges and rejects sibling or public package imports', () => {
  const cases = [
    ['transitive local', (root) => {
      write(root, 'packages/docx/src/paint/helper.ts',
        "import { measure } from '../paragraph-measure.js';\nexport const helper = measure;\n");
      write(root, 'packages/docx/src/paragraph-measure.ts', 'export const measure = 1;\n');
      return "import { helper } from './helper.js';\nexport const paint = helper;\n";
    }],
    ['sibling package', () => "import { render } from '@silurus/ooxml-pptx';\nexport const paint = render;\n"],
    ['public package', () => "import { render } from '@silurus/ooxml';\nexport const paint = render;\n"],
    ['bare side effect', () => "import '@silurus/ooxml-core';\nexport const paint = true;\n"],
  ];
  for (const [name, source] of cases) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-paint-graph-${name}-`);
    write(root, 'packages/docx/src/paint/graph.ts', source(root));
    expectDiagnostic(root, 'FORBIDDEN_PAINT_EDGE', name, '--final');
  }
});

test('paint cannot import the retained page graph implementation', () => {
  const rejected = initializeCanonicalFixture('docx-layout-boundary-page-graph-');
  write(rejected, 'packages/docx/src/paint/graph.ts',
    "import { PAGE_LAYER_IDS } from '../layout/page-graph.js';\nexport const paint = PAGE_LAYER_IDS;\n");
  expectDiagnostic(rejected, 'FORBIDDEN_PAINT_EDGE', 'page graph edge', '--final');
});

test('layout-only treatment stays outside paint while layout resource implementations stay forbidden', () => {
  const isolated = initializeCanonicalFixture('docx-layout-boundary-layout-only-treatment-');
  write(isolated, 'packages/docx/src/layout/border-treatment.ts',
    "import { docxBorderDashArray } from '@silurus/ooxml-core';\n"
      + "export const treatment = docxBorderDashArray('dotDash', 1);\n");
  const isolatedResult = runChecker(isolated, '--final');
  assert.equal(isolatedResult.status, 0, isolatedResult.output);

  for (const [name, source] of [
    ['value', "import { createRegistry } from '../layout/paint-resources.js';\nexport const paint = createRegistry();\n"],
    ['type', "import type { Registry } from '../layout/paint-resources.js';\nexport type PaintRegistry = Registry;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-layout-resource-${name}-`);
    write(root, 'packages/docx/src/layout/paint-resources.ts',
      'export interface Registry {}\nexport function createRegistry() { return {}; }\n');
    write(root, 'packages/docx/src/paint/resource.ts', source);
    expectDiagnostic(root, 'FORBIDDEN_PAINT_EDGE', name, '--final');
  }
});

test('computed paint measurement and TSX layout display inputs are audited', () => {
  const paint = initializeCanonicalFixture('docx-layout-boundary-computed-measurement-');
  write(paint, 'packages/docx/src/paint/computed.ts',
    "export const width = (ctx) => ctx['measureText']('x').width;\n");
  expectDiagnostic(paint, 'PAINT_CAPABILITY', 'computed measureText', '--final');

  const layout = initializeCanonicalFixture('docx-layout-boundary-layout-tsx-');
  write(layout, 'packages/docx/src/layout/display.tsx',
    'export const Layout = ({ dpr }: { dpr: number }) => dpr;\n');
  expectDiagnostic(layout, 'LAYOUT_DISPLAY_CAPABILITY', 'TSX dpr', '--final');
});

test('paint CommonJS edges cannot bypass the dependency graph', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-paint-require-');
  write(root, 'packages/docx/src/paint/require-edge.ts',
    "const helper = require('./require-helper.js');\nexport { helper };\n");
  write(root, 'packages/docx/src/paint/require-helper.ts',
    "const measured = require('../paragraph-measure.js');\nexport { measured };\n");
  write(root, 'packages/docx/src/paragraph-measure.ts', 'export const measured = true;\n');
  expectDiagnostic(root, 'FORBIDDEN_PAINT_EDGE', 'CommonJS paint edge', '--final');
});

test('allows only the exact parser normalization gateway and erased contracts', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-parser-gateway-');
  write(root, 'packages/docx/src/parser-model.ts',
    'export function normalizeInternalDocumentModel(document) { return { document, mathOccurrences: [] }; }\n');
  write(root, 'packages/docx/src/layout/resources.ts',
    "import { normalizeInternalDocumentModel } from '../parser-model.js';\n"
      + 'export function documentMathOccurrences(doc) {\n'
      + '  return [...normalizeInternalDocumentModel(doc).mathOccurrences];\n'
      + '}\n');
  const allowed = runChecker(root, '--final');
  assert.equal(allowed.status, 0, allowed.output);

  write(root, 'packages/docx/src/layout/resources.ts',
    "import { normalizeInternalDocumentModel as normalize } from '../parser-model.js';\n"
      + 'export function documentMathOccurrences(doc) { return [...normalize(doc).mathOccurrences]; }\n');
  expectDiagnostic(root, 'LAYOUT_PARSER_MODEL_DEPENDENCY', 'aliased gateway', '--final');
});

test('layout rejects direct and transitive parser-model runtime dependencies', () => {
  for (const [name, setup] of [
    ['direct', (root) => write(root, 'packages/docx/src/layout/parser-edge.ts',
      "import { parserFact } from '../parser-model.js';\nexport const fact = parserFact;\n")],
    ['transitive', (root) => {
      write(root, 'packages/docx/src/layout/parser-edge.ts',
        "import { parserFact } from '../parser-bridge.js';\nexport const fact = parserFact;\n");
      write(root, 'packages/docx/src/parser-bridge.ts',
        "import { parserFact } from './parser-model.js';\nexport { parserFact };\n");
    }],
    ['literal dynamic', (root) => write(root, 'packages/docx/src/layout/parser-edge.ts',
      "export const load = () => import('../parser-model.js');\n")],
    ['CommonJS', (root) => write(root, 'packages/docx/src/layout/parser-edge.ts',
      "export const load = () => require('../parser-model.js');\n")],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-parser-${name}-`);
    write(root, 'packages/docx/src/parser-model.ts', 'export const parserFact = true;\n');
    setup(root);
    expectDiagnostic(root, 'LAYOUT_PARSER_MODEL_DEPENDENCY', name, '--final');
  }
});

test('parser normalization gateway rejects aliases, re-exports, leaks, and extra bindings', () => {
  const exact =
    "import { normalizeInternalDocumentModel } from '../parser-model.js';\n"
    + 'export function documentMathOccurrences(doc) {\n'
    + '  return [...normalizeInternalDocumentModel(doc).mathOccurrences];\n'
    + '}\n';
  const sources = [
    "import { normalizeInternalDocumentModel, parserFact } from '../parser-model.js';\n"
      + 'export const leaked = [normalizeInternalDocumentModel, parserFact];\n',
    "import * as parserModel from '../parser-model.js';\nexport const leaked = parserModel;\n",
    "export { normalizeInternalDocumentModel } from '../parser-model.js';\n",
    "export * from '../parser-model.js';\n",
    `${exact}export { normalizeInternalDocumentModel };\n`,
    `${exact}export const leaked = normalizeInternalDocumentModel;\n`,
    "import { normalizeInternalDocumentModel } from '../parser-model.js';\n"
      + 'const normalize = normalizeInternalDocumentModel;\n'
      + 'export function documentMathOccurrences(doc) { return [...normalize(doc).mathOccurrences]; }\n',
  ];
  for (const [index, source] of sources.entries()) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-parser-gateway-${index}-`);
    write(root, 'packages/docx/src/parser-model.ts',
      'export const parserFact = true;\n'
        + 'export function normalizeInternalDocumentModel(document) { return { document, mathOccurrences: [] }; }\n');
    write(root, 'packages/docx/src/layout/resources.ts', source);
    expectDiagnostic(root, 'LAYOUT_PARSER_MODEL_DEPENDENCY', `gateway case ${index}`, '--final');
  }
});

test('layout may share parser-independent values and erased parser-adjacent contracts', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-parser-erased-');
  write(root, 'packages/docx/src/parser-model.ts', 'export const parserFact = true;\n');
  write(root, 'packages/docx/src/parser-contract.ts',
    "import { parserFact } from './parser-model.js';\nexport interface ParserContract { value: typeof parserFact }\n");
  write(root, 'packages/docx/src/shared-primitive.ts', 'export const nextTabStop = 36;\n');
  write(root, 'packages/docx/src/layout/parser-contract-user.ts',
    "import type { ParserContract } from '../parser-contract.js';\n"
      + "import { nextTabStop } from '../shared-primitive.js';\n"
      + 'export type Contract = ParserContract;\nexport const stop = nextTabStop;\n');
  const result = runChecker(root, '--final');
  assert.equal(result.status, 0, result.output);
});

test('body-layout input adapter has one exact parser acquisition projection', () => {
  const root = initializeBodyLayoutAdapterFixture();
  assert.equal(runChecker(root, '--final').status, 0);
});

test('body-layout input adapter rejects non-exact import sets', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-import-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    "import type { DocxDocumentModel } from './types.js';\n",
    '',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_IMPORT', 'incomplete import set', '--final');
});

test('body-layout input adapter rejects incomplete bindings within a reviewed module', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-incomplete-binding-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    'import { projectBodyLayoutInput, type BodyLayoutInput }',
    'import { projectBodyLayoutInput }',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_IMPORT', 'incomplete reviewed binding', '--final');
});

test('body-layout input adapter rejects duplicate, default, and unreviewed imports', () => {
  for (const [name, source] of [
    ['duplicate', canonicalBodyLayoutAdapter.replace(
      "import type { DocxDocumentModel } from './types.js';\n",
      "import type { DocxDocumentModel } from './types.js';\n"
        + "import type { DocxDocumentModel as DuplicateModel } from './types.js';\n",
    )],
    ['default', canonicalBodyLayoutAdapter.replace(
      "import type { DocxDocumentModel } from './types.js';",
      "import DocxDocumentModel from './types.js';",
    )],
    ['unreviewed', `${canonicalBodyLayoutAdapter}import { extra } from './extra.js';\nvoid extra;\n`],
  ]) {
    const root = initializeBodyLayoutAdapterFixture(`docx-layout-boundary-input-import-${name}-`);
    write(root, 'packages/docx/src/extra.ts', 'export const extra = true;\n');
    write(root, 'packages/docx/src/body-layout-input.ts', source);
    expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_IMPORT', name, '--final');
  }
});

test('body-layout input adapter rejects a non-literal import specifier', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-dynamic-specifier-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    "import { bodyLayoutAcquisitionInput } from './parser-model.js';",
    'import { bodyLayoutAcquisitionInput } from parserModule;',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_IMPORT', 'non-literal import', '--final');
});

test('body-layout input adapter rejects aliased or incorrectly typed bindings', () => {
  for (const [name, source] of [
    ['alias', canonicalBodyLayoutAdapter.replace(
      '{ bodyLayoutAcquisitionInput }',
      '{ bodyLayoutAcquisitionInput as acquireBodyLayoutInput }',
    )],
    ['value-as-type', canonicalBodyLayoutAdapter.replace(
      'import type { DocxDocumentModel }',
      'import { DocxDocumentModel }',
    )],
  ]) {
    const root = initializeBodyLayoutAdapterFixture(`docx-layout-boundary-input-binding-${name}-`);
    write(root, 'packages/docx/src/body-layout-input.ts', source);
    expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_BINDING', name, '--final');
  }
});

test('body-layout input adapter rejects additional declarations', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-declaration-');
  write(root, 'packages/docx/src/body-layout-input.ts',
    `${canonicalBodyLayoutAdapter}\nfunction hiddenProjection() {}\n`);
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_DECLARATION', 'additional declaration', '--final');
});

test('body-layout input adapter requires exactly one declaration', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-missing-declaration-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    /export function createBodyLayoutInput[\s\S]*?\n}\n/,
    '',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_DECLARATION', 'missing declaration', '--final');
});

test('body-layout input adapter rejects alternate export forms', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-export-');
  write(root, 'packages/docx/src/body-layout-input.ts',
    `${canonicalBodyLayoutAdapter}\nexport { createBodyLayoutInput as default };\n`);
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_EXPORT', 'export declaration', '--final');
});

test('body-layout input adapter rejects a default declaration export', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-default-export-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    'export function createBodyLayoutInput',
    'export default function createBodyLayoutInput',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_EXPORT', 'default declaration', '--final');
});

test('body-layout input adapter rejects any body other than the exact projection', () => {
  const root = initializeBodyLayoutAdapterFixture('docx-layout-boundary-input-body-');
  write(root, 'packages/docx/src/body-layout-input.ts', canonicalBodyLayoutAdapter.replace(
    'return projectBodyLayoutInput(bodyLayoutAcquisitionInput(document));',
    'return projectBodyLayoutInput(document);',
  ));
  expectDiagnostic(root, 'BODY_LAYOUT_ADAPTER_BODY', 'non-acquisition body', '--final');
});

test('retained body paint cannot reacquire layout while legacy story layout stays scoped outside it', () => {
  const root = initializeLegacyStoryPaintFixture();
  assert.equal(runChecker(root, '--base-ref', 'main').status, 0);

  write(root, 'packages/docx/src/renderer.ts', `
import { paintPlacedParagraphLayout } from './paint/canvas-text.js';
import { layoutLines } from './line-layout.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
function createConcreteBodyLayoutKernel(doc, ctx, localMetrics) { return { doc, ctx, localMetrics }; }
function createLayoutServices(doc, ctx, localMetrics) {
  const services = {};
  attachBodyLayoutKernel(services, createConcreteBodyLayoutKernel(doc, ctx, localMetrics));
  return services;
}
function renderBodyElements(elements) {
  for (const element of elements) {
    if (element.type === 'paragraph') layoutLines(element);
  }
}
function renderDocumentToCanvas(services, input, pageIndex) {
  return selectDocumentLayoutPage(services, input, pageIndex);
}
function renderHeaderFooterStory(story) { return layoutLines(story); }
`);
  expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', 'body reacquisition', '--base-ref', 'main');
});

test('retained body paint rejects aliased, transitive, callback, and unresolved layout calls', () => {
  const renderers = [
    ['alias', `const paintRetained = measureParagraph;
function renderBodyElements(elements) {
  for (const element of elements) if (element.type === 'paragraph') paintRetained(element);
}`],
    ['transitive', `function paintRetained() { acquireParagraphLayout(); }
function renderBodyElements(elements) {
  for (const element of elements) if (element.type === 'paragraph') paintRetained(element);
}`],
    ['callback', `function renderBodyElements(elements) {
  for (const element of elements) if (element.type === 'paragraph') paintPlacedParagraphLayout({
    deferFrontDrawing: () => measureParagraph(element),
  });
}`],
    ['unresolved', `function renderBodyElements(elements) {
  for (const element of elements) if (element.type === 'paragraph') opaquePaint(element);
}`],
  ];
  for (const [name, renderer] of renderers) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-body-${name}-`);
    write(root, 'packages/docx/src/renderer.ts', renderer);
    expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', name, '--final');
  }
});

test('every retained-body layout acquisition capability is rejected directly', () => {
  const calls = [
    'acquireParagraphLayout()',
    'acquireRetainedFrameGroup()',
    'buildSegments()',
    'contextualSpacingAdjust()',
    'estimateParagraphHeight()',
    'layoutLines()',
    'measureParagraph()',
    'measureText()',
    'paragraphGapAdjustment()',
    'paragraphLayoutFromMeasurement()',
    'parasShareBorderBox()',
    'renderFrameParagraph()',
    'renderParagraph()',
    'resolveParagraphBorderEdges()',
    'resolveFrameBox()',
  ];
  for (const call of calls) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-body-capability-${call}-`);
    write(root, 'packages/docx/src/renderer.ts',
      `function renderBodyElements(elements) {
        for (const element of elements) if (element.type === 'paragraph') ${call};
      }
      `);
    expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', call, '--final');
  }
});

test('late mutation of retained body sidecars is rejected before paint dispatch', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-body-sidecar-mutation-');
  write(root, 'packages/docx/src/renderer.ts',
    `const bodyFlowFragments = Object.freeze(Object.assign(new WeakMap(), {
      sourceIndices: new WeakMap(),
      framePlacement: new WeakMap(),
    }));
    bodyFlowFragments.sourceIndices.retainedTableMeasureBySource = new WeakMap();
    function renderBodyElements(elements) {
      for (const element of elements) {
        if (element.type === 'paragraph') state.onTextRun(element);
      }
    }
    `);
  expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', 'late sidecar mutation', '--final');
});

test('retained body paint fails closed without an auditable paragraph branch', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-body-fail-closed-');
  write(root, 'packages/docx/src/renderer.ts',
    'function renderBodyElements(elements) { return dispatchBody(elements); }\n');
  const result = expectDiagnostic(
    root,
    'BODY_PAINT_LAYOUT_CAPABILITY',
    'opaque body dispatch',
    '--final',
  );
  assert.match(result.output, /no statically auditable paragraph branch/);
});

test('retained body paint follows the else branch of a negated paragraph dispatch', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-body-negated-');
  write(root, 'packages/docx/src/renderer.ts',
    `function renderBodyElements(elements) {
      for (const element of elements) {
        if (element.type !== 'paragraph') paintTable(element);
        else measureParagraph(element);
      }
    }
    function paintTable() {}
    `);
  expectDiagnostic(root, 'BODY_PAINT_LAYOUT_CAPABILITY', 'negated paragraph dispatch', '--final');
});

test('final renderer rejects transitional exports, hidden algorithms, and non-layout imports', () => {
  for (const [name, source, diagnostic] of [
    ['deleted pagination export', canonicalRenderer + '\nexport function paginateDocument() {}\n', 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER'],
    ['star export', canonicalRenderer + "\nexport * from './layout/document.js';\n", 'FINAL_ADAPTER_EXPORT'],
    ['hidden algorithm', canonicalRenderer + '\nexport function accidentalAlgorithm() {}\n', 'FINAL_ADAPTER_DECLARATION'],
    ['inline layout loop', canonicalRenderer.replace(
      'const selection = selectDocumentLayoutPage(services, input, pageIndex);',
      'const selection = selectDocumentLayoutPage(services, input, pageIndex);\n'
        + '  for (const item of input) paintLayoutPage(item);',
    ), 'FINAL_ADAPTER_BODY'],
    ['non-layout import', "import { hidden } from './hidden.js';\n" + canonicalRenderer, 'FINAL_ADAPTER_IMPORT'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-adapter-${name}-`);
    write(root, 'packages/docx/src/hidden.ts', 'export function hidden() {}\n');
    write(root, 'packages/docx/src/renderer.ts', source);
    expectDiagnostic(root, diagnostic, name, '--final');
  }
});

test('final renderer import boundary rejects dynamic, bare, and unresolved imports exactly', () => {
  for (const [name, prefix] of [
    ['dynamic', 'import(globalThis.moduleName);\n'],
    ['bare', "import '@silurus/ooxml-core';\n"],
    ['unresolved', "import { missing } from './layout/missing.js';\nvoid missing;\n"],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-final-import-${name}-`);
    write(root, 'packages/docx/src/renderer.ts', `${prefix}${canonicalRenderer}`);
    expectDiagnostic(root, 'FINAL_ADAPTER_IMPORT', name, '--final');
  }
});

test('test-only adapters stay excluded from production imports', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-test-support-');
  write(root, 'packages/docx/src/canonical-layout-test-adapter.test.ts',
    'export function acquireForTest() {}\n');
  assert.equal(runChecker(root, '--final').status, 0);

  write(root, 'packages/docx/src/document.ts',
    "import { acquireForTest } from './canonical-layout-test-adapter.test.js';\n"
      + 'export const value = acquireForTest();\n');
  expectDiagnostic(root, 'PRODUCTION_TEST_SUPPORT_IMPORT', 'production test adapter import', '--final');
});

test('canonical producer must validate and deeply freeze its retained document layout', () => {
  for (const [name, source, diagnostic] of [
    ['second producer',
      'export function paginateBody() { return {}; }\nexport function alternateBodyProducer() { return {}; }\n',
      'CANONICAL_LAYOUT_PRODUCER'],
    ['missing validation',
      'export function paginateBody() { const layout = { pages: [] }; return deepFreezeDocumentLayout(layout); }\n',
      'RETAINED_LAYOUT_IMMUTABILITY'],
    ['mutable return',
      'export function paginateBody() { const layout = { pages: [] }; assertDocumentLayout(layout); return layout; }\n',
      'RETAINED_LAYOUT_IMMUTABILITY'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-producer-${name}-`);
    write(root, 'packages/docx/src/layout/body-paginator.ts', source);
    expectDiagnostic(root, diagnostic, name, '--final');
  }
});

test('selected variant owns normalization and page validation', () => {
  for (const [name, source] of [
    ['silent missing store',
      'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
        + '  const store = layoutVariantStoreOf(services);\n'
        + '  return store?.selectPage(layoutOptionsForRender(input), pageIndex);\n}\n'],
    ['raw layout selection',
      'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
        + '  const store = layoutVariantStoreOf(services);\n'
        + "  if (!store) throw new Error('missing');\n"
        + '  return store.layoutFor(layoutOptionsForRender(input)).pages[pageIndex];\n}\n'],
    ['unnormalized selection',
      'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
        + '  const store = layoutVariantStoreOf(services);\n'
        + "  if (!store) throw new Error('missing');\n"
        + '  return store.selectPage(input, pageIndex);\n}\n'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-selection-${name}-`);
    write(root, 'packages/docx/src/layout/document-layout-variants.ts', source);
    expectDiagnostic(root, 'SELECTED_LAYOUT_VARIANT', name, '--final');
  }
});

test('worker retains keyed parity and never duplicates selected-page validation', () => {
  for (const [name, source] of [
    ['duplicate worker selection',
      "import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';\n"
        + 'export function render() { return selectDocumentLayoutPage(services, input, pageIndex); }\n'],
    ['raw worker pages', 'export const pages = [];\n'],
    ['unkeyed builder',
      "import { renderDocumentToCanvas } from './renderer.js';\n"
        + 'export function render() { return renderDocumentToCanvas(); }\n'],
  ]) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-worker-${name}-`);
    write(root, 'packages/docx/src/render-worker.ts', source);
    expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', name, '--final');
  }
});

test('a transitional baseline is forbidden in explicit final mode', () => {
  const root = initializeLegacyStoryPaintFixture();
  expectDiagnostic(root, 'FINAL_BASELINE_PRESENT', 'final baseline', '--final');
});

test('reports an unknown CLI argument exactly', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-cli-');
  expectDiagnostic(root, 'UNKNOWN_ARGUMENT', 'unknown argument', '--definitely-unknown');
});

test('reports a missing final renderer exactly', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-missing-renderer-');
  rmSync(join(root, 'packages/docx/src/renderer.ts'));
  expectDiagnostic(root, 'FINAL_ADAPTER_MISSING', 'missing renderer', '--final');
});

test('reports final legacy inventory without accepting an alternate diagnostic', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-final-legacy-');
  write(root, 'packages/docx/src/layout/legacy-route.ts',
    'export const useOldLayoutEngine = true;\n');
  expectDiagnostic(root, 'FINAL_LEGACY_BOUNDARY', 'legacy migration flag', '--final');
});

test('reports non-literal module edges from paint exactly', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-nonliteral-paint-');
  write(root, 'packages/docx/src/paint/dynamic.ts',
    "const moduleName = './canvas-page.js';\nexport const load = () => import(moduleName);\n");
  expectDiagnostic(root, 'NON_LITERAL_MODULE_EDGE', 'paint dynamic import', '--final');
});

test('reports non-literal module edges from layout exactly', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-nonliteral-layout-');
  write(root, 'packages/docx/src/layout/dynamic.ts',
    "const moduleName = './types.js';\nexport const load = () => import(moduleName);\n");
  expectDiagnostic(root, 'NON_LITERAL_LAYOUT_MODULE_EDGE', 'layout dynamic import', '--final');
});

test('production rejects both test and test-support module imports', () => {
  for (const suffix of ['test', 'test-support']) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-production-${suffix}-`);
    write(root, `packages/docx/src/fixture.${suffix}.ts`, 'export const fixtureValue = 1;\n');
    write(root, 'packages/docx/src/document.ts',
      `import { fixtureValue } from './fixture.${suffix}.js';\nexport const value = fixtureValue;\n`);
    expectDiagnostic(root, 'PRODUCTION_TEST_SUPPORT_IMPORT', suffix, '--final');
  }
});

test('every deleted final producer and legacy stamp identifier is rejected exactly', () => {
  const deletedProducers = [
    'bodyFragmentFor',
    'bodyLayoutFallback',
    'computePages',
    'paginateDocument',
    'paginateWithHeaderFooterReserve',
    'PaginatedBodyElement',
    'physicalPageSizeForPage',
    'prebuiltPages',
    'retainedLayout',
    'sectionBreakSpacer',
    'collapsedSpacer',
    'leadsCollapsedRun',
    'hiddenCollapsed',
  ];
  const deletedStampProperties = [
    'colIndex',
    'colGeom',
    'colTopPt',
    'sectionHF',
    'sectionGeom',
    'sectionPageNumType',
    'sectionTextDirection',
  ];
  for (const name of deletedProducers) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-deleted-exact-${name}-`);
    write(root, 'packages/docx/src/deleted-path.ts', `export const ${name} = true;\n`);
    expectDiagnostic(root, 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER', name, '--final');
  }
  for (const name of deletedStampProperties) {
    const root = initializeCanonicalFixture(`docx-layout-boundary-deleted-property-${name}-`);
    write(root, 'packages/docx/src/deleted-path.ts',
      `export const legacyStamp = { ${name}: true };\n`);
    expectDiagnostic(root, 'FORBIDDEN_PAGE_PRODUCER_IDENTIFIER', name, '--final');
  }
});

test('legacy stamp property guards do not ban unrelated local variable names', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-local-column-index-');
  write(root, 'packages/docx/src/layout/local-index.ts',
    'export function localIndex() { const colIndex = 0; return colIndex + 1; }\n');
  const result = runChecker(root, '--final');
  assert.equal(result.status, 0, result.output);
});

test('the concrete body kernel is private even though its declaration is allowed', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-exported-kernel-');
  write(root, 'packages/docx/src/renderer.ts', canonicalRenderer.replace(
    'function createConcreteBodyLayoutKernel',
    'export function createConcreteBodyLayoutKernel',
  ));
  expectDiagnostic(root, 'FINAL_ADAPTER_EXPORT', 'exported concrete kernel', '--final');
});

test('writing a first transitional baseline remains supported', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-first-baseline-');
  git(root, 'init', '-b', 'main');
  git(root, 'config', 'user.email', 'boundary-test@example.invalid');
  git(root, 'config', 'user.name', 'Boundary Test');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'base without boundary baseline');
  git(root, 'switch', '-c', 'boundary');
  const result = runChecker(root, '--write-transitional-baseline', '--base-ref', 'main');
  assert.equal(result.status, 0, result.output);
  assert.equal(JSON.parse(readFileSync(
    join(root, 'scripts/docx-layout-boundary-baseline.json'),
    'utf8',
  )).version, 2);
});

test('reports an existing merge-base transitional baseline exactly', () => {
  const root = initializeTransitionalRepository();
  expectDiagnostic(
    root,
    'TRANSITIONAL_BASELINE_EXISTS',
    'baseline rewrite',
    '--write-transitional-baseline',
    '--base-ref',
    'main',
  );
});

test('reports invalid current baseline structure exactly', () => {
  const root = initializeTransitionalRepository();
  write(root, 'scripts/docx-layout-boundary-baseline.json', '{"version":1}\n');
  expectDiagnostic(root, 'INVALID_BASELINE', 'invalid current baseline', '--base-ref', 'main');
});

test('reports malformed current baseline JSON as INVALID_BASELINE', () => {
  const root = initializeTransitionalRepository();
  write(root, 'scripts/docx-layout-boundary-baseline.json', '{not-json\n');
  expectDiagnostic(root, 'INVALID_BASELINE', 'malformed current baseline', '--base-ref', 'main');
});

test('reports invalid merge-base baseline structure exactly', () => {
  const root = initializeTransitionalRepository();
  git(root, 'switch', 'main');
  write(root, 'scripts/docx-layout-boundary-baseline.json', '{"version":1}\n');
  git(root, 'add', '.');
  git(root, 'commit', '-m', 'invalid baseline');
  git(root, 'switch', '-c', 'invalid-base-child');
  expectDiagnostic(root, 'INVALID_BASELINE', 'invalid merge-base baseline', '--base-ref', 'main');
});

test('reports Git command failures exactly', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-git-error-');
  write(root, 'scripts/docx-layout-boundary-baseline.json',
    `${JSON.stringify(validEmptyBaseline(), null, 2)}\n`);
  expectDiagnostic(root, 'GIT_ERROR', 'not a Git repository', '--base-ref', 'main');
});

test('baseline expansion guards each tracked allowance category', () => {
  const mutations = [
    ['legacy symbol count', (root) => write(
      root,
      'packages/docx/src/legacy-symbol.ts',
      'export function renderShapeText() {}\n',
    )],
    ['migration identifier count', (root) => write(
      root,
      'packages/docx/src/migration-flag.ts',
      'export const useOldLayoutEngine = true;\n',
    )],
    ['non-layout declaration', (root) => write(
      root,
      'packages/docx/src/new-helper.ts',
      'export const newHelper = true;\n',
    )],
    ['unapproved legacy declaration hash', (root) => {
      const path = join(root, 'scripts/docx-layout-boundary-baseline.json');
      const baseline = JSON.parse(readFileSync(path, 'utf8'));
      baseline.legacyDeclarationHashes['packages/docx/src/ghost.ts#FunctionDeclaration#renderShapeText'] = 'ghost';
      write(root, 'scripts/docx-layout-boundary-baseline.json',
        `${JSON.stringify(baseline, null, 2)}\n`);
    }],
    ['renderer import edge', (root) => {
      write(root, 'packages/docx/src/layout-context.ts', 'export const legacyContext = true;\n');
      const rendererPath = join(root, 'packages/docx/src/renderer.ts');
      write(root, 'packages/docx/src/renderer.ts',
        `import { legacyContext } from './layout-context.js';\n${readFileSync(rendererPath, 'utf8')}\nvoid legacyContext;\n`);
    }],
  ];
  for (const [name, mutate] of mutations) {
    const root = initializeTransitionalRepository();
    mutate(root);
    expectDiagnostic(root, 'BASELINE_EXPANSION', name, '--base-ref', 'main');
  }
});

test('reports a changed frozen legacy declaration exactly', () => {
  const root = initializeTransitionalRepository();
  const rendererPath = join(root, 'packages/docx/src/renderer.ts');
  write(root, 'packages/docx/src/renderer.ts', readFileSync(rendererPath, 'utf8').replace(
    'function renderShapeText() { return 1; }',
    'function renderShapeText() { return 2; }',
  ));
  expectDiagnostic(root, 'LEGACY_DECLARATION_CHANGED', 'changed declaration', '--base-ref', 'main');
});

test('reports an exact-baseline mismatch after permitted shrinkage', () => {
  const root = initializeTransitionalRepository();
  const baselinePath = join(root, 'scripts/docx-layout-boundary-baseline.json');
  const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
  baseline.nonLayoutDeclarationKeys = [];
  write(root, 'scripts/docx-layout-boundary-baseline.json',
    `${JSON.stringify(baseline, null, 2)}\n`);
  expectDiagnostic(root, 'BASELINE_MISMATCH', 'baseline shrink mismatch', '--base-ref', 'main');
});

test('the canonical body producer file is mandatory', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-producer-missing-');
  rmSync(join(root, 'packages/docx/src/layout/body-paginator.ts'));
  expectDiagnostic(root, 'CANONICAL_LAYOUT_PRODUCER', 'missing producer', '--final');
});

test('the canonical body producer rejects alternate exported value producers', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-producer-value-');
  const path = join(root, 'packages/docx/src/layout/body-paginator.ts');
  write(root, 'packages/docx/src/layout/body-paginator.ts',
    `${readFileSync(path, 'utf8')}\nexport const alternateBodyProducer = () => ({ pages: [] });\n`);
  expectDiagnostic(root, 'CANONICAL_LAYOUT_PRODUCER', 'alternate value producer', '--final');
});

test('the canonical body producer rejects runtime re-exports', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-producer-reexport-');
  const path = join(root, 'packages/docx/src/layout/body-paginator.ts');
  write(root, 'packages/docx/src/layout/alternate-producer.ts',
    'export const alternateBodyProducer = () => ({ pages: [] });\n');
  write(root, 'packages/docx/src/layout/body-paginator.ts',
    `${readFileSync(path, 'utf8')}\nexport { alternateBodyProducer } from './alternate-producer.js';\n`);
  expectDiagnostic(root, 'CANONICAL_LAYOUT_PRODUCER', 'runtime re-export', '--final');
});

test('retained layout validation and freezing must target the same returned value', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-immutability-identity-');
  write(root, 'packages/docx/src/layout/body-paginator.ts',
    'export function paginateBody() {\n'
      + '  const validated = { pages: [], diagnostics: [] };\n'
      + '  const returned = { pages: [], diagnostics: [] };\n'
      + '  assertDocumentLayout(validated);\n'
      + '  return deepFreezeDocumentLayout(returned);\n'
      + '}\n');
  expectDiagnostic(root, 'RETAINED_LAYOUT_IMMUTABILITY', 'different validated value', '--final');
});

test('retained layout validation uses the reviewed invariants import', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-immutability-impostor-');
  write(root, 'packages/docx/src/layout/body-paginator.ts',
    'function assertDocumentLayout(value) { return value; }\n'
      + 'function deepFreezeDocumentLayout(value) { return value; }\n'
      + 'export function paginateBody() {\n'
      + '  const layout = { pages: [], diagnostics: [] };\n'
      + '  assertDocumentLayout(layout);\n'
      + '  return deepFreezeDocumentLayout(layout);\n'
      + '}\n');
  expectDiagnostic(root, 'RETAINED_LAYOUT_IMMUTABILITY', 'local impostor invariants', '--final');
});

test('the selected-layout variant module is mandatory', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-selection-missing-');
  rmSync(join(root, 'packages/docx/src/layout/document-layout-variants.ts'));
  expectDiagnostic(root, 'SELECTED_LAYOUT_VARIANT', 'missing selection module', '--final');
});

test('selected-layout validation rejects extra raw page selection even with the canonical statements present', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-selection-extra-');
  write(root, 'packages/docx/src/layout/document-layout-variants.ts',
    'export function selectDocumentLayoutPage(services, input, pageIndex) {\n'
      + '  if (input.rawLayout) return input.rawLayout.pages[pageIndex];\n'
      + '  const store = layoutVariantStoreOf(services);\n'
      + "  if (!store) throw new Error('missing');\n"
      + '  return store.selectPage(layoutOptionsForRender(input), pageIndex);\n'
      + '}\n');
  expectDiagnostic(root, 'SELECTED_LAYOUT_VARIANT', 'extra raw path', '--final');
});

test('renderer must delegate page validation to the selected-layout boundary exactly once', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-renderer-selection-');
  write(root, 'packages/docx/src/renderer.ts', canonicalRenderer.replace(
    'const selection = selectDocumentLayoutPage(services, input, pageIndex);',
    'const selection = { page: input.pages[pageIndex] };',
  ));
  expectDiagnostic(root, 'SELECTED_LAYOUT_VARIANT', 'renderer raw page selection', '--final');
});

test('the worker canonical selection route is mandatory', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-worker-missing-');
  rmSync(join(root, 'packages/docx/src/render-worker.ts'));
  expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', 'missing worker', '--final');
});

test('worker selection validates call ownership rather than only call counts', () => {
  const root = initializeCanonicalFixture('docx-layout-boundary-worker-call-shape-');
  write(root, 'packages/docx/src/render-worker.ts',
    'export function initializeWorker(model, layoutServices) {\n'
      + '  layoutDocument();\n'
      + '  return attachDocumentLayoutVariants({ buildLayout: () => ({ pages: [] }) });\n'
      + '}\n'
      + 'export function renderWorkerPage() { return renderDocumentToCanvas(); }\n'
      + 'export function collectWorkerRuns() { return renderDocumentToCanvas(); }\n');
  expectDiagnostic(root, 'WORKER_LAYOUT_SELECTION', 'wrong worker call ownership', '--final');
});
