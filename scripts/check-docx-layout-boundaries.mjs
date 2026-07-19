#!/usr/bin/env node

import { existsSync, mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { basename, dirname, join, relative, resolve, sep } from 'node:path';
import { spawnSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';

const require = createRequire(new URL('../packages/docx/package.json', import.meta.url));
const ts = require('typescript');

const BASELINE_PATH = 'scripts/docx-layout-boundary-baseline.json';
const DOCX_SOURCE = 'packages/docx/src';
const PAINT_SOURCE = `${DOCX_SOURCE}/paint`;
const LAYOUT_SOURCE = `${DOCX_SOURCE}/layout`;
const PARSER_MODEL = `${DOCX_SOURCE}/parser-model.ts`;
const BODY_LAYOUT_ADAPTER = `${DOCX_SOURCE}/body-layout-input.ts`;
const PARAGRAPH_ANCHOR_FRAME_ADAPTER = `${DOCX_SOURCE}/paragraph-anchor-frame-adapter.ts`;
const WORKER_LAYOUT_RETENTION = `${DOCX_SOURCE}/render-worker-layout.ts`;
const TEXT_RUN_PROJECTION_ADAPTER = `${DOCX_SOURCE}/text-run-projection.ts`;
const LAYOUT_PARSER_MODEL_GATEWAY = `${LAYOUT_SOURCE}/resources.ts`;
const LAYOUT_AFFINE = `${LAYOUT_SOURCE}/affine.ts`;
const LAYOUT_PARSER_MODEL_GATEWAY_IMPORT = '../parser-model.js';
const LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL = 'normalizeInternalDocumentModel';

const FINAL_RENDERER_EXPORTS = new Set([
  'DocxTextRunInfo',
  'RenderDocumentOptions',
  'clearResolvedLocalFonts',
  'documentHasMath',
  'dropColorReplacedCache',
  'prepareMathRuns',
  'renderDocumentToCanvas',
  'setResolvedLocalFonts',
]);

const FINAL_RENDERER_DECLARATIONS = new Set([
  ...FINAL_RENDERER_EXPORTS,
  'createConcreteBodyLayoutKernel',
  'createLayoutServices',
  'normalizeRenderOptions',
]);

const A5_STATE_OWNER_DECLARATIONS = new Set([
  'RetainedTableRecord',
  'commitFloatRegistryDelta',
  'prepareFittingOuterFragment',
  'reacquirePageBlock',
  'retainedTableRecord',
]);

const BODY_LAYOUT_ADAPTER_DECLARATIONS = new Set(['createBodyLayoutInput']);
const PARAGRAPH_ANCHOR_FRAME_ADAPTER_DECLARATIONS = new Set([
  'ParagraphAnchorReferenceFrameSnapshot',
  'paragraphAnchorReferenceFrames',
]);
const WORKER_LAYOUT_RETENTION_DECLARATIONS = new Set([
  'RetainedRenderWorkerDocumentLayout',
  'retainRenderWorkerDocumentLayout',
]);
const TEXT_RUN_PROJECTION_ADAPTER_DECLARATIONS = new Set([
  'SelectedTextRunsForPageOptions',
  'TextRunsForPageOptions',
  'projectTextRun',
  'textRunsForPage',
  'textRunsForSelectedPage',
]);
const LAYOUT_AFFINE_EXPORTS = new Set([
  'composeAffine',
  'inverseMapAffinePoint',
  'inverseMapAffineVector',
  'mapAffinePoint',
  'quarterTurnAffine',
  'scaleAffine',
  'translationAffine',
]);
const BODY_KERNEL_IMPLEMENTATION_DECLARATIONS = new Set(['createConcreteBodyLayoutKernel']);
const BODY_LAYOUT_ADAPTER_IMPORT_BINDINGS = new Map([
  [PARSER_MODEL, new Map([['bodyLayoutAcquisitionInput', 'value']])],
  [`${DOCX_SOURCE}/types.ts`, new Map([['DocxDocumentModel', 'type']])],
  [`${LAYOUT_SOURCE}/body-layout-input.ts`, new Map([
    ['projectBodyLayoutInput', 'value'],
    ['BodyLayoutInput', 'type'],
  ])],
]);

const PLANNED_NON_LAYOUT_MODULES = new Set([
  `${DOCX_SOURCE}/parser-model.ts`,
]);

const SHARED_PAINT_IMPORTS = new Map([
  ['@silurus/ooxml-core', new Map([
    ['autoContrastColor', 'value'],
    ['canvasFontString', 'value'],
    ['crispOffset', 'value'],
    ['drawImageCropped', 'value'],
    ['doubleRailGeometry', 'value'],
    ['fillDoubleBorder', 'value'],
    ['HyperlinkTarget', 'type'],
    ['paintDrawingMLShape', 'value'],
    // Shared fill resolution keeps gradient/no-fill semantics identical across
    // DOCX, PPTX, and XLSX painters; paint may consume it but not layout APIs.
    ['resolveFill', 'value'],
    ['renderChart', 'value'],
  ])],
]);

const LEGACY_SYMBOLS = [
  'fitMeasureReuseEnabled',
  'fragmentPaintEnabled',
  'lineReuseEnabled',
  'isFragmentPaintableParagraph',
  'layoutLinesInputs',
  'stampParagraphLines',
  'renderBodyParagraphLines',
  'renderShapeText',
  'tableRequiresLegacyPaint',
  'isFragmentPaintableTable',
  'tableReuseEnabled',
  'renderTableFragment',
  'computePages',
  'computeTableLayout',
  'calculateRowHeight',
  'measureCellContentHeightPx',
  'buildTableCellBlocks',
  'renderHeaderFooter',
  'measureFootnoteHeight',
  'deferFront',
  'deferFrontDrawing',
  'deferBehindDrawing',
  'deferredFrontPaint',
  'deferredPaintWrapper',
  'bodyDrawingPass',
  'sectionBreakSpacer',
  'collapsedSpacer',
  'leadsCollapsedRun',
  'hiddenCollapsed',
  'tableColWidthsPt',
  'tableRowHeightsPt',
  'tableLayoutInputs',
];

const DELETED_PAGE_PRODUCER_IDENTIFIERS = new Set([
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
]);

const DELETED_LEGACY_STAMP_PROPERTIES = new Set([
  'colIndex',
  'colGeom',
  'colTopPt',
  'sectionHF',
  'sectionGeom',
  'sectionPageNumType',
  'sectionTextDirection',
]);

const LEGACY_RENDERER_IMPORTS = new Set([
  'layout-context.ts',
  'layout-fragments.ts',
  'line-layout.ts',
  'paragraph-measure.ts',
  'table-fragments.ts',
  'table-geometry.ts',
]);

function fail(code, detail) {
  throw new Error(`${code}: ${detail}`);
}

function posixPath(path) {
  return path.split(sep).join('/');
}

function isProductionTypeScript(path) {
  return /\.tsx?$/.test(path)
    && !path.endsWith('.d.ts')
    && !/\.(test|spec|stories|test-support)\.tsx?$/.test(path)
    && !path.includes('/wasm/');
}

function assertNoProductionTestSupportImports(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    for (const edge of moduleEdges(path)) {
      if (!edge.literal || !edge.specifier.startsWith('.')) continue;
      const dependency = resolveLocalImport(path, edge.specifier);
      if (dependency && /\.(?:test|test-support)\.tsx?$/.test(dependency)) {
        fail(
          'PRODUCTION_TEST_SUPPORT_IMPORT',
          `${posixPath(relative(root, path))} -> ${posixPath(relative(root, dependency))}`,
        );
      }
    }
  }
}

function listFiles(root) {
  if (!existsSync(root)) return [];
  const files = [];
  for (const entry of readdirSync(root)) {
    const path = join(root, entry);
    if (statSync(path).isDirectory()) files.push(...listFiles(path));
    else files.push(path);
  }
  return files;
}

function sourceFile(path) {
  const kind = path.endsWith('.tsx') ? ts.ScriptKind.TSX : ts.ScriptKind.TS;
  return ts.createSourceFile(path, readFileSync(path, 'utf8'), ts.ScriptTarget.Latest, true, kind);
}

function importIsTypeOnly(statement) {
  const clause = statement.importClause;
  if (!clause) return false;
  if (clause.isTypeOnly) return true;
  if (clause.name || !clause.namedBindings || !ts.isNamedImports(clause.namedBindings)) return false;
  return clause.namedBindings.elements.length > 0
    && clause.namedBindings.elements.every((element) => element.isTypeOnly);
}

function exportIsTypeOnly(statement) {
  if (statement.isTypeOnly) return true;
  return statement.exportClause
    && ts.isNamedExports(statement.exportClause)
    && statement.exportClause.elements.length > 0
    && statement.exportClause.elements.every((element) => element.isTypeOnly);
}

function moduleEdges(path) {
  const source = sourceFile(path);
  const edges = [];
  for (const statement of source.statements) {
    if (ts.isImportDeclaration(statement)
      && ts.isStringLiteral(statement.moduleSpecifier)) {
      const bindings = statement.importClause?.namedBindings;
      const importedNames = statement.importClause
        ? [
            ...(statement.importClause.name ? ['default'] : []),
            ...(bindings && ts.isNamespaceImport(bindings) ? ['*'] : []),
            ...(bindings && ts.isNamedImports(bindings)
              ? bindings.elements.map((element) => element.propertyName?.text ?? element.name.text)
              : []),
          ]
        : [];
      edges.push({
        kind: 'import',
        specifier: statement.moduleSpecifier.text,
        typeOnly: importIsTypeOnly(statement),
        literal: true,
        importedNames,
        aliased: !!(bindings && ts.isNamedImports(bindings)
          && bindings.elements.some((element) => element.propertyName && element.propertyName.text !== element.name.text)),
        bare: !statement.importClause,
      });
    }
    if (ts.isExportDeclaration(statement)
      && statement.moduleSpecifier
      && ts.isStringLiteral(statement.moduleSpecifier)) {
      edges.push({
        kind: 'export',
        specifier: statement.moduleSpecifier.text,
        typeOnly: exportIsTypeOnly(statement),
        literal: true,
        importedNames: statement.exportClause && ts.isNamedExports(statement.exportClause)
          ? statement.exportClause.elements.map((element) => element.propertyName?.text ?? element.name.text)
          : ['*'],
      });
    }
  }
  const visit = (node) => {
    if (ts.isCallExpression(node)
      && node.expression.kind === ts.SyntaxKind.ImportKeyword) {
      const argument = node.arguments[0];
      edges.push(argument && (ts.isStringLiteral(argument) || ts.isNoSubstitutionTemplateLiteral(argument))
        ? { kind: 'dynamic-import', specifier: argument.text, typeOnly: false, literal: true }
        : { kind: 'dynamic-import', specifier: '<dynamic>', typeOnly: false, literal: false });
    }
    if (ts.isCallExpression(node)
      && ts.isIdentifier(node.expression)
      && node.expression.text === 'require') {
      const argument = node.arguments[0];
      edges.push(argument && (ts.isStringLiteral(argument) || ts.isNoSubstitutionTemplateLiteral(argument))
        ? { kind: 'require', specifier: argument.text, typeOnly: false, literal: true }
        : { kind: 'require', specifier: '<dynamic>', typeOnly: false, literal: false });
    }
    ts.forEachChild(node, visit);
  };
  ts.forEachChild(source, visit);
  return edges;
}

function resolveLocalImport(importer, specifier) {
  if (!specifier.startsWith('.')) return null;
  const clean = specifier.split('?')[0].split('#')[0];
  const unresolved = resolve(dirname(importer), clean);
  const withoutJs = unresolved.replace(/\.(mjs|cjs|js)$/, '');
  const candidates = [
    unresolved,
    `${withoutJs}.ts`,
    `${withoutJs}.tsx`,
    join(unresolved, 'index.ts'),
  ];
  return candidates.find((candidate) => existsSync(candidate) && statSync(candidate).isFile()) ?? null;
}

function dependencyGraph(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  const files = listFiles(sourceRoot).filter(isProductionTypeScript);
  const graph = new Map();
  for (const file of files) {
    graph.set(file, moduleEdges(file));
  }
  return graph;
}

function paintBoundaryViolations(root) {
  const graph = dependencyGraph(root);
  const paintRoot = resolve(root, PAINT_SOURCE);
  const layoutTypes = resolve(root, LAYOUT_SOURCE, 'types.ts');
  const layoutAffine = resolve(root, LAYOUT_AFFINE);
  const paintAffine = resolve(root, PAINT_SOURCE, 'affine.ts');
  const entries = [...graph.keys()].filter((path) => path.startsWith(`${paintRoot}${sep}`));
  const violations = [];
  const nonLiteral = [];

  for (const entry of entries) {
    const stack = [{ path: entry, chain: [entry] }];
    const visited = new Set([entry]);
    while (stack.length > 0) {
      const current = stack.pop();
      for (const edge of graph.get(current.path) ?? []) {
        if (!edge.literal) {
          nonLiteral.push(posixPath(relative(root, current.path)));
          continue;
        }
        if (edge.bare) {
          violations.push([...current.chain.map((path) => posixPath(relative(root, path))), edge.specifier]);
          continue;
        }
        if (!edge.specifier.startsWith('.')) {
          const allowedNames = SHARED_PAINT_IMPORTS.get(edge.specifier);
          const allowed = edge.kind === 'import'
            && !edge.aliased
            && allowedNames
            && edge.importedNames?.length > 0
            && edge.importedNames?.every((name) => (
              allowedNames.get(name) === (edge.typeOnly ? 'type' : 'value')
            ));
          if (!allowed) {
            violations.push([...current.chain.map((path) => posixPath(relative(root, path))), edge.specifier]);
          }
          continue;
        }
        const dependency = resolveLocalImport(current.path, edge.specifier);
        if (!dependency) {
          violations.push([...current.chain.map((path) => posixPath(relative(root, path))), edge.specifier]);
          continue;
        }
        const chain = [...current.chain, dependency];
        const insidePaint = dependency.startsWith(`${paintRoot}${sep}`);
        const allowedAffineContract = current.path === paintAffine
          && dependency === layoutAffine
          && edge.kind === 'export'
          && !edge.typeOnly
          && edge.importedNames?.length === LAYOUT_AFFINE_EXPORTS.size
          && edge.importedNames.every((name) => LAYOUT_AFFINE_EXPORTS.has(name));
        const allowedContract = (edge.typeOnly && dependency === layoutTypes)
          || allowedAffineContract;
        if (!insidePaint && !allowedContract) {
          violations.push(chain.map((path) => posixPath(relative(root, path))));
          continue;
        }
        if (insidePaint && !visited.has(dependency)) {
          visited.add(dependency);
          stack.push({ path: dependency, chain });
        }
      }
    }
  }
  return { violations, nonLiteral };
}

function identifierText(node) {
  if (ts.isIdentifier(node) || ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) return node.text;
  return null;
}

function assertCapabilityBoundaries(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const rel = posixPath(relative(root, path));
    const inPaint = rel.startsWith(`${PAINT_SOURCE}/`);
    const inLayout = rel.startsWith(`${LAYOUT_SOURCE}/`);
    if (!inPaint && !inLayout) continue;
    const source = sourceFile(path);
    const visit = (node) => {
      const text = identifierText(node);
      if (inPaint && text === 'measureText') fail('PAINT_CAPABILITY', `${rel} uses measureText`);
      if (inPaint && text && /^(?:resolve|merge|combine|apply|fold|compose|inherit).*(?:Style|Properties|Pr|Cascade|Format|Formatting)$/i.test(text)) {
        fail('PAINT_CAPABILITY', `${rel} uses ${text}`);
      }
      if (inLayout && text && /^(?:resolve|merge|combine|apply|fold|compose|inherit).*(?:Style|Properties|Pr|Cascade|Format|Formatting)$/i.test(text)) {
        fail('LAYOUT_STYLE_CAPABILITY', `${rel} uses ${text}`);
      }
      if (inLayout && text && /^(?:dpr|displayScale|devicePixelRatio|CanvasRenderingContext2D|OffscreenCanvasRenderingContext2D)$/.test(text)) {
        fail('LAYOUT_DISPLAY_CAPABILITY', `${rel} uses ${text}`);
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
}

function assertOccurrenceProjectionRuntimeDependencies(root) {
  const occurrence = resolve(root, LAYOUT_SOURCE, 'occurrence-projection.ts');
  const translation = resolve(root, LAYOUT_SOURCE, 'retained-geometry-translation.ts');
  const plainData = resolve(root, LAYOUT_SOURCE, 'plain-data.ts');
  const guarded = [occurrence, translation, plainData];
  for (const path of guarded) {
    if (!existsSync(path)) {
      fail('OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', `missing ${posixPath(relative(root, path))}`);
    }
  }
  const allowedTargets = new Map([
    [occurrence, new Set([translation, plainData])],
    [translation, new Set()],
    [plainData, new Set()],
  ]);
  for (const current of guarded) {
    for (const edge of moduleEdges(current)) {
      if (edge.typeOnly) continue;
      const detail = `${posixPath(relative(root, current))} -> ${edge.specifier}`;
      if (!edge.literal || edge.kind === 'dynamic-import' || edge.kind === 'require'
        || edge.bare || edge.specifier.includes('?') || edge.specifier.includes('#')
        || !edge.specifier.startsWith('.')) {
        fail('OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', detail);
      }
      const dependency = resolveLocalImport(current, edge.specifier);
      if (!dependency || !allowedTargets.get(current)?.has(dependency)) {
        fail('OCCURRENCE_PROJECTION_RUNTIME_DEPENDENCY', detail);
      }
    }
  }
}

function assertCoordinateSpaceRuntimeDependencies(root) {
  const coordinate = resolve(root, LAYOUT_SOURCE, 'coordinate-space.ts');
  const pageFactory = resolve(root, LAYOUT_SOURCE, 'page-factory.ts');
  for (const path of [coordinate, pageFactory]) {
    if (!existsSync(path)) {
      fail('COORDINATE_SPACE_RUNTIME_DEPENDENCY', `missing ${posixPath(relative(root, path))}`);
    }
  }

  for (const edge of moduleEdges(coordinate)) {
    const detail = `${posixPath(relative(root, coordinate))} -> ${edge.specifier}`;
    if (edge.kind !== 'import' || !edge.literal || edge.bare || !edge.typeOnly
      || edge.specifier !== './types.js') {
      fail('COORDINATE_SPACE_RUNTIME_DEPENDENCY', detail);
    }
  }

  const runtimeTargets = new Set([
    resolve(root, LAYOUT_SOURCE, 'coordinate-space.ts'),
    // Page finalization owns retained section decoration geometry so paint
    // cannot reconstruct section bands from layout policy.
    resolve(root, LAYOUT_SOURCE, 'column-separators.ts'),
    resolve(root, LAYOUT_SOURCE, 'page-graph.ts'),
  ]);
  for (const edge of moduleEdges(pageFactory)) {
    if (edge.typeOnly) {
      if (edge.specifier.includes('?') || edge.specifier.includes('#')) {
        fail(
          'COORDINATE_SPACE_RUNTIME_DEPENDENCY',
          `${posixPath(relative(root, pageFactory))} -> ${edge.specifier}`,
        );
      }
      continue;
    }
    const detail = `${posixPath(relative(root, pageFactory))} -> ${edge.specifier}`;
    if (edge.kind !== 'import' || !edge.literal || edge.bare
      || edge.specifier.includes('?') || edge.specifier.includes('#')
      || !edge.specifier.startsWith('.')) {
      fail('COORDINATE_SPACE_RUNTIME_DEPENDENCY', detail);
    }
    const dependency = resolveLocalImport(pageFactory, edge.specifier);
    if (!dependency || !runtimeTargets.has(dependency)) {
      fail('COORDINATE_SPACE_RUNTIME_DEPENDENCY', detail);
    }
  }
}

function assertAffineRuntimeDependencies(root) {
  const affine = resolve(root, LAYOUT_AFFINE);
  if (!existsSync(affine)) {
    fail('AFFINE_RUNTIME_DEPENDENCY', `missing ${LAYOUT_AFFINE}`);
  }
  for (const edge of moduleEdges(affine)) {
    const detail = `${LAYOUT_AFFINE} -> ${edge.specifier}`;
    if (edge.kind !== 'import' || !edge.literal || edge.bare || !edge.typeOnly
      || edge.specifier !== './types.js') {
      fail('AFFINE_RUNTIME_DEPENDENCY', detail);
    }
  }
}

function assertTextRunProjectionAdapterBoundary(root) {
  const adapter = resolve(root, TEXT_RUN_PROJECTION_ADAPTER);
  if (!existsSync(adapter)) return;
  const declarationCounts = new Map();
  const source = sourceFile(adapter);
  for (const statement of source.statements) {
    if (ts.isExportDeclaration(statement)) {
      fail('TEXT_RUN_PROJECTION_DECLARATION', statement.getText(source));
    }
    for (const name of declarationNames(statement)) {
      if (!TEXT_RUN_PROJECTION_ADAPTER_DECLARATIONS.has(name)) {
        fail('TEXT_RUN_PROJECTION_DECLARATION', name);
      }
      declarationCounts.set(name, (declarationCounts.get(name) ?? 0) + 1);
      const shouldExport = name !== 'projectTextRun';
      if (hasExportModifier(statement) !== shouldExport) {
        fail('TEXT_RUN_PROJECTION_DECLARATION', name);
      }
    }
  }
  for (const name of TEXT_RUN_PROJECTION_ADAPTER_DECLARATIONS) {
    if (declarationCounts.get(name) !== 1) {
      fail(
        'TEXT_RUN_PROJECTION_DECLARATION',
        `${name} count ${declarationCounts.get(name) ?? 0}`,
      );
    }
  }
  const expected = new Map([
    ['@silurus/ooxml-core', [
      { typeOnly: false, names: ['PT_TO_PX', 'canvasFontString'] },
    ]],
    [`${DOCX_SOURCE}/renderer.ts`, [
      { typeOnly: true, names: ['DocxTextRunInfo'] },
    ]],
    [`${LAYOUT_SOURCE}/affine.ts`, [
      { typeOnly: false, names: ['composeAffine', 'mapAffinePoint', 'scaleAffine'] },
    ]],
    [`${LAYOUT_SOURCE}/document-layout-variants.ts`, [
      { typeOnly: false, names: ['selectDocumentLayoutPage'] },
    ]],
    [`${LAYOUT_SOURCE}/text-index.ts`, [
      { typeOnly: false, names: ['textRunGeometryForPage'] },
      { typeOnly: true, names: ['TextRunGeometry'] },
    ]],
    [`${LAYOUT_SOURCE}/types.ts`, [
      { typeOnly: true, names: ['DocumentLayout', 'LayoutServices', 'Matrix2DData'] },
    ]],
    [`${PAINT_SOURCE}/affine.ts`, [
      { typeOnly: false, names: ['cssTransformFor'] },
    ]],
  ]);
  const actual = new Map();
  for (const edge of moduleEdges(adapter)) {
    if (edge.kind !== 'import' || !edge.literal || edge.bare || edge.aliased
      || edge.specifier.includes('?') || edge.specifier.includes('#')) {
      fail('TEXT_RUN_PROJECTION_IMPORT', `${TEXT_RUN_PROJECTION_ADAPTER} -> ${edge.specifier}`);
    }
    const key = edge.specifier.startsWith('.')
      ? (() => {
          const target = resolveLocalImport(adapter, edge.specifier);
          return target ? posixPath(relative(root, target)) : null;
        })()
      : edge.specifier;
    if (key === null || !expected.has(key)) {
      fail('TEXT_RUN_PROJECTION_IMPORT', `${TEXT_RUN_PROJECTION_ADAPTER} -> ${edge.specifier}`);
    }
    const entries = actual.get(key) ?? [];
    entries.push({
      typeOnly: edge.typeOnly,
      names: [...(edge.importedNames ?? [])].sort(),
    });
    actual.set(key, entries);
  }
  const normalize = (entries) => [...entries]
    .map(({ typeOnly, names }) => `${typeOnly ? 'type' : 'value'}:${[...names].sort().join(',')}`)
    .sort();
  for (const [key, expectedEntries] of expected) {
    const actualEntries = actual.get(key) ?? [];
    if (JSON.stringify(normalize(actualEntries)) !== JSON.stringify(normalize(expectedEntries))) {
      fail(
        'TEXT_RUN_PROJECTION_IMPORT',
        `${TEXT_RUN_PROJECTION_ADAPTER} -> ${key} expected ${normalize(expectedEntries).join('|')}, received ${normalize(actualEntries).join('|')}`,
      );
    }
  }
}

/**
 * Body paint consumes retained placements. Letting this adapter call a layout
 * entry point would silently reintroduce a second layout pass whose result can
 * diverge from pagination (especially for grouped frames). Keep the rule tied
 * to the adapter's AST instead of relying on naming conventions in paint files:
 * renderer.ts intentionally still owns legacy header/footer story layout.
 */
function assertBodyPaintConsumesRetainedLayout(root) {
  const path = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(path)) return;
  const program = ts.createProgram({
    rootNames: [path],
    options: {
      target: ts.ScriptTarget.Latest,
      module: ts.ModuleKind.ESNext,
      noResolve: true,
      skipLibCheck: true,
    },
  });
  const source = program.getSourceFile(path);
  if (!source) return;
  const checker = program.getTypeChecker();
  const declaration = source.statements.find((statement) => (
    ts.isFunctionDeclaration(statement)
      && statement.name?.text === 'renderBodyElements'
  ));
  if (!declaration?.body) return;

  const forbidden = new Set([
    'acquireParagraphLayout',
    'acquireRetainedFrameGroup',
    'buildSegments',
    'contextualSpacingAdjust',
    'estimateParagraphHeight',
    'layoutLines',
    'measureParagraph',
    'measureText',
    'paragraphGapAdjustment',
    'paragraphLayoutFromMeasurement',
    'parasShareBorderBox',
    'renderFrameParagraph',
    'renderParagraph',
    'resolveParagraphBorderEdges',
    'resolveFrameBox',
  ]);
  const retainedCanvasMethods = new Set([
    'restore',
    'save',
    'scale',
    'translate',
  ]);
  const isRetainedPropertyBoundary = (target) => (
    (target.name === 'onTextRun' && target.receiver === 'state')
    || (retainedCanvasMethods.has(target.name) && target.receiver === 'state.ctx')
  );
  const retainedImportBoundaries = new Map([
    ['./paint/canvas-drawing.js', new Set(['paintDrawingLayout'])],
    ['./paint/canvas-text.js', new Set([
      'paintParagraphLayout',
      'paintPlacedParagraphLayout',
      'paintPlacedTextBoxLayout',
      'paintTextBoxLayout',
    ])],
    ['./vertical-text.js', new Set(['verticalTextLayerPlacement'])],
  ]);

  const unwrapExpression = (expression) => {
    let current = expression;
    while (ts.isParenthesizedExpression(current)
      || ts.isAsExpression(current)
      || ts.isTypeAssertionExpression(current)
      || ts.isNonNullExpression(current)
      || ts.isSatisfiesExpression(current)) {
      current = current.expression;
    }
    return current;
  };

  const isUnshadowedGlobalIdentifier = (node, name) => {
    if (!ts.isIdentifier(node) || node.text !== name) return false;
    const symbol = checker.getSymbolAtLocation(node);
    const declarations = symbol?.declarations ?? [];
    return declarations.length > 0
      && declarations.every((item) => item.getSourceFile().isDeclarationFile);
  };

  const isCanonicalWeakMapConstruction = (expression) => {
    const value = unwrapExpression(expression);
    return ts.isNewExpression(value)
      && isUnshadowedGlobalIdentifier(value.expression, 'WeakMap')
      && (value.arguments?.length ?? 0) === 0;
  };

  const isGlobalObjectCall = (expression, method) => (
    ts.isCallExpression(expression)
    && ts.isPropertyAccessExpression(expression.expression)
    && expression.expression.name.text === method
    && isUnshadowedGlobalIdentifier(expression.expression.expression, 'Object')
  );

  const isCanonicalBodyFragmentMapInitializer = (expression) => {
    const frozen = unwrapExpression(expression);
    if (!isGlobalObjectCall(frozen, 'freeze') || frozen.arguments.length !== 1) return false;
    const assigned = unwrapExpression(frozen.arguments[0]);
    if (!isGlobalObjectCall(assigned, 'assign')
      || assigned.arguments.length !== 2
      || !isCanonicalWeakMapConstruction(assigned.arguments[0])) return false;
    const sidecars = unwrapExpression(assigned.arguments[1]);
    if (!ts.isObjectLiteralExpression(sidecars) || sidecars.properties.length !== 2) return false;
    const expectedSidecars = new Set(['framePlacement', 'sourceIndices']);
    const seen = new Set();
    for (const property of sidecars.properties) {
      if (!ts.isPropertyAssignment(property) || !ts.isIdentifier(property.name)) return false;
      const name = property.name.text;
      if (!expectedSidecars.has(name)
        || seen.has(name)
        || !isCanonicalWeakMapConstruction(property.initializer)) return false;
      seen.add(name);
    }
    return seen.size === expectedSidecars.size;
  };

  const isCanonicalBodyFragmentReceiver = (identifier) => {
    const symbol = checker.getSymbolAtLocation(identifier);
    const declarations = symbol?.declarations ?? [];
    const declaration = declarations.length === 1 ? declarations[0] : undefined;
    return declaration !== undefined
      && ts.isVariableDeclaration(declaration)
      && ts.isVariableDeclarationList(declaration.parent)
      && (declaration.parent.flags & ts.NodeFlags.Const) !== 0
      && ts.isVariableStatement(declaration.parent.parent)
      && declaration.parent.parent.parent === source
      && declaration.initializer !== undefined
      && isCanonicalBodyFragmentMapInitializer(declaration.initializer);
  };

  const isExactBodyFragmentLookup = (item) => {
    if (item.name?.text !== 'bodyFragmentFor'
      || item.parameters.length !== 1
      || !ts.isIdentifier(item.parameters[0].name)
      || item.body?.statements.length !== 1) return false;
    const statement = item.body.statements[0];
    if (!ts.isReturnStatement(statement) || !statement.expression) return false;
    const call = unwrapExpression(statement.expression);
    if (!ts.isCallExpression(call)
      || !ts.isPropertyAccessExpression(call.expression)
      || !ts.isIdentifier(call.expression.expression)
      || call.expression.expression.text !== 'bodyFlowFragments'
      || !isCanonicalBodyFragmentReceiver(call.expression.expression)
      || call.expression.name.text !== 'get'
      || call.arguments.length !== 1) return false;
    const argument = unwrapExpression(call.arguments[0]);
    return ts.isIdentifier(argument)
      && checker.getSymbolAtLocation(argument)
        === checker.getSymbolAtLocation(item.parameters[0].name);
  };

  const staticString = (expression, resolving = new Set()) => {
    const value = unwrapExpression(expression);
    if (ts.isStringLiteralLike(value)) return value.text;
    if (ts.isBinaryExpression(value)
      && value.operatorToken.kind === ts.SyntaxKind.PlusToken) {
      const left = staticString(value.left, resolving);
      const right = staticString(value.right, resolving);
      return left === null || right === null ? null : left + right;
    }
    if (!ts.isIdentifier(value)) return null;
    const symbol = checker.getSymbolAtLocation(value);
    if (!symbol || resolving.has(symbol)) return null;
    resolving.add(symbol);
    const declarations = symbol.declarations ?? [];
    const values = declarations.flatMap((item) => (
      ts.isVariableDeclaration(item) && item.initializer
        ? [staticString(item.initializer, resolving)]
        : []
    ));
    resolving.delete(symbol);
    return values.length > 0 && values.every((item) => item === values[0])
      ? values[0]
      : null;
  };

  const propertyNameText = (name) => {
    if (ts.isIdentifier(name) || ts.isStringLiteralLike(name)) return name.text;
    if (ts.isComputedPropertyName(name)) return staticString(name.expression);
    return null;
  };

  const canonicalBodyFragmentSidecarPath = (expression) => {
    const path = [];
    let current = unwrapExpression(expression);
    while (ts.isPropertyAccessExpression(current) || ts.isElementAccessExpression(current)) {
      const name = ts.isPropertyAccessExpression(current)
        ? current.name.text
        : current.argumentExpression
          ? staticString(current.argumentExpression)
          : null;
      if (name === null) return null;
      path.unshift(name);
      current = unwrapExpression(current.expression);
    }
    return ts.isIdentifier(current)
      && current.text === 'bodyFlowFragments'
      && isCanonicalBodyFragmentReceiver(current)
      && (path[0] === 'sourceIndices' || path[0] === 'framePlacement')
      ? path
      : null;
  };
  const assignmentOperators = new Set([
    ts.SyntaxKind.EqualsToken,
    ts.SyntaxKind.PlusEqualsToken,
    ts.SyntaxKind.MinusEqualsToken,
    ts.SyntaxKind.AsteriskEqualsToken,
    ts.SyntaxKind.AsteriskAsteriskEqualsToken,
    ts.SyntaxKind.SlashEqualsToken,
    ts.SyntaxKind.PercentEqualsToken,
    ts.SyntaxKind.LessThanLessThanEqualsToken,
    ts.SyntaxKind.GreaterThanGreaterThanEqualsToken,
    ts.SyntaxKind.GreaterThanGreaterThanGreaterThanEqualsToken,
    ts.SyntaxKind.AmpersandEqualsToken,
    ts.SyntaxKind.BarEqualsToken,
    ts.SyntaxKind.CaretEqualsToken,
    ts.SyntaxKind.BarBarEqualsToken,
    ts.SyntaxKind.AmpersandAmpersandEqualsToken,
    ts.SyntaxKind.QuestionQuestionEqualsToken,
  ]);
  const lateSidecarMutations = [];
  const findLateSidecarMutations = (current) => {
    if (ts.isCallExpression(current)
      && (isGlobalObjectCall(current, 'assign') || isGlobalObjectCall(current, 'defineProperty'))
      && current.arguments[0]
      && canonicalBodyFragmentSidecarPath(current.arguments[0])) {
      lateSidecarMutations.push(current.getText(source));
    } else if (ts.isBinaryExpression(current)
      && assignmentOperators.has(current.operatorToken.kind)
      && canonicalBodyFragmentSidecarPath(current.left)) {
      lateSidecarMutations.push(current.getText(source));
    }
    ts.forEachChild(current, findLateSidecarMutations);
  };
  findLateSidecarMutations(source);
  if (lateSidecarMutations.length > 0) {
    fail(
      'BODY_PAINT_LAYOUT_CAPABILITY',
      `${DOCX_SOURCE}/renderer.ts mutates canonical body fragment sidecar after initialization`,
    );
  }

  let resolveCallTarget;
  const resolveObjectProperty = (expression, propertyName, resolving) => {
    const value = unwrapExpression(expression);
    if (ts.isObjectLiteralExpression(value)) {
      const property = value.properties.find((item) => (
        'name' in item && item.name && propertyNameText(item.name) === propertyName
      ));
      if (property && ts.isPropertyAssignment(property)) {
        return resolveCallTarget(property.initializer, resolving);
      }
      if (property && ts.isShorthandPropertyAssignment(property)) {
        return resolveCallTarget(property.name, resolving);
      }
      if (property && ts.isMethodDeclaration(property) && property.body) {
        return [{ kind: 'body', body: property.body, detail: propertyName }];
      }
      return null;
    }
    if (ts.isConditionalExpression(value)) {
      const whenTrue = resolveObjectProperty(value.whenTrue, propertyName, resolving);
      const whenFalse = resolveObjectProperty(value.whenFalse, propertyName, resolving);
      return whenTrue && whenFalse ? [...whenTrue, ...whenFalse] : null;
    }
    if (!ts.isIdentifier(value)) return null;
    const symbol = checker.getSymbolAtLocation(value);
    if (!symbol || resolving.has(symbol)) return null;
    resolving.add(symbol);
    const targets = (symbol.declarations ?? []).flatMap((item) => (
      ts.isVariableDeclaration(item) && item.initializer
        ? resolveObjectProperty(item.initializer, propertyName, resolving) ?? []
        : []
    ));
    resolving.delete(symbol);
    return targets.length > 0 ? targets : null;
  };

  resolveCallTarget = (expression, resolving = new Set()) => {
    const value = unwrapExpression(expression);
    if (ts.isPropertyAccessExpression(value)) {
      return resolveObjectProperty(value.expression, value.name.text, resolving)
        ?? [{ kind: 'property', name: value.name.text, receiver: value.expression.getText(source) }];
    }
    if (ts.isElementAccessExpression(value)) {
      const name = value.argumentExpression && staticString(value.argumentExpression);
      return name === null || name === undefined
        ? [{ kind: 'unresolved', detail: value.getText(source) }]
        : resolveObjectProperty(value.expression, name, resolving)
          ?? [{ kind: 'property', name, receiver: value.expression.getText(source) }];
    }
    if (ts.isConditionalExpression(value)) {
      return [
        ...resolveCallTarget(value.whenTrue, resolving),
        ...resolveCallTarget(value.whenFalse, resolving),
      ];
    }
    if (!ts.isIdentifier(value)) {
      return [{ kind: 'unresolved', detail: value.getText(source) }];
    }
    const symbol = checker.getSymbolAtLocation(value);
    if (!symbol) return [{ kind: 'name', name: value.text }];
    if (resolving.has(symbol)) {
      return [{ kind: 'unresolved', detail: value.text }];
    }
    resolving.add(symbol);
    const targets = [];
    for (const item of symbol.declarations ?? []) {
      if (ts.isFunctionDeclaration(item) && item.body) {
        const name = item.name?.text;
        targets.push(isExactBodyFragmentLookup(item)
          ? { kind: 'local-boundary', name: 'bodyFragmentFor' }
          : { kind: 'body', body: item.body, detail: name ?? value.text });
      } else if (ts.isVariableDeclaration(item) && item.initializer) {
        const initializer = unwrapExpression(item.initializer);
        if (ts.isArrowFunction(initializer) || ts.isFunctionExpression(initializer)) {
          targets.push({ kind: 'body', body: initializer.body, detail: value.text });
        } else {
          targets.push(...resolveCallTarget(initializer, resolving));
        }
      } else if (ts.isImportSpecifier(item)) {
        let parent = item.parent;
        while (parent && !ts.isImportDeclaration(parent)) parent = parent.parent;
        targets.push({
          kind: 'import',
          name: item.propertyName?.text ?? item.name.text,
          specifier: parent && ts.isStringLiteral(parent.moduleSpecifier)
            ? parent.moduleSpecifier.text
            : null,
        });
      } else if (ts.isBindingElement(item)) {
        const name = !item.propertyName
          ? ts.isIdentifier(item.name) ? item.name.text : null
          : propertyNameText(item.propertyName);
        const variable = ts.isObjectBindingPattern(item.parent)
          && ts.isVariableDeclaration(item.parent.parent)
          ? item.parent.parent
          : null;
        const propertyTargets = name !== null && variable?.initializer
          ? resolveObjectProperty(variable.initializer, name, resolving)
          : null;
        targets.push(...(propertyTargets ?? [name === null
          ? { kind: 'unresolved', detail: item.getText(source) }
          : {
              kind: 'property',
              name,
              receiver: variable?.initializer?.getText(source) ?? '<destructured>',
            }]));
      } else if (ts.isParameter(item)) {
        targets.push({ kind: 'unresolved', detail: value.text });
      }
    }
    resolving.delete(symbol);
    return targets.length > 0
      ? targets
      : [{ kind: 'unresolved', detail: value.text }];
  };

  const directCallTargets = (node) => {
    const targets = [];
    const visit = (current) => {
      const isNamedLocalCallable = current !== node && (
        ts.isFunctionDeclaration(current)
        || ((ts.isFunctionExpression(current) || ts.isArrowFunction(current))
          && ts.isVariableDeclaration(current.parent)
          && ts.isIdentifier(current.parent.name))
      );
      if (isNamedLocalCallable) return;
      if (ts.isCallExpression(current)) {
        targets.push(...resolveCallTarget(current.expression));
      }
      ts.forEachChild(current, visit);
    };
    visit(node);
    return targets;
  };
  const paragraphBranchOf = (statement) => {
    if (!ts.isBinaryExpression(statement.expression)) return null;
    const { left, operatorToken, right } = statement.expression;
    const isTypeAccess = (node) => ts.isPropertyAccessExpression(node)
      && node.name.text === 'type';
    const isParagraph = (node) => ts.isStringLiteralLike(node)
      && node.text === 'paragraph';
    if (!((isTypeAccess(left) && isParagraph(right))
      || (isParagraph(left) && isTypeAccess(right)))) return null;
    if (operatorToken.kind === ts.SyntaxKind.EqualsEqualsEqualsToken
      || operatorToken.kind === ts.SyntaxKind.EqualsEqualsToken) {
      return statement.thenStatement;
    }
    if (operatorToken.kind === ts.SyntaxKind.ExclamationEqualsEqualsToken
      || operatorToken.kind === ts.SyntaxKind.ExclamationEqualsToken) {
      return statement.elseStatement ?? null;
    }
    return null;
  };

  const entryCalls = new Set();
  let foundParagraphBranch = false;
  const findParagraphBranches = (node) => {
    if (ts.isIfStatement(node)) {
      const branch = paragraphBranchOf(node);
      if (branch) {
        foundParagraphBranch = true;
        for (const target of directCallTargets(branch)) entryCalls.add(target);
        return;
      }
    }
    ts.forEachChild(node, findParagraphBranches);
  };
  findParagraphBranches(declaration.body);
  if (!foundParagraphBranch) {
    fail(
      'BODY_PAINT_LAYOUT_CAPABILITY',
      `${DOCX_SOURCE}/renderer.ts#renderBodyElements has no statically auditable paragraph branch`,
    );
  }

  const violations = new Set();
  const visitedBodies = new Set();
  const pending = [...entryCalls];
  while (pending.length > 0) {
    const target = pending.pop();
    if (!target) continue;
    if (target.kind === 'unresolved') {
      violations.add(`unresolved call ${target.detail}`);
      continue;
    }
    if (target.kind === 'local-boundary') continue;
    if (target.kind === 'import') {
      if (retainedImportBoundaries.get(target.specifier)?.has(target.name)) continue;
      violations.add(forbidden.has(target.name)
        ? target.name
        : `unresolved call ${target.name} from ${target.specifier ?? '<unknown import>'}`);
      continue;
    }
    if (target.kind === 'property') {
      if (isRetainedPropertyBoundary(target)) continue;
      violations.add(forbidden.has(target.name)
        ? target.name
        : `unresolved call ${target.name}`);
      continue;
    }
    if (target.kind === 'name') {
      violations.add(forbidden.has(target.name)
        ? target.name
        : `unresolved call ${target.name}`);
      continue;
    }
    if (visitedBodies.has(target.body)) continue;
    visitedBodies.add(target.body);
    for (const called of directCallTargets(target.body)) pending.push(called);
  }

  if (violations.size > 0) {
    fail(
      'BODY_PAINT_LAYOUT_CAPABILITY',
      `${DOCX_SOURCE}/renderer.ts#renderBodyElements reaches ${[...violations].sort().join(', ')}`,
    );
  }
}

function isExactLayoutParserModelGatewayImportEdge(currentRel, edge, dependency, parserModel) {
  return currentRel === LAYOUT_PARSER_MODEL_GATEWAY
    && dependency === parserModel
    && edge.kind === 'import'
    && edge.specifier === LAYOUT_PARSER_MODEL_GATEWAY_IMPORT
    && edge.typeOnly === false
    && edge.aliased === false
    && edge.bare === false
    && edge.importedNames?.length === 1
    && edge.importedNames[0] === LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL;
}

function hasExactLayoutParserModelGatewayProjection(path) {
  const source = sourceFile(path);
  let bindingReferences = 0;
  const countBindingReferences = (node) => {
    if (ts.isIdentifier(node) && node.text === LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL) {
      bindingReferences += 1;
    }
    ts.forEachChild(node, countBindingReferences);
  };
  countBindingReferences(source);
  if (bindingReferences !== 2) return false;

  const projections = source.statements.filter((statement) => (
    ts.isFunctionDeclaration(statement)
    && statement.name?.text === 'documentMathOccurrences'
  ));
  if (projections.length !== 1) return false;
  const projection = projections[0];
  const exported = projection.modifiers?.some((modifier) => (
    modifier.kind === ts.SyntaxKind.ExportKeyword
  ));
  if (!exported || !projection.body || projection.parameters.length !== 1) return false;
  const parameter = projection.parameters[0];
  if (!ts.isIdentifier(parameter.name) || projection.body.statements.length !== 1) return false;
  const returned = projection.body.statements[0];
  if (!ts.isReturnStatement(returned)
    || !returned.expression
    || !ts.isArrayLiteralExpression(returned.expression)
    || returned.expression.elements.length !== 1) return false;
  const spread = returned.expression.elements[0];
  if (!ts.isSpreadElement(spread)
    || !ts.isPropertyAccessExpression(spread.expression)
    || spread.expression.name.text !== 'mathOccurrences') return false;
  const call = spread.expression.expression;
  return ts.isCallExpression(call)
    && ts.isIdentifier(call.expression)
    && call.expression.text === LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL
    && call.arguments.length === 1
    && ts.isIdentifier(call.arguments[0])
    && call.arguments[0].text === parameter.name.text;
}

function layoutParserModelBoundaryViolations(root) {
  const graph = dependencyGraph(root);
  const parserModel = resolve(root, PARSER_MODEL);
  const entries = [...graph.keys()].filter((path) => (
    posixPath(relative(root, path)).startsWith(`${LAYOUT_SOURCE}/`)
  ));
  const violations = [];
  const nonLiteral = [];

  for (const entry of entries) {
    const stack = [{ path: entry, chain: [entry] }];
    const visited = new Set([entry]);
    while (stack.length > 0) {
      const current = stack.pop();
      const currentRel = posixPath(relative(root, current.path));
      for (const edge of graph.get(current.path) ?? []) {
        if (!edge.literal) {
          nonLiteral.push(current.chain.map((path) => posixPath(relative(root, path))));
          continue;
        }
        if (!edge.specifier.startsWith('.')) continue;
        const dependency = resolveLocalImport(current.path, edge.specifier);
        if (!dependency) continue;
        const chain = [...current.chain, dependency];
        if (dependency === parserModel) {
          // The parser-model gateway permits exactly one projection edge. Only
          // that edge is terminal; every other resources.ts dependency remains
          // part of the transitive runtime graph and is inspected normally.
          if (isExactLayoutParserModelGatewayImportEdge(currentRel, edge, dependency, parserModel)) {
            if (hasExactLayoutParserModelGatewayProjection(current.path)) continue;
            violations.push([
              ...chain.map((path) => posixPath(relative(root, path))),
              `invalid use of ${LAYOUT_PARSER_MODEL_GATEWAY_SYMBOL}`,
            ]);
            continue;
          }
          violations.push(chain.map((path) => posixPath(relative(root, path))));
          continue;
        }
        // A type-only edge is erased and cannot create a runtime parser-model
        // dependency through the referenced contract. A direct parser-model
        // type import was rejected above so layout stays parser-model-free.
        if (edge.typeOnly) continue;
        if (graph.has(dependency) && !visited.has(dependency)) {
          visited.add(dependency);
          stack.push({ path: dependency, chain });
        }
      }
    }
  }
  return { violations, nonLiteral };
}

function assertLayoutParserModelBoundaries(root) {
  const { violations, nonLiteral } = layoutParserModelBoundaryViolations(root);
  if (nonLiteral.length > 0) {
    fail(
      'NON_LITERAL_LAYOUT_MODULE_EDGE',
      nonLiteral.map((chain) => chain.join(' -> ')).join('\n'),
    );
  }
  if (violations.length > 0) {
    fail(
      'LAYOUT_PARSER_MODEL_DEPENDENCY',
      violations.map((chain) => chain.join(' -> ')).join('\n'),
    );
  }
}

function assertBodyLayoutAdapterBoundary(root) {
  const adapter = resolve(root, BODY_LAYOUT_ADAPTER);
  if (!existsSync(adapter)) return;
  const source = sourceFile(adapter);
  const seenImports = new Map();
  const declarations = [];
  for (const statement of source.statements) {
    if (ts.isExportDeclaration(statement) || ts.isExportAssignment(statement)) {
      fail('BODY_LAYOUT_ADAPTER_EXPORT', statement.getText(source));
    }
    if (statement.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.DefaultKeyword)) {
      fail('BODY_LAYOUT_ADAPTER_EXPORT', statement.getText(source));
    }
    if (ts.isImportDeclaration(statement)) {
      if (!ts.isStringLiteralLike(statement.moduleSpecifier)) {
        fail('BODY_LAYOUT_ADAPTER_IMPORT', '<dynamic>');
      }
      const dependency = resolveLocalImport(adapter, statement.moduleSpecifier.text);
      const dependencyRelative = dependency ? posixPath(relative(root, dependency)) : null;
      const reviewedBindings = dependencyRelative
        ? BODY_LAYOUT_ADAPTER_IMPORT_BINDINGS.get(dependencyRelative)
        : undefined;
      const clause = statement.importClause;
      if (!reviewedBindings || !clause || clause.name || !clause.namedBindings
        || !ts.isNamedImports(clause.namedBindings)) {
        fail('BODY_LAYOUT_ADAPTER_IMPORT', statement.getText(source));
      }
      if (seenImports.has(dependencyRelative)) {
        fail('BODY_LAYOUT_ADAPTER_IMPORT', `duplicate:${dependencyRelative}`);
      }
      const actualBindings = new Map();
      for (const element of clause.namedBindings.elements) {
        const importedName = element.propertyName?.text ?? element.name.text;
        const kind = clause.isTypeOnly || element.isTypeOnly ? 'type' : 'value';
        if (element.name.text !== importedName || reviewedBindings.get(importedName) !== kind) {
          fail('BODY_LAYOUT_ADAPTER_BINDING', `${dependencyRelative}#${importedName}:${kind}`);
        }
        actualBindings.set(importedName, kind);
      }
      if (actualBindings.size !== reviewedBindings.size
        || [...reviewedBindings].some(([name, kind]) => actualBindings.get(name) !== kind)) {
        fail('BODY_LAYOUT_ADAPTER_IMPORT', `incomplete:${dependencyRelative}`);
      }
      seenImports.set(dependencyRelative, actualBindings);
    }
    for (const name of declarationNames(statement)) {
      declarations.push(name);
      if (!BODY_LAYOUT_ADAPTER_DECLARATIONS.has(name)) {
        fail('BODY_LAYOUT_ADAPTER_DECLARATION', name);
      }
    }
  }
  if (declarations.length !== 1 || declarations[0] !== 'createBodyLayoutInput') {
    fail('BODY_LAYOUT_ADAPTER_DECLARATION', declarations.join(','));
  }
  if (seenImports.size !== BODY_LAYOUT_ADAPTER_IMPORT_BINDINGS.size
    || [...BODY_LAYOUT_ADAPTER_IMPORT_BINDINGS.keys()].some((path) => !seenImports.has(path))) {
    fail('BODY_LAYOUT_ADAPTER_IMPORT', 'exact-import-set-required');
  }
  const declaration = source.statements.find(ts.isFunctionDeclaration);
  const returned = declaration?.body?.statements.length === 1
    && ts.isReturnStatement(declaration.body.statements[0])
    ? declaration.body.statements[0].expression
    : null;
  const isAcquisitionCall = returned && ts.isCallExpression(returned)
    && ts.isIdentifier(returned.expression)
    && returned.expression.text === 'projectBodyLayoutInput'
    && returned.arguments.length === 1
    && ts.isCallExpression(returned.arguments[0])
    && ts.isIdentifier(returned.arguments[0].expression)
    && returned.arguments[0].expression.text === 'bodyLayoutAcquisitionInput'
    && returned.arguments[0].arguments.length === 1
    && ts.isIdentifier(returned.arguments[0].arguments[0])
    && returned.arguments[0].arguments[0].text === 'document';
  if (!declaration
    || declaration.name?.text !== 'createBodyLayoutInput'
    || declaration.parameters.length !== 1
    || !ts.isIdentifier(declaration.parameters[0].name)
    || declaration.parameters[0].name.text !== 'document'
    || !isAcquisitionCall) {
    fail('BODY_LAYOUT_ADAPTER_BODY', declaration?.getText(source) ?? '<missing>');
  }
}

function assertParagraphAnchorFrameAdapterBoundary(root, allowTransitionalAdapter) {
  const adapter = resolve(root, PARAGRAPH_ANCHOR_FRAME_ADAPTER);
  if (!existsSync(adapter)) return;
  if (!allowTransitionalAdapter) {
    fail('FINAL_PARAGRAPH_ANCHOR_ADAPTER', PARAGRAPH_ANCHOR_FRAME_ADAPTER);
  }

  const source = sourceFile(adapter);
  const imports = source.statements.filter(ts.isImportDeclaration);
  if (imports.length !== 1) {
    fail('PARAGRAPH_ANCHOR_ADAPTER_IMPORT', 'exact-import-set-required');
  }
  const importStatement = imports[0];
  const clause = importStatement.importClause;
  const bindings = clause?.namedBindings;
  if (!ts.isStringLiteralLike(importStatement.moduleSpecifier)
    || importStatement.moduleSpecifier.text !== './layout/anchor-frame.js'
    || !clause?.isTypeOnly
    || clause.name
    || !bindings
    || !ts.isNamedImports(bindings)
    || bindings.elements.length !== 1
    || bindings.elements[0]?.name.text !== 'AnchorReferenceFramesInput'
    || bindings.elements[0]?.propertyName) {
    fail('PARAGRAPH_ANCHOR_ADAPTER_IMPORT', importStatement.getText(source));
  }

  const declarations = [];
  for (const statement of source.statements) {
    if (ts.isExportDeclaration(statement) || ts.isExportAssignment(statement)
      || statement.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.DefaultKeyword)) {
      fail('PARAGRAPH_ANCHOR_ADAPTER_EXPORT', statement.getText(source));
    }
    for (const name of declarationNames(statement)) {
      declarations.push(name);
      if (!PARAGRAPH_ANCHOR_FRAME_ADAPTER_DECLARATIONS.has(name)) {
        fail('PARAGRAPH_ANCHOR_ADAPTER_DECLARATION', name);
      }
    }
  }
  const exactDeclarations = [...PARAGRAPH_ANCHOR_FRAME_ADAPTER_DECLARATIONS].sort();
  if (declarations.length !== exactDeclarations.length
    || declarations.sort().some((name, index) => name !== exactDeclarations[index])) {
    fail('PARAGRAPH_ANCHOR_ADAPTER_DECLARATION', declarations.join(','));
  }

  const runtimeConsumers = [];
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    if (path === adapter) continue;
    for (const edge of moduleEdges(path)) {
      if (!edge.literal || edge.typeOnly || !edge.specifier.startsWith('.')) continue;
      if (resolveLocalImport(path, edge.specifier) !== adapter) continue;
      const consumer = posixPath(relative(root, path));
      runtimeConsumers.push(consumer);
      if (consumer !== `${DOCX_SOURCE}/renderer.ts`
        || edge.kind !== 'import'
        || edge.aliased
        || edge.importedNames.length !== 1
        || edge.importedNames[0] !== 'paragraphAnchorReferenceFrames') {
        fail('PARAGRAPH_ANCHOR_ADAPTER_CONSUMER', consumer);
      }
    }
  }
  if (runtimeConsumers.length !== 1
    || runtimeConsumers[0] !== `${DOCX_SOURCE}/renderer.ts`) {
    fail('PARAGRAPH_ANCHOR_ADAPTER_CONSUMER', runtimeConsumers.join(','));
  }
}

function assertBodyKernelServiceOwner(root) {
  const renderer = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(renderer)) return;
  const source = sourceFile(renderer);
  const owner = source.statements.find((statement) => (
    ts.isFunctionDeclaration(statement) && statement.name?.text === 'createLayoutServices'
  ));
  const implementation = source.statements.find((statement) => (
    ts.isFunctionDeclaration(statement) && statement.name?.text === 'createConcreteBodyLayoutKernel'
  ));
  if (!owner?.body || !implementation?.body) {
    fail('BODY_KERNEL_SERVICE_OWNER', 'missing owner or implementation');
  }
  const calls = [];
  const identifiers = [];
  const visit = (node) => {
    if (ts.isIdentifier(node) && node.text === 'createConcreteBodyLayoutKernel') identifiers.push(node);
    if (ts.isCallExpression(node)
      && ts.isIdentifier(node.expression)
      && node.expression.text === 'createConcreteBodyLayoutKernel') calls.push(node);
    ts.forEachChild(node, visit);
  };
  visit(source);
  const call = calls[0];
  const parentCall = call?.parent;
  const owned = call
    && parentCall
    && ts.isCallExpression(parentCall)
    && ts.isIdentifier(parentCall.expression)
    && parentCall.expression.text === 'attachBodyLayoutKernel'
    && parentCall.arguments.length === 2
    && ts.isIdentifier(parentCall.arguments[0])
    && parentCall.arguments[0].text === 'services'
    && parentCall.arguments[1] === call
    && call.arguments.length === 3
    && call.arguments.every((argument, index) => (
      ts.isIdentifier(argument) && argument.text === ['doc', 'ctx', 'localMetrics'][index]
    ));
  let insideOwner = false;
  for (let node = parentCall; node; node = node.parent) {
    if (node === owner) insideOwner = true;
  }
  if (calls.length !== 1 || identifiers.length !== 2 || !owned || !insideOwner) {
    fail('BODY_KERNEL_SERVICE_OWNER', `calls:${calls.length};identifiers:${identifiers.length}`);
  }
}

const MIGRATION_IDENTIFIER = /(?:legacy|(?:use|enable|prefer|require)[a-z0-9]*(?:old|previous|alternate)[a-z0-9]*(?:engine|layout|path|algorithm)|(?:reuse|paint)enabled|requireslegacy|dryrun)/i;

function matchingIdentifierCounts(root, predicate) {
  const counts = {};
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const source = sourceFile(path);
    const file = posixPath(relative(root, path));
    const visit = (node) => {
      if (ts.isIdentifier(node) && predicate(node.text)) {
        const key = `${file}#${node.text}`;
        counts[key] = (counts[key] ?? 0) + 1;
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
  return Object.fromEntries(Object.entries(counts).sort(([left], [right]) => left.localeCompare(right)));
}

function assertNoDeletedPageProducerIdentifiers(root) {
  const matches = {};
  const sourceRoot = resolve(root, DOCX_SOURCE);
  const record = (file, name) => {
    const key = `${file}#${name}`;
    matches[key] = (matches[key] ?? 0) + 1;
  };
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const source = sourceFile(path);
    const file = posixPath(relative(root, path));
    const visit = (node) => {
      if (ts.isIdentifier(node) && DELETED_PAGE_PRODUCER_IDENTIFIERS.has(node.text)) {
        record(file, node.text);
      } else if ((ts.isIdentifier(node) || ts.isStringLiteralLike(node))
        && DELETED_LEGACY_STAMP_PROPERTIES.has(node.text)) {
        const parent = node.parent;
        const isStaticPropertyName = (
          (ts.isPropertyAccessExpression(parent) && parent.name === node)
          || (ts.isElementAccessExpression(parent) && parent.argumentExpression === node)
          || (ts.isPropertyAssignment(parent) && parent.name === node)
          || (ts.isShorthandPropertyAssignment(parent) && parent.name === node)
          || (ts.isPropertyDeclaration(parent) && parent.name === node)
          || (ts.isPropertySignature(parent) && parent.name === node)
          || (ts.isMethodDeclaration(parent) && parent.name === node)
          || (ts.isGetAccessorDeclaration(parent) && parent.name === node)
          || (ts.isSetAccessorDeclaration(parent) && parent.name === node)
          || (ts.isBindingElement(parent) && (parent.propertyName ?? parent.name) === node)
        );
        if (isStaticPropertyName) record(file, node.text);
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
  }
  if (Object.keys(matches).length > 0) {
    fail('FORBIDDEN_PAGE_PRODUCER_IDENTIFIER', stableJson(matches).trim());
  }
}

function callsNamed(node, name) {
  const calls = [];
  const visit = (current) => {
    if (ts.isCallExpression(current)
      && ts.isIdentifier(current.expression)
      && current.expression.text === name) calls.push(current);
    ts.forEachChild(current, visit);
  };
  visit(node);
  return calls;
}

function unwrapStaticExpression(expression) {
  let current = expression;
  while (ts.isParenthesizedExpression(current)
    || ts.isAsExpression(current)
    || ts.isTypeAssertionExpression(current)
    || ts.isNonNullExpression(current)
    || ts.isSatisfiesExpression(current)) {
    current = current.expression;
  }
  return current;
}

function callOf(expression, name) {
  const value = expression && unwrapStaticExpression(expression);
  return value
    && ts.isCallExpression(value)
    && ts.isIdentifier(value.expression)
    && value.expression.text === name
    ? value
    : null;
}

function hasExactInvariantImports(source) {
  return source.statements.some((statement) => {
    if (!ts.isImportDeclaration(statement)
      || !ts.isStringLiteral(statement.moduleSpecifier)
      || statement.moduleSpecifier.text !== './invariants.js'
      || !statement.importClause
      || statement.importClause.isTypeOnly
      || statement.importClause.name
      || !statement.importClause.namedBindings
      || !ts.isNamedImports(statement.importClause.namedBindings)) return false;
    const elements = statement.importClause.namedBindings.elements;
    return elements.length === 2
      && elements.every((element) => !element.isTypeOnly && !element.propertyName)
      && new Set(elements.map((element) => element.name.text)).size === 2
      && elements.some((element) => element.name.text === 'assertDocumentLayout')
      && elements.some((element) => element.name.text === 'deepFreezeDocumentLayout');
  });
}

function isCanonicalSelectedLayoutFunction(statement, source) {
  if (!ts.isFunctionDeclaration(statement)
    || statement.name?.text !== 'selectDocumentLayoutPage'
    || !hasExportModifier(statement)
    || !statement.body
    || statement.parameters.length !== 3
    || statement.parameters.some((parameter) => !ts.isIdentifier(parameter.name))
    || statement.body.statements.length !== 3) return false;
  const [services, input, pageIndex] = statement.parameters.map((parameter) => parameter.name.text);
  const [storeDeclaration, missingStore, selectedPage] = statement.body.statements;
  if (!ts.isVariableStatement(storeDeclaration)
    || (storeDeclaration.declarationList.flags & ts.NodeFlags.Const) === 0
    || storeDeclaration.declarationList.declarations.length !== 1) return false;
  const [store] = storeDeclaration.declarationList.declarations;
  const storeCall = callOf(store.initializer, 'layoutVariantStoreOf');
  if (!ts.isIdentifier(store.name)
    || !storeCall
    || storeCall.arguments.length !== 1
    || storeCall.arguments[0].getText(source) !== services) return false;
  const storeName = store.name.text;
  if (!ts.isIfStatement(missingStore)
    || missingStore.elseStatement
    || !ts.isPrefixUnaryExpression(missingStore.expression)
    || missingStore.expression.operator !== ts.SyntaxKind.ExclamationToken
    || missingStore.expression.operand.getText(source) !== storeName
    || !ts.isThrowStatement(missingStore.thenStatement)) return false;
  if (!ts.isReturnStatement(selectedPage) || !selectedPage.expression) return false;
  const selection = unwrapStaticExpression(selectedPage.expression);
  if (!ts.isCallExpression(selection)
    || !ts.isPropertyAccessExpression(selection.expression)
    || selection.expression.expression.getText(source) !== storeName
    || selection.expression.name.text !== 'selectPage'
    || selection.arguments.length !== 2) return false;
  const normalized = callOf(selection.arguments[0], 'layoutOptionsForRender');
  return normalized !== null
    && normalized.arguments.length === 1
    && normalized.arguments[0].getText(source) === input
    && selection.arguments[1].getText(source) === pageIndex;
}

function objectProperty(object, name) {
  return object.properties.find((property) => (
    (ts.isPropertyAssignment(property)
      || ts.isShorthandPropertyAssignment(property)
      || ts.isMethodDeclaration(property))
    && (ts.isIdentifier(property.name) || ts.isStringLiteralLike(property.name))
    && property.name.text === name
  ));
}

function isCanonicalWorkerVariantAttachment(call, source) {
  if (call.arguments.length !== 1) return false;
  const value = unwrapStaticExpression(call.arguments[0]);
  if (!ts.isObjectLiteralExpression(value)) return false;
  const model = objectProperty(value, 'model');
  const services = objectProperty(value, 'services');
  const buildLayout = objectProperty(value, 'buildLayout');
  if (!model || !services || !ts.isPropertyAssignment(buildLayout)) return false;
  const modelExpression = ts.isShorthandPropertyAssignment(model) ? model.name : model.initializer;
  const servicesExpression = ts.isShorthandPropertyAssignment(services) ? services.name : services.initializer;
  const builder = unwrapStaticExpression(buildLayout.initializer);
  if (!ts.isArrowFunction(builder)
    || builder.parameters.length !== 1
    || !ts.isIdentifier(builder.parameters[0].name)) return false;
  const layoutCall = callOf(builder.body, 'layoutDocument');
  return layoutCall !== null
    && layoutCall.arguments.length === 3
    && layoutCall.arguments[0].getText(source) === modelExpression.getText(source)
    && layoutCall.arguments[1].getText(source) === servicesExpression.getText(source)
    && layoutCall.arguments[2].getText(source) === builder.parameters[0].name.text;
}

function exactObjectPropertyIdentities(object, identities, source) {
  if (object.properties.length !== identities.size) return false;
  for (const [name, expected] of identities) {
    const property = objectProperty(object, name);
    const value = ts.isShorthandPropertyAssignment(property)
      ? property.name
      : ts.isPropertyAssignment(property)
        ? property.initializer
        : null;
    if (value?.getText(source) !== expected) return false;
  }
  return true;
}

function returnedObjectLiteral(body) {
  const returns = body.statements.filter(ts.isReturnStatement);
  if (returns.length !== 1 || !returns[0].expression) return null;
  let expression = unwrapStaticExpression(returns[0].expression);
  if (ts.isCallExpression(expression)
    && ts.isPropertyAccessExpression(expression.expression)
    && expression.expression.expression.getText() === 'Object'
    && expression.expression.name.text === 'freeze'
    && expression.arguments.length === 1) {
    expression = unwrapStaticExpression(expression.arguments[0]);
  }
  return ts.isObjectLiteralExpression(expression) ? expression : null;
}

function workerRenderCallIsCanonical(call, source) {
  if (call.arguments.length !== 4) return false;
  const options = unwrapStaticExpression(call.arguments[3]);
  if (!ts.isObjectLiteralExpression(options)) return false;
  const services = objectProperty(options, 'layoutServices');
  const defaultDate = objectProperty(options, 'defaultCurrentDateMs');
  return ts.isPropertyAssignment(services)
    && services.initializer.getText(source) === 'doc.layoutServices'
    && ts.isPropertyAssignment(defaultDate)
    && defaultDate.initializer.getText(source) === 'doc.defaultCurrentDateMs';
}

function workerRetentionCallIsCanonical(call, source) {
  return call.arguments.length === 3
    && call.arguments[0].getText(source) === 'model'
    && call.arguments[1].getText(source) === 'layoutServices'
    && call.arguments[2].getText(source) === 'req.defaultCurrentDateMs';
}

function workerRetentionSeamIsCanonical(source) {
  const declarations = source.statements.flatMap(declarationNames).sort();
  const exactDeclarations = [...WORKER_LAYOUT_RETENTION_DECLARATIONS].sort();
  if (declarations.length !== exactDeclarations.length
    || declarations.some((name, index) => name !== exactDeclarations[index])) return false;
  const retention = source.statements.find((statement) => (
    ts.isFunctionDeclaration(statement)
    && statement.name?.text === 'retainRenderWorkerDocumentLayout'
  ));
  if (!retention?.body
    || retention.parameters.length !== 3
    || retention.parameters.some((parameter) => !ts.isIdentifier(parameter.name))
    || retention.parameters.map((parameter) => parameter.name.getText(source)).join(',')
      !== 'model,layoutServices,defaultCurrentDateMs') return false;
  const attachments = callsNamed(retention.body, 'attachDocumentLayoutVariants');
  if (attachments.length !== 1
    || callsNamed(retention.body, 'layoutDocument').length !== 1
    || !isCanonicalWorkerVariantAttachment(attachments[0], source)) return false;
  const attachment = unwrapStaticExpression(attachments[0].arguments[0]);
  if (!ts.isObjectLiteralExpression(attachment)) return false;
  const model = objectProperty(attachment, 'model');
  const services = objectProperty(attachment, 'services');
  const defaultDate = objectProperty(attachment, 'defaultCurrentDateMs');
  const modelExpression = ts.isShorthandPropertyAssignment(model)
    ? model.name
    : ts.isPropertyAssignment(model)
      ? model.initializer
      : null;
  const servicesExpression = ts.isShorthandPropertyAssignment(services)
    ? services.name
    : ts.isPropertyAssignment(services)
      ? services.initializer
      : null;
  const defaultDateExpression = ts.isShorthandPropertyAssignment(defaultDate)
    ? defaultDate.name
    : ts.isPropertyAssignment(defaultDate)
      ? defaultDate.initializer
      : null;
  const returned = returnedObjectLiteral(retention.body);
  return modelExpression?.getText(source) === 'model'
    && servicesExpression?.getText(source) === 'layoutServices'
    && defaultDateExpression?.getText(source) === 'defaultCurrentDateMs'
    && returned !== null
    && exactObjectPropertyIdentities(returned, new Map([
      ['model', 'model'],
      ['layoutServices', 'layoutServices'],
      ['layoutVariants', 'variants.store'],
      ['defaultCurrentDateMs', 'defaultCurrentDateMs'],
    ]), source);
}

function variableDeclarationsNamed(node, name) {
  const declarations = [];
  const visit = (current) => {
    if (ts.isVariableDeclaration(current)
      && ts.isIdentifier(current.name)
      && current.name.text === name) declarations.push(current);
    ts.forEachChild(current, visit);
  };
  visit(node);
  return declarations;
}

function workerMetadataRouteIsCanonical(source) {
  const layouts = variableDeclarationsNamed(source, 'layout');
  const pageSizeLists = variableDeclarationsNamed(source, 'pageSizes');
  const metadata = variableDeclarationsNamed(source, 'meta');
  if (layouts.length !== 1
    || layouts[0].initializer?.getText(source) !== 'doc.layoutVariants.defaultLayout'
    || pageSizeLists.length !== 1
    || metadata.length !== 1) return false;
  const pageSizesInitializer = pageSizeLists[0].initializer
    && unwrapStaticExpression(pageSizeLists[0].initializer);
  if (!pageSizesInitializer
    || !ts.isCallExpression(pageSizesInitializer)
    || !ts.isPropertyAccessExpression(pageSizesInitializer.expression)
    || pageSizesInitializer.expression.name.text !== 'map'
    || pageSizesInitializer.expression.expression.getText(source) !== 'layout.pages') return false;
  const value = metadata[0].initializer && unwrapStaticExpression(metadata[0].initializer);
  if (!value || !ts.isObjectLiteralExpression(value)) return false;
  const pageCount = objectProperty(value, 'pageCount');
  const pageSizes = objectProperty(value, 'pageSizes');
  const bookmarks = objectProperty(value, 'bookmarkPages');
  if (!ts.isPropertyAssignment(pageCount)
    || pageCount.initializer.getText(source) !== 'layout.pages.length'
    || (!ts.isShorthandPropertyAssignment(pageSizes)
      && !(ts.isPropertyAssignment(pageSizes)
        && pageSizes.initializer.getText(source) === 'pageSizes'))
    || !ts.isPropertyAssignment(bookmarks)
    || !ts.isArrayLiteralExpression(bookmarks.initializer)
    || bookmarks.initializer.elements.length !== 1
    || !ts.isSpreadElement(bookmarks.initializer.elements[0])) return false;
  const bookmarkCall = callOf(bookmarks.initializer.elements[0].expression, 'buildBookmarkPageMap');
  return bookmarkCall?.arguments.length === 1
    && bookmarkCall.arguments[0].getText(source) === 'layout';
}

function assertCanonicalCutoverBoundaries(root) {
  const paginatorPath = resolve(root, LAYOUT_SOURCE, 'body-paginator.ts');
  if (!existsSync(paginatorPath)) {
    fail('CANONICAL_LAYOUT_PRODUCER', `${LAYOUT_SOURCE}/body-paginator.ts#paginateBody`);
  } else {
    const source = sourceFile(paginatorPath);
    const exportedValues = source.statements.filter((statement) => (
      hasExportModifier(statement)
      && (ts.isFunctionDeclaration(statement)
        || ts.isVariableStatement(statement)
        || ts.isClassDeclaration(statement)
        || ts.isEnumDeclaration(statement))
    ));
    const runtimeExportForms = source.statements.filter((statement) => (
      ts.isExportAssignment(statement)
      || (ts.isExportDeclaration(statement) && !exportIsTypeOnly(statement))
    ));
    const producer = exportedValues.find((statement) => (
      ts.isFunctionDeclaration(statement) && statement.name?.text === 'paginateBody'
    ));
    if (!producer?.body || exportedValues.length !== 1 || runtimeExportForms.length !== 0) {
      fail('CANONICAL_LAYOUT_PRODUCER', `${LAYOUT_SOURCE}/body-paginator.ts#paginateBody`);
    }
    const validation = producer.body.statements.at(-2);
    const returned = producer.body.statements.at(-1);
    const validationCall = validation && ts.isExpressionStatement(validation)
      ? callOf(validation.expression, 'assertDocumentLayout')
      : null;
    const frozenCall = returned && ts.isReturnStatement(returned)
      ? callOf(returned.expression, 'deepFreezeDocumentLayout')
      : null;
    if (!hasExactInvariantImports(source)
      || callsNamed(producer.body, 'assertDocumentLayout').length !== 1
      || callsNamed(producer.body, 'deepFreezeDocumentLayout').length !== 1
      || validationCall?.arguments.length !== 1
      || frozenCall?.arguments.length !== 1
      || validationCall.arguments[0].getText(source)
        !== frozenCall.arguments[0].getText(source)) {
      fail('RETAINED_LAYOUT_IMMUTABILITY', `${LAYOUT_SOURCE}/body-paginator.ts#paginateBody`);
    }
  }

  const variantsPath = resolve(root, LAYOUT_SOURCE, 'document-layout-variants.ts');
  if (!existsSync(variantsPath)) {
    fail('SELECTED_LAYOUT_VARIANT', `${LAYOUT_SOURCE}/document-layout-variants.ts#selectDocumentLayoutPage`);
  } else {
    const source = sourceFile(variantsPath);
    const selection = source.statements.find((statement) => (
      ts.isFunctionDeclaration(statement)
      && statement.name?.text === 'selectDocumentLayoutPage'
    ));
    if (!selection || !isCanonicalSelectedLayoutFunction(selection, source)) {
      fail('SELECTED_LAYOUT_VARIANT', `${LAYOUT_SOURCE}/document-layout-variants.ts#selectDocumentLayoutPage`);
    }
  }

  const workerPath = resolve(root, DOCX_SOURCE, 'render-worker.ts');
  const workerRetentionPath = resolve(root, DOCX_SOURCE, 'render-worker-layout.ts');
  if (!existsSync(workerPath)) {
    fail('WORKER_LAYOUT_SELECTION', `${DOCX_SOURCE}/render-worker.ts`);
  } else if (!existsSync(workerRetentionPath)
    || !workerRetentionSeamIsCanonical(sourceFile(workerRetentionPath))) {
    fail('WORKER_LAYOUT_SELECTION', `${DOCX_SOURCE}/render-worker-layout.ts`);
  } else {
    const source = sourceFile(workerPath);
    const topLevelPages = source.statements.some((statement) => (
      ts.isVariableStatement(statement)
      && statement.declarationList.declarations.some((declaration) => (
        ts.isIdentifier(declaration.name) && declaration.name.text === 'pages'
      ))
    ));
    let duplicateSelection = false;
    const visit = (node) => {
      if (ts.isIdentifier(node) && node.text === 'selectDocumentLayoutPage') {
        duplicateSelection = true;
      }
      ts.forEachChild(node, visit);
    };
    visit(source);
    const variantAttachments = callsNamed(source, 'attachDocumentLayoutVariants');
    const retentionCalls = callsNamed(source, 'retainRenderWorkerDocumentLayout');
    const rendererCalls = callsNamed(source, 'renderDocumentToCanvas');
    if (topLevelPages
      || duplicateSelection
      || variantAttachments.length !== 0
      || callsNamed(source, 'layoutDocument').length !== 0
      || retentionCalls.length !== 1
      || !workerRetentionCallIsCanonical(retentionCalls[0], source)
      || rendererCalls.length !== 1
      || rendererCalls.some((call) => !workerRenderCallIsCanonical(call, source))
      || !workerMetadataRouteIsCanonical(source)) {
      fail('WORKER_LAYOUT_SELECTION', `${DOCX_SOURCE}/render-worker.ts`);
    }

    const rendererPath = resolve(root, DOCX_SOURCE, 'renderer.ts');
    if (existsSync(rendererPath)
      && callsNamed(sourceFile(rendererPath), 'selectDocumentLayoutPage').length !== 1) {
      fail('SELECTED_LAYOUT_VARIANT', `${DOCX_SOURCE}/renderer.ts#selectDocumentLayoutPage`);
    }
  }
}

function identifierCounts(root) {
  const counts = {};
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const fileCounts = Object.fromEntries(LEGACY_SYMBOLS.map((symbol) => [symbol, 0]));
    const source = sourceFile(path);
    const visit = (node) => {
      if (ts.isIdentifier(node) && Object.hasOwn(fileCounts, node.text)) fileCounts[node.text] += 1;
      ts.forEachChild(node, visit);
    };
    visit(source);
    const file = posixPath(relative(root, path));
    for (const [symbol, count] of Object.entries(fileCounts)) {
      if (count > 0) counts[`${file}#${symbol}`] = count;
    }
  }
  return Object.fromEntries(Object.entries(counts).sort(([left], [right]) => left.localeCompare(right)));
}

function bindingNames(name, names = []) {
  if (ts.isIdentifier(name)) names.push(name.text);
  else for (const element of name.elements) {
    if (!ts.isOmittedExpression(element)) bindingNames(element.name, names);
  }
  return names;
}

function declarationNames(statement) {
  if ((ts.isFunctionDeclaration(statement)
      || ts.isClassDeclaration(statement)
      || ts.isInterfaceDeclaration(statement)
      || ts.isTypeAliasDeclaration(statement)
      || ts.isEnumDeclaration(statement)
      || ts.isModuleDeclaration(statement))
    && statement.name) {
    return [statement.name.text];
  }
  if (ts.isVariableStatement(statement)) {
    return statement.declarationList.declarations.flatMap((declaration) => bindingNames(declaration.name));
  }
  return [];
}

function declarationKind(statement) {
  if (ts.isVariableStatement(statement)) return 'variable';
  return ts.SyntaxKind[statement.kind];
}

function normalizedNodeHash(node, source) {
  const normalized = node.getText(source).replace(/\s+/g, ' ').trim();
  return createHash('sha256').update(normalized).digest('hex');
}

function normalizedSyntaxHash(node, source) {
  const shape = (current) => {
    const text = ts.isIdentifier(current) || ts.isLiteralExpression(current)
      ? current.getText(source)
      : undefined;
    const children = [];
    current.forEachChild((child) => children.push(shape(child)));
    return [ts.SyntaxKind[current.kind], text, ...children];
  };
  return createHash('sha256').update(JSON.stringify(shape(node))).digest('hex');
}

/** A2 permits one mechanically constrained edit to computePages: append the
 * two dependency parameters, then append those identifiers to its existing
 * buildMeasureState call. A5 additionally replaces the exact floating-slice
 * table stamp fold with one retained-slice size lookup. Everything else remains
 * represented in the hash. */
function normalizedComputePagesHash(node, source) {
  const compactText = (current, currentSource) =>
    current.getText(currentSource).replace(/\s+/g, '');
  // A6 effective table-flow ([MS-OI29500] §2.1.162): the paginator decides
  // ordinary-flow participation and floating placement through
  // `tableParticipatesInOrdinaryFlow` / `effectiveTablePositioning` instead of a
  // lexical `tblpPr` presence test. Normalize ONLY the exact paired two-site
  // transformation back onto the frozen lexical base:
  //   1. `if (!tableParticipatesInOrdinaryFlow(t)) continue;` -> `if (t.tblpPr) continue;`
  //   2. `const tp = effectiveTablePositioning(tbl); if (tp) { … }` folds back to
  //      `if (tbl.tblpPr) { const tp = tbl.tblpPr; … }` (the pre-resolution moves
  //      into the branch as its first statement).
  // Both sites must be present exactly once; one-site, duplicated, moved,
  // predicate/callee/argument-altered, or side-effect-adjacent forms fall through
  // and stay fully represented in the hash. Section-occurrence routing is never
  // touched here.
  //
  // Reconstruction rewrites both sites back to plain `tblpPr` value syntax, which
  // erases any TypeScript-only syntax the effective form might carry: a variable
  // type annotation (`const tp: TblpPr = …`), a definite-assignment `!`, an
  // optional-call `?.(…)`, or call type arguments (`effectiveTablePositioning<…>(tbl)`).
  // If such syntax were tolerated it would vanish from the resulting hash and go
  // unaudited, so every matcher below rejects it and lets the declaration stay
  // fully represented instead.
  const effectiveFlowSkipSites = [];
  const findEffectiveFlowSkip = (current) => {
    if (ts.isIfStatement(current)
      && !current.elseStatement
      && ts.isPrefixUnaryExpression(current.expression)
      && current.expression.operator === ts.SyntaxKind.ExclamationToken
      && ts.isCallExpression(current.expression.operand)
      && current.expression.operand.questionDotToken === undefined
      && current.expression.operand.typeArguments === undefined
      && ts.isIdentifier(current.expression.operand.expression)
      && current.expression.operand.expression.text === 'tableParticipatesInOrdinaryFlow'
      && current.expression.operand.arguments.length === 1
      && ts.isIdentifier(current.expression.operand.arguments[0])
      && current.expression.operand.arguments[0].text === 't'
      && ts.isContinueStatement(current.thenStatement)
      && !current.thenStatement.label) {
      effectiveFlowSkipSites.push(current);
    }
    ts.forEachChild(current, findEffectiveFlowSkip);
  };
  findEffectiveFlowSkip(node);
  const isEffectiveTpDeclaration = (statement) => {
    if (!ts.isVariableStatement(statement)
      || (statement.declarationList.flags & ts.NodeFlags.Const) === 0
      || statement.declarationList.declarations.length !== 1) return false;
    const [declaration] = statement.declarationList.declarations;
    return ts.isIdentifier(declaration.name)
      && declaration.name.text === 'tp'
      && declaration.type === undefined
      && declaration.exclamationToken === undefined
      && declaration.initializer !== undefined
      && ts.isCallExpression(declaration.initializer)
      && declaration.initializer.questionDotToken === undefined
      && declaration.initializer.typeArguments === undefined
      && ts.isIdentifier(declaration.initializer.expression)
      && declaration.initializer.expression.text === 'effectiveTablePositioning'
      && declaration.initializer.arguments.length === 1
      && ts.isIdentifier(declaration.initializer.arguments[0])
      && declaration.initializer.arguments[0].text === 'tbl';
  };
  const effectiveFlowAcquireSites = [];
  const findEffectiveFlowAcquire = (current) => {
    const statements = ts.isBlock(current) || ts.isSourceFile(current)
      ? current.statements
      : undefined;
    if (statements) {
      for (let index = 0; index + 1 < statements.length; index += 1) {
        const declaration = statements[index];
        const branch = statements[index + 1];
        if (isEffectiveTpDeclaration(declaration)
          && ts.isIfStatement(branch)
          && !branch.elseStatement
          && ts.isIdentifier(branch.expression)
          && branch.expression.text === 'tp'
          && ts.isBlock(branch.thenStatement)) {
          effectiveFlowAcquireSites.push({ declaration, branch });
        }
      }
    }
    ts.forEachChild(current, findEffectiveFlowAcquire);
  };
  findEffectiveFlowAcquire(node);
  if (effectiveFlowSkipSites.length === 1 && effectiveFlowAcquireSites.length === 1) {
    const nodeStart = node.getStart(source);
    const nodeText = node.getText(source);
    const skip = effectiveFlowSkipSites[0];
    const { declaration, branch } = effectiveFlowAcquireSites[0];
    const blockOpen = branch.thenStatement.getStart(source) + 1;
    const edits = [
      [skip.getStart(source), skip.getEnd(), 'if (t.tblpPr) continue;'],
      [declaration.getStart(source), declaration.getEnd(), ''],
      [branch.expression.getStart(source), branch.expression.getEnd(), 'tbl.tblpPr'],
      [blockOpen, blockOpen, '\nconst tp = tbl.tblpPr;'],
    ].sort((left, right) => right[0] - left[0]);
    let virtualText = nodeText;
    for (const [start, end, replacement] of edits) {
      virtualText = virtualText.slice(0, start - nodeStart)
        + replacement
        + virtualText.slice(end - nodeStart);
    }
    const virtualSource = ts.createSourceFile(
      'compute-pages-a6-effective-flow-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((statement) => (
      ts.isFunctionDeclaration(statement) && statement.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const occurrenceOwnerReplacements = [
    ['computeTableRowHeights(ms, t, colWPt, j)', 'computeTableRowHeights(ms, t, colWPt)'],
    [
      'estimateTableHeight(measureState, nxt as unknown as DocTable, colW(), startIdx)',
      'estimateTableHeight(measureState, nxt as unknown as DocTable, colW())',
    ],
    [
      'measureState, para, i, colW, suppressBefore, colX,',
      'measureState, para, colW, suppressBefore, colX,',
    ],
    [
      `        const occurrenceEl = { ...el } as PaginatedElementWithLines;
        attachBodyParagraphFragment(occurrenceEl, para, measureState, i, {
          paragraphXPt: colX(),
          availableWidthPt: colW(),
          suppressSpaceBefore: suppressBefore,
          columnIndex: colIndex,
        }, fitMeasured);
        pushTagged(occurrenceEl);`,
      `        attachBodyParagraphFragment(el as PaginatedElementWithLines, para, measureState, i, {
          paragraphXPt: colX(),
          availableWidthPt: colW(),
          suppressSpaceBefore: suppressBefore,
          columnIndex: colIndex,
        }, fitMeasured);
        pushTagged(el as PaginatedBodyElement);`,
    ],
    [
      'attachBodyParagraphFragment(el as PaginatedElementWithLines, para, measureState, i, {',
      'attachBodyParagraphFragment(el as PaginatedElementWithLines, para, measureState, {',
    ],
    ['computeTableLayout(tbl, cW, measureState, i)', 'computeTableLayout(tbl, cW, measureState)'],
    ['            const retainedRecord = retainedTableRecord(measureState, i);\n', ''],
    [
      `              const prepared = prepareFittingOuterFragment(
                tbl, i, retainedRecord, finalState, box,
              );`,
      `              const prepared = bodyFlowFragments.sourceIndices.retainedTableMeasureBySource
                .prepareFittingOuterFragment(tbl, finalState, box);`,
    ],
    [
      `        const occurrenceEl = { ...el } as PaginatedBodyElement;
        withColumnBand(() => {
          stampTableLayout(
            occurrenceEl,
            first.layout.colWidths,
            first.layout.rowHeights,
            first.contentWPt,
            i,
            retainedTableRecord(measureState, i),
            measureState,
            undefined,
            acceptedPrepared,
          );
          const side = floatTableWrapSide(first.box, measureState);
          registerTableFloat(
            first.box, tp, measureState, side, tbl.overlap !== 'never', true,
          );
        });
        pushTagged(occurrenceEl);`,
      `        withColumnBand(() => {
          stampTableLayout(
            el as PaginatedBodyElement,
            first.layout.colWidths,
            first.layout.rowHeights,
            first.contentWPt,
            i,
            retainedTableRecord(measureState, i),
            measureState,
            undefined,
            acceptedPrepared,
          );
          const side = floatTableWrapSide(first.box, measureState);
          registerTableFloat(
            first.box, tp, measureState, side, tbl.overlap !== 'never', true,
          );
        });
        pushTagged(el as PaginatedBodyElement);`,
    ],
    [
      `            first.contentWPt,
            i,
            retainedTableRecord(measureState, i),
            measureState,
            undefined,
            acceptedPrepared,`,
      `            first.contentWPt,
            undefined,
            acceptedPrepared,`,
    ],
    [
      '            { sourceIndex: i, record: retainedTableRecord(measureState, i), state: measureState },\n',
      '',
    ],
    ['computeTablePtLayout(measureState, tbl, bandPt, i)', 'computeTablePtLayout(measureState, tbl, bandPt)'],
    [
      `          bandPt,
          i,
          retainedTableRecord(measureState, i),
          measureState,
          {`,
      `          bandPt,
          {`,
    ],
    [
      'computeTablePtLayout(measureState, tbl, tblContentWPt, i)',
      'computeTablePtLayout(measureState, tbl, tblContentWPt)',
    ],
    [
      `        tableContentH,
        measureState,
        i,
        true,`,
      `        tableContentH,
        measureState,
        true,`,
    ],
    [
      `              tblContentWPt,
              measureState,
              i,
              {`,
      `              tblContentWPt,
              measureState,
              {`,
    ],
    ['                fragment: meta.fragment,\n', ''],
    [
      `          () => currentSectionFrame.textDirection,
          i,
        );`,
      `          () => currentSectionFrame.textDirection,
        );`,
    ],
    [
      `          tblContentWPt,
          measureState,
          i,
          {`,
      `          tblContentWPt,
          measureState,
          {`,
    ],
  ];
  const currentOccurrenceOwnerText = node.getText(source);
  const priorFittingProbeOccurrenceReplacements = occurrenceOwnerReplacements.filter(
    (_replacement, index) => ![5, 6, 7, 8, 9].includes(index),
  );
  const hasExactPriorFittingProbe = currentOccurrenceOwnerText.includes(
    'const layout = computeTableLayout(tbl, cW, measureState);',
  ) && currentOccurrenceOwnerText.includes(
    'stampTableLayout(\n            el as PaginatedBodyElement,\n            first.layout.colWidths,',
  );
  const selectedOccurrenceOwnerReplacements = hasExactPriorFittingProbe
    ? priorFittingProbeOccurrenceReplacements
    : occurrenceOwnerReplacements;
  let occurrenceOwnerText = currentOccurrenceOwnerText;
  let hasExactOccurrenceOwnerThreading = occurrenceOwnerText.includes(
    '{ sourceIndex: i, record: retainedTableRecord(measureState, i), state: measureState }',
  );
  if (hasExactOccurrenceOwnerThreading) {
    for (const [current, previous] of selectedOccurrenceOwnerReplacements) {
      if (occurrenceOwnerText.split(current).length !== 2) {
        hasExactOccurrenceOwnerThreading = false;
        break;
      }
      occurrenceOwnerText = occurrenceOwnerText.replace(current, previous);
    }
  }
  if (hasExactOccurrenceOwnerThreading) {
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-occurrence-owner-virtual.ts',
      occurrenceOwnerText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((candidate) => (
      ts.isFunctionDeclaration(candidate) && candidate.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const exactFittingOuterProbe = `
    const measureFloat = () =>
      withColumnBand(() => {
        const cW = colW() * measureState.scale;
        const layout = computeTableLayout(tbl, cW, measureState);
        const physicalPageIndex = pages.length - 1;
        const pageContext = measureState.layoutServices
          ? fieldAcquisitionContextOf(measureState.layoutServices)
            .resolveDestinationPage?.(physicalPageIndex)
          : undefined;
        const finalState: RenderState = {
          ...measureState,
          pageIndex: physicalPageIndex,
          displayPageNumber: pageContext?.displayPageNumber ?? physicalPageIndex + 1,
          pageNumberFormat: pageContext?.pageNumberFormat ?? measureState.pageNumberFormat,
        };
        let finalHeightPt = layout.rowHeights.reduce((sum, height) => sum + height, 0);
        let box = computeFloatTableBox(
          tp, finalState, finalState.y, layout.tableW, finalHeightPt, false,
          { allowOverlap: tbl.overlap !== 'never' },
        );
        for (let pass = 0; pass < 4; pass += 1) {
          const prepared = bodyFlowFragments.sourceIndices.retainedTableMeasureBySource
            .prepareFittingOuterFragment(tbl, finalState, box);
          if (!prepared.fragment) {
            const rawBox = computeFloatTableBox(
              tp, finalState, finalState.y, layout.tableW, finalHeightPt, true,
            );
            return {
              box,
              rawBox,
              layout,
              requiresCanonicalSplit: true as const,
              contentWPt: cW / finalState.scale,
            };
          }
          const nextHeightPt = prepared.fragment.advancePt;
          const nextBox = computeFloatTableBox(
            tp, finalState, finalState.y, layout.tableW, nextHeightPt, false,
            { allowOverlap: tbl.overlap !== 'never' },
          );
          const rawBox = computeFloatTableBox(
            tp, finalState, finalState.y, layout.tableW, nextHeightPt, true,
          );
          if (nextHeightPt === finalHeightPt
            && nextBox.x === box.x && nextBox.y === box.y) {
            return {
              box: nextBox,
              rawBox,
              layout,
              prepared,
              requiresCanonicalSplit: false as const,
              contentWPt: cW / finalState.scale,
            };
          }
          finalHeightPt = nextHeightPt;
          box = nextBox;
        }
        throw new Error('Fitting outer table final-frame probe did not converge');
      });
  `.replace(/\s+/g, '');
  const exactPreviousFittingOuterProbe = exactFittingOuterProbe
    .replace(
      "computeFloatTableBox(tp,finalState,finalState.y,layout.tableW,finalHeightPt,false,{allowOverlap:tbl.overlap!=='never'},)",
      'computeFloatTableBox(tp,finalState,finalState.y,layout.tableW,finalHeightPt,)',
    )
    .replace(
      "computeFloatTableBox(tp,finalState,finalState.y,layout.tableW,nextHeightPt,false,{allowOverlap:tbl.overlap!=='never'},)",
      'computeFloatTableBox(tp,finalState,finalState.y,layout.tableW,nextHeightPt,)',
    );
  const exactFittingOuterSplitCondition =
    'first.requiresCanonicalSplit||(isTextAnchored&&tableOverflowsHere)||pageAnchoredOverflows';
  const exactFittingOuterAcceptance = `
    if (!first.prepared) {
      throw new Error('Fitting outer table acceptance requires a whole prepared fragment');
    }
    const acceptedFragment = first.prepared.fragment;
    if (!acceptedFragment) {
      throw new Error('Fitting outer table acceptance requires a whole prepared fragment');
    }
    const acceptedPrepared = {
      ...first.prepared,
      fragment: acceptedFragment,
      box: first.box,
    };
    withColumnBand(() => {
      stampTableLayout(
        el as PaginatedBodyElement,
        first.layout.colWidths,
        first.layout.rowHeights,
        first.contentWPt,
        undefined,
        acceptedPrepared,
      );
      const side = floatTableWrapSide(first.box, measureState);
      registerTableFloat(
        first.box, tp, measureState, side, tbl.overlap !== 'never', true,
      );
    });
  `.replace(/\s+/g, '');
  const exactPreviousFittingOuterAcceptance = exactFittingOuterAcceptance.replace(
    "registerTableFloat(first.box,tp,measureState,side,tbl.overlap!=='never',true,);",
    "registerTableFloat(first.box,tp,measureState,side,tbl.overlap!=='never');",
  );
  const fittingOuterProbes = [];
  const fittingOuterSplitConditions = [];
  const fittingOuterAcceptances = [];
  const findFittingOuterProbeTransaction = (current) => {
    if (ts.isVariableStatement(current)) {
      const compact = compactText(current, source);
      if (compact === exactFittingOuterProbe || compact === exactPreviousFittingOuterProbe) {
        fittingOuterProbes.push({
          node: current,
          variant: compact === exactFittingOuterProbe ? 'pre-resolved' : 'previous',
        });
      }
    }
    if (ts.isIfStatement(current)
      && compactText(current.expression, source) === exactFittingOuterSplitCondition) {
      fittingOuterSplitConditions.push(current.expression);
    }
    if (ts.isBlock(current)) {
      for (let index = 0; index + 4 < current.statements.length; index += 1) {
        const statements = current.statements.slice(index, index + 5);
        const compact = statements.map((statement) => compactText(statement, source)).join('');
        if (compact === exactFittingOuterAcceptance
          || compact === exactPreviousFittingOuterAcceptance) {
          fittingOuterAcceptances.push({
            range: [statements[0], statements[4]],
            variant: compact === exactFittingOuterAcceptance ? 'pre-resolved' : 'previous',
          });
        }
      }
    }
    ts.forEachChild(current, findFittingOuterProbeTransaction);
  };
  findFittingOuterProbeTransaction(node);
  if (fittingOuterProbes.length === 1
    && fittingOuterSplitConditions.length === 1
    && fittingOuterAcceptances.length === 1
    && fittingOuterProbes[0].variant === fittingOuterAcceptances[0].variant) {
    const nodeStart = node.getStart(source);
    const legacyProbe = `
      const measureFloat = () =>
        withColumnBand(() => {
          const cW = colW() * measureState.scale;
          const layout = computeTableLayout(tbl, cW, measureState);
          const tableH = layout.rowHeights.reduce((s, x) => s + x, 0);
          const box = computeFloatTableBox(
            tp, measureState, measureState.y, layout.tableW, tableH,
          );
          const rawBox = computeFloatTableBox(
            tp, measureState, measureState.y, layout.tableW, tableH, true,
          );
          return { box, rawBox, layout, contentWPt: cW / measureState.scale };
        });
    `;
    const legacySplitCondition =
      '(isTextAnchored && tableOverflowsHere) || pageAnchoredOverflows';
    const legacyAcceptance = `
      withColumnBand(() => {
        stampTableLayout(
          el as PaginatedBodyElement,
          first.layout.colWidths,
          first.layout.rowHeights,
          first.contentWPt,
        );
        const side = floatTableWrapSide(first.box, measureState);
        registerTableFloat(first.box, tp, measureState, side, tbl.overlap !== 'never');
      });
    `;
    const [acceptanceStart, acceptanceEnd] = fittingOuterAcceptances[0].range;
    const replacements = [
      [
        fittingOuterProbes[0].node.getStart(source),
        fittingOuterProbes[0].node.getEnd(),
        legacyProbe,
      ],
      [
        fittingOuterSplitConditions[0].getStart(source),
        fittingOuterSplitConditions[0].getEnd(),
        legacySplitCondition,
      ],
      [acceptanceStart.getStart(source), acceptanceEnd.getEnd(), legacyAcceptance],
    ].sort((left, right) => right[0] - left[0]);
    let virtualText = node.getText(source);
    for (const [start, end, replacement] of replacements) {
      virtualText = virtualText.slice(0, start - nodeStart)
        + replacement
        + virtualText.slice(end - nodeStart);
    }
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-fitting-probe-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((candidate) => (
      ts.isFunctionDeclaration(candidate) && candidate.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const exactConvergedSplitParentResolverSource = `
    (sliceTp, tableWidthPt, tableHeightPt, externalRegistry) =>
      withColumnBand(() => {
        const skipVClamp = sliceTp.vertAnchor === 'page' || sliceTp.vertAnchor === 'margin';
        const stateAgainstExternalRegistry: RenderState = {
          ...measureState,
          floats: [...externalRegistry.floats],
          floatParaSeq: externalRegistry.nextParagraphId,
        };
        return computeFloatTableBox(
          sliceTp,
          stateAgainstExternalRegistry,
          measureState.y,
          tableWidthPt * measureState.scale,
          tableHeightPt * measureState.scale,
          skipVClamp,
          { allowOverlap: tbl.overlap !== 'never' },
        );
      })
  `;
  const exactConvergedSplitEmitterSource = '(sliceEl) => pushTagged(sliceEl)';
  const convergedSplitCallbacksSource = ts.createSourceFile(
    'compute-pages-a5-converged-split-parent-expected.ts',
    `const callbacks = [${exactConvergedSplitParentResolverSource}, ${exactConvergedSplitEmitterSource}];`,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
  const convergedSplitCallbacksDeclaration = convergedSplitCallbacksSource.statements[0]
    ?.declarationList?.declarations?.[0];
  const convergedSplitCallbacks = convergedSplitCallbacksDeclaration?.initializer?.elements;
  const syntaxPrinter = ts.createPrinter({ removeComments: true });
  const printedSyntax = (current, currentSource) => syntaxPrinter
    .printNode(ts.EmitHint.Unspecified, current, currentSource)
    .replace(/\s+/g, '');
  const expectedConvergedSplitResolver = convergedSplitCallbacks?.[0]
    ? printedSyntax(convergedSplitCallbacks[0], convergedSplitCallbacksSource)
    : '';
  const expectedConvergedSplitEmitter = convergedSplitCallbacks?.[1]
    ? printedSyntax(convergedSplitCallbacks[1], convergedSplitCallbacksSource)
    : '';
  const exactSplitFloatLivePageSource = '() => pages.length - 1';
  const splitFloatLivePageSource = ts.createSourceFile(
    'compute-pages-a5-split-float-live-page-expected.ts',
    `const callback = ${exactSplitFloatLivePageSource};`,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
  const splitFloatLivePageDeclaration = splitFloatLivePageSource.statements[0]
    ?.declarationList?.declarations?.[0];
  const splitFloatLivePageCallback = splitFloatLivePageDeclaration?.initializer;
  const expectedSplitFloatLivePage = splitFloatLivePageCallback
    ? printedSyntax(splitFloatLivePageCallback, splitFloatLivePageSource)
    : '';
  const splitFloatLivePageCalls = [];
  const findSplitFloatLivePageCall = (current) => {
    if (ts.isCallExpression(current)
      && ts.isIdentifier(current.expression)
      && current.expression.text === 'splitFloatTableAcrossPages'
      && current.arguments.length >= 3) {
      const resolver = current.arguments[current.arguments.length - 3];
      const emitter = current.arguments[current.arguments.length - 2];
      const livePage = current.arguments[current.arguments.length - 1];
      if (printedSyntax(resolver, source) === expectedConvergedSplitResolver
        && printedSyntax(emitter, source) === expectedConvergedSplitEmitter
        && printedSyntax(livePage, source) === expectedSplitFloatLivePage) {
        splitFloatLivePageCalls.push({ emitter, livePage });
      }
    }
    ts.forEachChild(current, findSplitFloatLivePageCall);
  };
  findSplitFloatLivePageCall(node);
  if (splitFloatLivePageCalls.length === 1) {
    const { emitter, livePage } = splitFloatLivePageCalls[0];
    const nodeStart = node.getStart(source);
    const start = emitter.getEnd() - nodeStart;
    const end = livePage.getEnd() - nodeStart;
    const nodeText = node.getText(source);
    const virtualText = nodeText.slice(0, start) + nodeText.slice(end);
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-split-float-live-page-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((candidate) => (
      ts.isFunctionDeclaration(candidate) && candidate.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const convergedSplitCalls = [];
  const findConvergedSplitCall = (current) => {
    if (ts.isCallExpression(current)
      && ts.isIdentifier(current.expression)
      && current.expression.text === 'splitFloatTableAcrossPages'
      && current.arguments.length >= 2) {
      const resolver = current.arguments[current.arguments.length - 2];
      const emitter = current.arguments[current.arguments.length - 1];
      if (printedSyntax(resolver, source) === expectedConvergedSplitResolver
        && printedSyntax(emitter, source) === expectedConvergedSplitEmitter) {
        convergedSplitCalls.push({ resolver, emitter });
      }
    }
    ts.forEachChild(current, findConvergedSplitCall);
  };
  findConvergedSplitCall(node);
  if (convergedSplitCalls.length === 1) {
    const { resolver, emitter } = convergedSplitCalls[0];
    const nodeStart = node.getStart(source);
    const start = resolver.getStart(source) - nodeStart;
    const end = emitter.getEnd() - nodeStart;
    const previousCallback = `
      (sliceEl) => {
        pushTagged(sliceEl);
        return withColumnBand(() => {
          const sp = sliceEl as PaginatedBodyElement;
          const sliceTp = (sliceEl as unknown as DocTable).tblpPr as TblpPr;
          const { widthPx: tableW, heightPx: sliceH } = retainedTableSliceSize(
            sp, measureState.scale,
          );
          const skipVClamp = sliceTp.vertAnchor === 'page' || sliceTp.vertAnchor === 'margin';
          return computeFloatTableBox(
            sliceTp, measureState, measureState.y, tableW, sliceH, skipVClamp,
            { allowOverlap: tbl.overlap !== 'never' },
          );
        });
      }
    `;
    const nodeText = node.getText(source);
    const virtualText = nodeText.slice(0, start) + previousCallback + nodeText.slice(end);
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-converged-split-parent-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((candidate) => (
      ts.isFunctionDeclaration(candidate) && candidate.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const exactSplitParentPreResolutionSource = `
    (sliceEl) => {
      pushTagged(sliceEl);
      return withColumnBand(() => {
        const sp = sliceEl as PaginatedBodyElement;
        const sliceTp = (sliceEl as unknown as DocTable).tblpPr as TblpPr;
        const { widthPx: tableW, heightPx: sliceH } = retainedTableSliceSize(
          sp, measureState.scale,
        );
        const skipVClamp = sliceTp.vertAnchor === 'page' || sliceTp.vertAnchor === 'margin';
        return computeFloatTableBox(
          sliceTp, measureState, measureState.y, tableW, sliceH, skipVClamp,
          { allowOverlap: tbl.overlap !== 'never' },
        );
      });
    }
  `;
  const expectedSplitCallbackSource = ts.createSourceFile(
    'compute-pages-a5-split-parent-commit-expected.ts',
    `const callback = ${exactSplitParentPreResolutionSource};`,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
  const expectedSplitCallbackDeclaration = expectedSplitCallbackSource.statements[0]
    ?.declarationList?.declarations?.[0];
  const expectedSplitCallback = expectedSplitCallbackDeclaration?.initializer;
  const exactSplitParentPreResolutionSyntax = expectedSplitCallback
    ? printedSyntax(expectedSplitCallback, expectedSplitCallbackSource)
    : '';
  const splitParentPreResolutions = [];
  const findSplitParentPreResolution = (current) => {
    if (ts.isArrowFunction(current)
      && printedSyntax(current, source) === exactSplitParentPreResolutionSyntax) {
      splitParentPreResolutions.push(current);
    }
    ts.forEachChild(current, findSplitParentPreResolution);
  };
  findSplitParentPreResolution(node);
  if (splitParentPreResolutions.length === 1) {
    const callback = splitParentPreResolutions[0];
    const nodeStart = node.getStart(source);
    const start = callback.getStart(source) - nodeStart;
    const end = callback.getEnd() - nodeStart;
    const legacyCallback = `
      (sliceEl) => {
        withColumnBand(() => {
          const sp = sliceEl as PaginatedBodyElement;
          const sliceTp = (sliceEl as unknown as DocTable).tblpPr as TblpPr;
          const { widthPx: tableW, heightPx: sliceH } = retainedTableSliceSize(
            sp, measureState.scale,
          );
          const skipVClamp = sliceTp.vertAnchor === 'page' || sliceTp.vertAnchor === 'margin';
          const sliceBox = computeFloatTableBox(
            sliceTp, measureState, measureState.y, tableW, sliceH, skipVClamp,
          );
          const side = floatTableWrapSide(sliceBox, measureState);
          registerTableFloat(sliceBox, sliceTp, measureState, side, tbl.overlap !== 'never');
        });
        pushTagged(sliceEl);
      }
    `;
    const nodeText = node.getText(source);
    const virtualText = nodeText.slice(0, start) + legacyCallback + nodeText.slice(end);
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-split-parent-commit-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((candidate) => (
      ts.isFunctionDeclaration(candidate) && candidate.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const exactFittingOuterFinalization = 'withColumnBand(()=>{stampTableLayout(elasPaginatedBodyElement,first.layout.colWidths,first.layout.rowHeights,first.contentWPt,);constside=floatTableWrapSide(first.box,measureState);registerTableFloat(first.box,tp,measureState,side,tbl.overlap!==\'never\');});';
  const fittingOuterStatements = [];
  const findFittingOuterStatement = (current) => {
    if (ts.isExpressionStatement(current)
      && compactText(current, source) === exactFittingOuterFinalization) {
      fittingOuterStatements.push(current);
    }
    ts.forEachChild(current, findFittingOuterStatement);
  };
  findFittingOuterStatement(node);
  if (fittingOuterStatements.length === 1) {
    const statement = fittingOuterStatements[0];
    const nodeStart = node.getStart(source);
    const relativeStart = statement.getStart(source) - nodeStart;
    const relativeEnd = statement.getEnd() - nodeStart;
    const nodeText = node.getText(source);
    const legacySequence = [
      'withColumnBand(() => {',
      '  const side = floatTableWrapSide(first.box, measureState);',
      "  registerTableFloat(first.box, tp, measureState, side, tbl.overlap !== 'never');",
      '});',
      'stampTableLayout(',
      '  el as PaginatedBodyElement,',
      '  first.layout.colWidths,',
      '  first.layout.rowHeights,',
      '  first.contentWPt,',
      ');',
    ].join('\n');
    const virtualText = nodeText.slice(0, relativeStart)
      + legacySequence
      + nodeText.slice(relativeEnd);
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-fitting-outer-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((candidate) => (
      ts.isFunctionDeclaration(candidate) && candidate.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const renamedEnvelopeCallees = [];
  const findRenamedEnvelopeCallees = (current) => {
    if (ts.isCallExpression(current)
      && ts.isIdentifier(current.expression)
      && current.expression.text === 'attachTableFragment') {
      renamedEnvelopeCallees.push(current.expression);
    }
    ts.forEachChild(current, findRenamedEnvelopeCallees);
  };
  findRenamedEnvelopeCallees(node);
  if (renamedEnvelopeCallees.length === 2) {
    const nodeStart = node.getStart(source);
    let virtualText = node.getText(source);
    for (const callee of renamedEnvelopeCallees
      .slice()
      .sort((left, right) => right.getStart(source) - left.getStart(source))) {
      const start = callee.getStart(source) - nodeStart;
      const end = callee.getEnd() - nodeStart;
      virtualText = virtualText.slice(0, start)
        + 'attachRetainedTableEnvelope'
        + virtualText.slice(end);
    }
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-envelope-rename-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((statement) => (
      ts.isFunctionDeclaration(statement) && statement.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const exactIntermediateUpright = `
    stampTableLayout(tableEl, colWidthsPt, rowHeightsPt, bandPt);
    if (y + h > effContentH() - tblReservePt && y > colTopY) nextColumnOrPage(i);
    const sourceIndex = bodySourceIndexFor(tbl);
    const retained = sourceIndex === undefined
      ? undefined
      : measureState.retainedTablesBySourceIndex?.get(sourceIndex);
    if (sourceIndex === undefined || retained === undefined) {
      throw new Error('Upright vertical table requires retained physical geometry');
    }
    attachRetainedTablePlacement(tableEl, retained.layout, sourceIndex, {
      xPt: colX(),
      yPt: measureState.y,
      widthPt: colW(),
      flowAdvancePt: h,
      columnIndex: colIndex,
    });
  `.replace(/\s+/g, '');
  const intermediateUprightTransactions = [];
  const findIntermediateUprightTransaction = (current) => {
    if (ts.isBlock(current)) {
      for (let index = 0; index + 5 < current.statements.length; index += 1) {
        const statements = current.statements.slice(index, index + 6);
        if (statements.map((statement) => compactText(statement, source)).join('')
          === exactIntermediateUpright) {
          intermediateUprightTransactions.push([statements[0], statements[5]]);
        }
      }
    }
    ts.forEachChild(current, findIntermediateUprightTransaction);
  };
  findIntermediateUprightTransaction(node);
  if (intermediateUprightTransactions.length === 1) {
    const [startStatement, endStatement] = intermediateUprightTransactions[0];
    const nodeStart = node.getStart(source);
    const relativeStart = startStatement.getStart(source) - nodeStart;
    const relativeEnd = endStatement.getEnd() - nodeStart;
    const canonicalUpright = [
      'stampTableLayout(tableEl, colWidthsPt, rowHeightsPt, bandPt);',
      'if (y + h > effContentH() - tblReservePt && y > colTopY) nextColumnOrPage(i);',
    ].join('\n');
    const nodeText = node.getText(source);
    const virtualText = nodeText.slice(0, relativeStart)
      + canonicalUpright
      + nodeText.slice(relativeEnd);
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-intermediate-upright-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((statement) => (
      ts.isFunctionDeclaration(statement) && statement.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const exactUprightRelocation =
    'if(y+h>effContentH()-tblReservePt&&y>colTopY)nextColumnOrPage(i);';
  const exactUprightStamp = 'withColumnBand(()=>stampTableLayout(tableEl,colWidthsPt,rowHeightsPt,bandPt,{...measureState,pageIndex:pages.length-1,displayPageNumber:pages.length,},));';
  let uprightPair = null;
  const findUprightPair = (current) => {
    if (uprightPair || !ts.isBlock(current)) {
      ts.forEachChild(current, findUprightPair);
      return;
    }
    const statements = current.statements;
    for (let index = 0; index + 1 < statements.length; index += 1) {
      if (compactText(statements[index], source) === exactUprightRelocation
        && compactText(statements[index + 1], source) === exactUprightStamp) {
        uprightPair = [statements[index], statements[index + 1]];
        return;
      }
    }
    ts.forEachChild(current, findUprightPair);
  };
  findUprightPair(node);
  if (uprightPair) {
    const [relocation, stamp] = uprightPair;
    const nodeStart = node.getStart(source);
    const relativeStart = relocation.getStart(source) - nodeStart;
    const relativeEnd = stamp.getEnd() - nodeStart;
    const nodeText = node.getText(source);
    const legacySequence = [
      'stampTableLayout(tableEl, colWidthsPt, rowHeightsPt, bandPt);',
      'if (y + h > effContentH() - tblReservePt && y > colTopY) nextColumnOrPage(i);',
    ].join('\n');
    const virtualText = nodeText.slice(0, relativeStart)
      + legacySequence
      + nodeText.slice(relativeEnd);
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-upright-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((statement) => (
      ts.isFunctionDeclaration(statement) && statement.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const exactRetainedSliceSize =
    'const{widthPx:tableW,heightPx:sliceH}=retainedTableSliceSize(sp,measureState.scale,);';
  const retainedSliceStatements = [];
  const findRetainedSliceStatement = (current) => {
    if (ts.isVariableStatement(current)
      && compactText(current, source) === exactRetainedSliceSize) {
      retainedSliceStatements.push(current);
    }
    ts.forEachChild(current, findRetainedSliceStatement);
  };
  findRetainedSliceStatement(node);
  if (retainedSliceStatements.length === 1) {
    const statement = retainedSliceStatements[0];
    const nodeStart = node.getStart(source);
    const relativeStart = statement.getStart(source) - nodeStart;
    const relativeEnd = statement.getEnd() - nodeStart;
    const nodeText = node.getText(source);
    const legacyFold = [
      'const tableW = (sp.tableColWidthsPt ?? []).reduce((s, w) => s + w, 0) * measureState.scale;',
      'const sliceH = (sp.tableRowHeightsPt ?? []).reduce((s, h) => s + h, 0) * measureState.scale;',
    ].join('\n');
    const virtualText = nodeText.slice(0, relativeStart)
      + legacyFold
      + nodeText.slice(relativeEnd);
    const virtualSource = ts.createSourceFile(
      'compute-pages-a5-virtual.ts',
      virtualText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    const virtualNode = virtualSource.statements.find((statement) => (
      ts.isFunctionDeclaration(statement) && statement.name?.text === 'computePages'
    ));
    if (virtualNode) return normalizedComputePagesHash(virtualNode, virtualSource);
  }
  const allowedNames = ['layoutServices', 'layoutOptions'];
  const allowedParameterSyntax = [
    'layoutServices?: LayoutServices',
    'layoutOptions?: LayoutOptions',
  ];
  const appendedParameters = node.parameters?.slice(-2) ?? [];
  const hasAllowedParameters = appendedParameters.length === 2
    && appendedParameters.every((parameter, index) => (
      ts.isIdentifier(parameter.name) && parameter.name.text === allowedNames[index]
      && parameter.getText(source).replace(/\s+/g, ' ').trim() === allowedParameterSyntax[index]
    ));
  const omittedParameters = new Set(hasAllowedParameters ? appendedParameters : []);
  const shape = (current) => {
    if (omittedParameters.has(current)) return null;
    if (ts.isCallExpression(current)
      && ts.isIdentifier(current.expression)
      && current.expression.text === 'buildMeasureState') {
      const tail = current.arguments.slice(-2);
      const hasAllowedArguments = tail.length === 2
        && tail.every((argument, index) => ts.isIdentifier(argument) && argument.text === allowedNames[index]);
      const args = hasAllowedArguments ? current.arguments.slice(0, -2) : current.arguments;
      return [
        ts.SyntaxKind[current.kind],
        undefined,
        shape(current.expression),
        ...args.map(shape),
      ];
    }
    const text = ts.isIdentifier(current) || ts.isLiteralExpression(current)
      ? current.getText(source)
      : undefined;
    const children = [];
    current.forEachChild((child) => {
      const childShape = shape(child);
      if (childShape !== null) children.push(childShape);
    });
    return [ts.SyntaxKind[current.kind], text, ...children];
  };
  return createHash('sha256').update(JSON.stringify(shape(node))).digest('hex');
}

/** A5 removes the body-table stamp/reuse prefix while leaving the B1
 * header/footer fallback byte-identical. Normalize only that complete replacement. */
function normalizedComputeTableLayoutHash(node, source) {
  if (!ts.isFunctionDeclaration(node) || !node.body) return normalizedSyntaxHash(node, source);
  const compact = (value) => value.replace(/\s+/g, '');
  const statements = node.body.statements;
  const b1Index = statements.findIndex((statement) => (
    ts.isVariableStatement(statement)
    && compact(statement.getText(source)).startsWith(
      'constcolWidths=resolveColumnWidths(table,contentWPt1,state).map(',
    )
  ));
  if (b1Index < 0) return normalizedSyntaxHash(node, source);
  const retainedPrefix = compact(statements.slice(0, b1Index)
    .map((statement) => statement.getText(source)).join(''));
  const sourceKeyedRetainedPrefix = compact(`
    const { scale } = state;
    const contentWPt1 = contentWPx / scale;
    if (state.retainedTableAcquisition && bodySourceIndexFor(table) !== undefined) {
      const retained = computeTablePtLayout(state, table, contentWPt1);
      const colWidths = retained.colWidthsPt.map((width) => width * scale);
      const rowHeights = retained.rowHeightsPt.map((height) => height * scale);
      return {
        colWidths,
        tableW: colWidths.reduce((sum, width) => sum + width, 0),
        rowContentHeights: retained.rowContentHeightsPt.map((height) => height * scale),
        rowHeights,
      };
    }
  `);
  const occurrenceOwnedRetainedPrefix = compact(`
    const { scale } = state;
    const contentWPt1 = contentWPx / scale;
    if (state.retainedTableAcquisition && sourceIndex !== undefined) {
      const retained = computeTablePtLayout(state, table, contentWPt1, sourceIndex);
      const colWidths = retained.colWidthsPt.map((width) => width * scale);
      const rowHeights = retained.rowHeightsPt.map((height) => height * scale);
      return {
        colWidths,
        tableW: colWidths.reduce((sum, width) => sum + width, 0),
        rowContentHeights: retained.rowContentHeightsPt.map((height) => height * scale),
        rowHeights,
      };
    }
  `);
  const occurrenceOwned = retainedPrefix === occurrenceOwnedRetainedPrefix;
  if (retainedPrefix !== sourceKeyedRetainedPrefix && !occurrenceOwned) {
    if (retainedPrefix.includes('computeTablePtLayout')) {
      return createHash('sha256')
        .update(`invalid-a5-retained-prefix:${retainedPrefix}`)
        .digest('hex');
    }
    return normalizedSyntaxHash(node, source);
  }
  const legacyPrefix = `
    const { scale } = state;
    const contentHeightsFromResolved = (rowHeights: number[]): number[] => {
      const footprints = applyTableRowBoundaryFootprints(
        table,
        new Array<number>(table.rows.length).fill(0),
        scale,
      );
      return rowHeights.map((height, index) => height - (footprints[index] ?? 0));
    };
    const stamped = table as PaginatedBodyElement;
    const contentWPt1 = contentWPx / scale;
    const placedFragment = bodyFlowFragments.get(table as object);
    const fragmentBandPt = tableFragmentBandPt.get(table as object);
    if (
      tableReuseEnabled &&
      placedFragment !== undefined &&
      placedFragment.fragment.kind === 'table' &&
      fragmentBandPt !== undefined &&
      placedFragment.fragment.rows.length === table.rows.length &&
      Math.abs(fragmentBandPt - contentWPt1) <= 1e-6 * Math.max(1, Math.abs(contentWPt1))
    ) {
      const fragment = placedFragment.fragment;
      const colWidths = fragment.columnWidthsPt.map((w) => w * scale);
      const rowHeights = fragment.rows.map((r) => r.heightPt * scale);
      return {
        colWidths,
        tableW: colWidths.reduce((s, w) => s + w, 0),
        rowContentHeights: contentHeightsFromResolved(rowHeights),
        rowHeights,
      };
    }
    const reuseInputs = stamped.tableLayoutInputs;
    const reuse =
      tableReuseEnabled &&
      reuseInputs !== undefined &&
      stamped.tableColWidthsPt !== undefined &&
      stamped.tableRowHeightsPt !== undefined &&
      reuseInputs.scale === 1 &&
      stamped.tableRowHeightsPt.length === table.rows.length &&
      Math.abs(reuseInputs.contentWPt - contentWPt1) <= 1e-6 * Math.max(1, Math.abs(contentWPt1));
    if (reuse) {
      const colWidths = (stamped.tableColWidthsPt as number[]).map((w) => w * scale);
      const rowHeights = (stamped.tableRowHeightsPt as number[]).map((h) => h * scale);
      return {
        colWidths,
        tableW: colWidths.reduce((s, w) => s + w, 0),
        rowContentHeights: contentHeightsFromResolved(rowHeights),
        rowHeights,
      };
    }
  `;
  const first = statements[0];
  const b1 = statements[b1Index];
  if (!first || !b1) return normalizedSyntaxHash(node, source);
  const nodeStart = node.getStart(source);
  const nodeText = node.getText(source);
  let virtualText = nodeText.slice(0, first.getStart(source) - nodeStart)
    + legacyPrefix
    + nodeText.slice(b1.getStart(source) - nodeStart);
  if (occurrenceOwned) {
    const exactParameter = ',\n  sourceIndex?: number,\n): {';
    if (virtualText.split(exactParameter).length !== 2) return normalizedSyntaxHash(node, source);
    virtualText = virtualText.replace(exactParameter, '\n): {');
  }
  const virtualSource = ts.createSourceFile(
    'compute-table-layout-a5-virtual.ts',
    virtualText,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
  const virtualNode = virtualSource.statements.find((statement) => (
    ts.isFunctionDeclaration(statement) && statement.name?.text === 'computeTableLayout'
  ));
  return virtualNode
    ? normalizedSyntaxHash(virtualNode, virtualSource)
    : normalizedSyntaxHash(node, source);
}

/** A3 deletes the production-wide fragment flag after retained table paint is
 * mandatory. Normalize only its exact first `if` conjunct; all table eligibility
 * predicates remain hash-frozen through A5. */
function normalizedIsFragmentPaintableTableHash(node, source) {
  const firstIf = node.body?.statements.find(ts.isIfStatement);
  const targetCondition = firstIf?.expression;
  const flattenOr = (current, operands = []) => {
    if (ts.isBinaryExpression(current)
      && current.operatorToken.kind === ts.SyntaxKind.BarBarToken) {
      flattenOr(current.left, operands);
      flattenOr(current.right, operands);
    } else {
      operands.push(current);
    }
    return operands;
  };
  const exactDeletedGate = (current) => ts.isPrefixUnaryExpression(current)
    && current.operator === ts.SyntaxKind.ExclamationToken
    && ts.isIdentifier(current.operand)
    && current.operand.text === 'fragmentPaintEnabled';
  let normalized = node.getText(source);
  if (targetCondition) {
    const operands = flattenOr(targetCondition);
    if (operands.length >= 2 && exactDeletedGate(operands[0])) {
      const start = operands[0].getStart(source) - node.getStart(source);
      const end = operands[1].getStart(source) - node.getStart(source);
      const exactPrefix = normalized.slice(start, end).replace(/\s+/g, '');
      if (exactPrefix === '!fragmentPaintEnabled||') {
        normalized = normalized.slice(0, start) + normalized.slice(end);
      }
    }
  }
  return createHash('sha256').update(normalized.replace(/\s+/g, ' ').trim()).digest('hex');
}

/** A2 routes service-produced shape text through the same immutable Canvas
 * route used by measurement. The text-box implementation remains hash frozen
 * except for exact Canvas-route threading and the spec-required numbering
 * marker snapshot -> shape -> retained-paint sequence below. */
function normalizedRenderShapeTextHash(node, source) {
  const omittedRouteParameters = new Set();
  const findAllowedRouteParameters = (current) => {
    if (ts.isVariableDeclaration(current)
      && ts.isIdentifier(current.name)
      && current.name.text === 'shapeLineMetrics'
      && current.initializer
      && ts.isArrowFunction(current.initializer)) {
      const tail = current.initializer.parameters.slice(-2);
      const exact = ['familyRoute?: CanvasFontRoute', 'familyEaRoute?: CanvasFontRoute'];
      if (tail.length === 2 && tail.every((parameter, index) => (
        parameter.getText(source).replace(/\s+/g, ' ').trim() === exact[index]
      ))) {
        tail.forEach((parameter) => omittedRouteParameters.add(parameter));
      }
    }
    ts.forEachChild(current, findAllowedRouteParameters);
  };
  findAllowedRouteParameters(node);
  const compactText = (current, currentSource) =>
    current.getText(currentSource).replace(/\s+/g, '');
  const exactMarkerInput =
    'constmarkerShapeInput=numberingMarkerShapeInput(block.numbering,block.fontSizePt);';
  const exactMarkerLayout =
    'constmarkerTextLayout=shapeNumberingMarkerText(markerShapeInput,markerText,scale,effState.layoutServices?.text,);';
  const exactMarkerWidth =
    'constmarkerW=markerTextLayout?.shape.advancePt??ctx.measureText(markerText).width;';
  const exactMarkerPaint = [
    'if(markerTextLayout){',
    'paintNumberingMarkerText(ctx,markerTextLayout,markerX,baseline,',
    'eaVertUpright?(paintCtx,text,drawX,drawBaseline,fontSizePx)=>{',
    'drawVerticalRun(paintCtx,text,drawX,drawBaseline,fontSizePx,0);',
    '}:undefined,);',
    '}elseif(eaVertUpright){',
    'drawVerticalRun(ctx,markerText,markerX,baseline,block.fontSizePt*scale,0);',
    '}else{ctx.fillText(markerText,markerX,baseline);}',
  ].join('');
  const markerMigrationCounts = [0, 0, 0, 0];
  const countMarkerMigration = (current) => {
    const compact = compactText(current, source);
    if (ts.isVariableStatement(current) && compact === exactMarkerInput) markerMigrationCounts[0] += 1;
    if (ts.isVariableStatement(current) && compact === exactMarkerLayout) markerMigrationCounts[1] += 1;
    if (ts.isVariableStatement(current) && compact === exactMarkerWidth) markerMigrationCounts[2] += 1;
    if (ts.isIfStatement(current) && compact === exactMarkerPaint) markerMigrationCounts[3] += 1;
    ts.forEachChild(current, countMarkerMigration);
  };
  countMarkerMigration(node);
  const exactMarkerMigration = markerMigrationCounts.every((count) => count === 1);
  let shape;
  const replacementShape = (text) => {
    const replacementSource = ts.createSourceFile(
      'render-shape-text-a2-replacement.ts',
      text,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    );
    return shape(replacementSource.statements[0], replacementSource);
  };
  shape = (current, currentSource = source) => {
    if (ts.isVariableStatement(current)) {
      const compact = compactText(current, currentSource);
      if (exactMarkerMigration && (compact === exactMarkerInput || compact === exactMarkerLayout)) {
        return null;
      }
      if (exactMarkerMigration && compact === exactMarkerWidth) {
        return replacementShape('const markerW = ctx.measureText(markerText).width;');
      }
    }
    if (exactMarkerMigration
      && ts.isIfStatement(current)
      && compactText(current, currentSource) === exactMarkerPaint) {
      return replacementShape(
        'if (eaVertUpright) { drawVerticalRun(ctx, markerText, markerX, baseline, block.fontSizePt * scale, 0); } else { ctx.fillText(markerText, markerX, baseline); }',
      );
    }
    // A partial/duplicated marker migration is intentionally left in the AST
    // hash. Only the complete four-node contract above can normalize away.
    if (ts.isVariableStatement(current)
      && current.declarationList.declarations.length === 1) {
      const [declaration] = current.declarationList.declarations;
      if (ts.isIdentifier(declaration.name)
        && declaration.name.text === 'measureRoute'
        && current.getText(currentSource).replace(/\s+/g, ' ').trim()
          === 'const measureRoute = eaIntended > asciiIntended ? familyEaRoute : familyRoute;') {
        return null;
      }
    }
    if (omittedRouteParameters.has(current)) return null;
    if (ts.isCallExpression(current)
      && ts.isIdentifier(current.expression)
      && current.expression.text === 'buildFont') {
      const route = current.arguments.at(-1);
      const hasAllowedRoute = current.arguments.length === 6
        && route != null
        && ((ts.isPropertyAccessExpression(route)
          && ts.isIdentifier(route.expression)
          && route.expression.text === 's'
          && route.name.text === 'fontRoute')
          || (ts.isIdentifier(route) && route.text === 'measureRoute'));
      const args = hasAllowedRoute ? current.arguments.slice(0, -1) : current.arguments;
      return [
        ts.SyntaxKind[current.kind],
        undefined,
        shape(current.expression, currentSource),
        ...args.map((argument) => shape(argument, currentSource)),
      ];
    }
    if (ts.isCallExpression(current)
      && ts.isIdentifier(current.expression)
      && current.expression.text === 'shapeLineMetrics') {
      const [fontRoute, eaFloorRoute] = current.arguments.slice(-2);
      const fontRouteText = fontRoute?.getText(source);
      const eaFloorRouteText = eaFloorRoute?.getText(source);
      const hasAllowedRoutes = current.arguments.length >= 3
        && (fontRouteText === 's.fontRoute' || fontRouteText === 'tallest?.fontRoute')
        && (eaFloorRouteText === 's.eaFloorRoute' || eaFloorRouteText === 'tallest?.eaFloorRoute');
      const args = hasAllowedRoutes ? current.arguments.slice(0, -2) : current.arguments;
      return [
        ts.SyntaxKind[current.kind],
        undefined,
        shape(current.expression, currentSource),
        ...args.map((argument) => shape(argument, currentSource)),
      ];
    }
    const text = ts.isIdentifier(current) || ts.isLiteralExpression(current)
      ? current.getText(currentSource)
      : undefined;
    const children = [];
    current.forEachChild((child) => {
      const childShape = shape(child, currentSource);
      if (childShape !== null) children.push(childShape);
    });
    return [ts.SyntaxKind[current.kind], text, ...children];
  };
  return createHash('sha256').update(JSON.stringify(shape(node))).digest('hex');
}

function declarationInventory(root, allowTransitionalAdapter = false) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  const nonLayoutDeclarationKeys = [];
  const legacyDeclarationHashes = {};
  for (const path of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    const file = posixPath(relative(root, path));
    const migrationOwner = file.startsWith(`${LAYOUT_SOURCE}/`)
      || file.startsWith(`${PAINT_SOURCE}/`)
      || file.startsWith(`${DOCX_SOURCE}/conformance/`)
      || PLANNED_NON_LAYOUT_MODULES.has(file);
    const source = sourceFile(path);
    for (const statement of source.statements) {
      for (const name of declarationNames(statement)) {
        const key = `${file}#${declarationKind(statement)}#${name}`;
        // The final renderer surface is fixed up front, so staged PRs may add
        // those named adapters without opening a route for arbitrary helpers.
        const plannedRendererAdapter = file === `${DOCX_SOURCE}/renderer.ts`
          && (FINAL_RENDERER_DECLARATIONS.has(name) || A5_STATE_OWNER_DECLARATIONS.has(name));
        const plannedBodyLayoutAdapter = file === BODY_LAYOUT_ADAPTER
          && BODY_LAYOUT_ADAPTER_DECLARATIONS.has(name);
        const transitionalParagraphAnchorAdapter = allowTransitionalAdapter
          && file === PARAGRAPH_ANCHOR_FRAME_ADAPTER
          && PARAGRAPH_ANCHOR_FRAME_ADAPTER_DECLARATIONS.has(name);
        const plannedWorkerLayoutRetention = file === WORKER_LAYOUT_RETENTION
          && WORKER_LAYOUT_RETENTION_DECLARATIONS.has(name);
        const plannedTextRunProjectionAdapter = file === TEXT_RUN_PROJECTION_ADAPTER
          && TEXT_RUN_PROJECTION_ADAPTER_DECLARATIONS.has(name);
        const plannedBodyKernelImplementation = file === `${DOCX_SOURCE}/renderer.ts`
          && BODY_KERNEL_IMPLEMENTATION_DECLARATIONS.has(name);
        if (!migrationOwner && !plannedRendererAdapter && !plannedBodyLayoutAdapter
          && !transitionalParagraphAnchorAdapter && !plannedWorkerLayoutRetention
          && !plannedTextRunProjectionAdapter && !plannedBodyKernelImplementation) {
          nonLayoutDeclarationKeys.push(key);
        }
        if (LEGACY_SYMBOLS.includes(name)) {
          legacyDeclarationHashes[key] = name === 'computePages'
            ? normalizedComputePagesHash(statement, source)
            : name === 'computeTableLayout'
              ? normalizedComputeTableLayoutHash(statement, source)
            : name === 'isFragmentPaintableTable'
              ? normalizedIsFragmentPaintableTableHash(statement, source)
            : name === 'renderShapeText'
              ? normalizedRenderShapeTextHash(statement, source)
            : normalizedNodeHash(statement, source);
        }
      }
    }
  }
  return {
    nonLayoutDeclarationKeys: [...new Set(nonLayoutDeclarationKeys)].sort(),
    legacyDeclarationHashes: Object.fromEntries(
      Object.entries(legacyDeclarationHashes).sort(([left], [right]) => left.localeCompare(right)),
    ),
  };
}

function rendererImportEdges(root) {
  const renderer = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(renderer)) return [];
  return [...new Set(moduleEdges(renderer)
    .filter((edge) => edge.literal)
    .map((edge) => resolveLocalImport(renderer, edge.specifier))
    .filter((path) => path && LEGACY_RENDERER_IMPORTS.has(basename(path)))
    .map((path) => `${DOCX_SOURCE}/renderer.ts -> ${posixPath(relative(root, path))}`))]
    .sort();
}

function currentAllowances(root, allowTransitionalAdapter = false) {
  const declarations = declarationInventory(root, allowTransitionalAdapter);
  return {
    version: 2,
    legacySymbolCounts: identifierCounts(root),
    migrationIdentifierCounts: matchingIdentifierCounts(root, (name) => MIGRATION_IDENTIFIER.test(name)),
    nonLayoutDeclarationKeys: declarations.nonLayoutDeclarationKeys,
    legacyDeclarationHashes: declarations.legacyDeclarationHashes,
    rendererImportEdges: rendererImportEdges(root),
  };
}

function parseBaselineJson(contents, detail) {
  try {
    return JSON.parse(contents);
  } catch {
    fail('INVALID_BASELINE', detail);
  }
}

function assertUniqueJsonObjectKeys(contents, detail) {
  const source = ts.parseJsonText(detail, contents);
  const visit = (node) => {
    if (ts.isObjectLiteralExpression(node)) {
      const keys = new Set();
      for (const property of node.properties) {
        const name = property.name;
        const key = name && (ts.isStringLiteralLike(name) || ts.isIdentifier(name))
          ? name.text
          : null;
        if (key !== null && keys.has(key)) {
          fail('INVALID_BASELINE', `${detail}: duplicate key ${key}`);
        }
        if (key !== null) keys.add(key);
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
}

function assertValidBaseline(value, detail) {
  if (value === null
    || typeof value !== 'object'
    || value.version !== 2
    || typeof value.legacySymbolCounts !== 'object'
    || value.legacySymbolCounts === null
    || Array.isArray(value.legacySymbolCounts)
    || typeof value.migrationIdentifierCounts !== 'object'
    || value.migrationIdentifierCounts === null
    || Array.isArray(value.migrationIdentifierCounts)
    || !Array.isArray(value.nonLayoutDeclarationKeys)
    || typeof value.legacyDeclarationHashes !== 'object'
    || value.legacyDeclarationHashes === null
    || Array.isArray(value.legacyDeclarationHashes)
    || !Array.isArray(value.rendererImportEdges)) {
    fail('INVALID_BASELINE', detail);
  }
  return value;
}

function readBaseline(path) {
  const contents = readFileSync(path, 'utf8');
  assertUniqueJsonObjectKeys(contents, path);
  return assertValidBaseline(parseBaselineJson(contents, path), path);
}

function git(root, args, allowFailure = false) {
  const result = spawnSync('git', args, { cwd: root, encoding: 'utf8' });
  if (result.status !== 0 && !allowFailure) fail('GIT_ERROR', `${args.join(' ')}: ${result.stderr.trim()}`);
  return result;
}

function mergeBaseBaseline(root, baseRef) {
  const mergeBase = git(root, ['merge-base', baseRef, 'HEAD']).stdout.trim();
  const shown = git(root, ['show', `${mergeBase}:${BASELINE_PATH}`], true);
  if (shown.status !== 0) return null;
  const detail = `${mergeBase}:${BASELINE_PATH}`;
  const value = assertValidBaseline(parseBaselineJson(shown.stdout, detail), detail);
  // The stored A1 hashes predate the A2-specific normalization. Recompute only
  // the two mechanically constrained declarations from the immutable merge-base
  // source; every other declaration continues to use the committed baseline.
  const renderer = git(root, ['show', `${mergeBase}:${DOCX_SOURCE}/renderer.ts`], true);
  if (renderer.status === 0) {
    const source = ts.createSourceFile('renderer.ts', renderer.stdout, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
    const declaration = (name) => source.statements.find((statement) => (
      ts.isFunctionDeclaration(statement) && statement.name?.text === name
    ));
    const computePages = declaration('computePages');
    if (computePages) {
      value.legacyDeclarationHashes[`${DOCX_SOURCE}/renderer.ts#FunctionDeclaration#computePages`]
        = normalizedComputePagesHash(computePages, source);
    }
    const computeTableLayout = declaration('computeTableLayout');
    if (computeTableLayout) {
      value.legacyDeclarationHashes[`${DOCX_SOURCE}/renderer.ts#FunctionDeclaration#computeTableLayout`]
        = normalizedComputeTableLayoutHash(computeTableLayout, source);
    }
    const renderShapeText = declaration('renderShapeText');
    if (renderShapeText) {
      value.legacyDeclarationHashes[`${DOCX_SOURCE}/renderer.ts#FunctionDeclaration#renderShapeText`]
        = normalizedRenderShapeTextHash(renderShapeText, source);
    }
    const isFragmentPaintableTable = declaration('isFragmentPaintableTable');
    if (isFragmentPaintableTable) {
      value.legacyDeclarationHashes[`${DOCX_SOURCE}/renderer.ts#FunctionDeclaration#isFragmentPaintableTable`]
        = normalizedIsFragmentPaintableTableHash(isFragmentPaintableTable, source);
    }
  }
  return value;
}

function assertNoExpansion(head, base) {
  for (const [symbol, count] of Object.entries(head.legacySymbolCounts)) {
    const baseCount = base.legacySymbolCounts[symbol] ?? 0;
    if (count > baseCount) fail('BASELINE_EXPANSION', `${symbol}: ${count} > ${baseCount}`);
  }
  for (const [identifier, count] of Object.entries(head.migrationIdentifierCounts)) {
    const baseCount = base.migrationIdentifierCounts[identifier] ?? 0;
    if (count > baseCount) fail('BASELINE_EXPANSION', `${identifier}: ${count} > ${baseCount}`);
  }
  const baseDeclarations = new Set(base.nonLayoutDeclarationKeys);
  for (const declaration of head.nonLayoutDeclarationKeys) {
    if (!baseDeclarations.has(declaration)) fail('BASELINE_EXPANSION', declaration);
  }
  for (const [declaration, hash] of Object.entries(head.legacyDeclarationHashes)) {
    const baseHash = base.legacyDeclarationHashes[declaration];
    if (!baseHash) fail('BASELINE_EXPANSION', declaration);
    if (hash !== baseHash) fail('LEGACY_DECLARATION_CHANGED', declaration);
  }
  const baseEdges = new Set(base.rendererImportEdges);
  for (const edge of head.rendererImportEdges) {
    if (!baseEdges.has(edge)) fail('BASELINE_EXPANSION', edge);
  }
}

function stableJson(value) {
  return `${JSON.stringify(value, null, 2)}\n`;
}

function assertExactBaseline(baseline, actual) {
  const expected = stableJson(baseline);
  const received = stableJson(actual);
  if (expected !== received) {
    const expectedLines = expected.split('\n');
    const receivedLines = received.split('\n');
    const index = expectedLines.findIndex((line, lineIndex) => line !== receivedLines[lineIndex]);
    fail(
      'BASELINE_MISMATCH',
      `baseline must exactly describe current legacy symbols and renderer import edges; first difference at line ${index + 1}: expected ${JSON.stringify(expectedLines[index])}, received ${JSON.stringify(receivedLines[index])}`,
    );
  }
}

function assertPaintBoundaries(root) {
  const { violations, nonLiteral } = paintBoundaryViolations(root);
  if (nonLiteral.length > 0) {
    fail('NON_LITERAL_MODULE_EDGE', nonLiteral.join('\n'));
  }
  if (violations.length > 0) {
    fail('FORBIDDEN_PAINT_EDGE', violations.map((chain) => chain.join(' -> ')).join('\n'));
  }
}

function hasExportModifier(statement) {
  return statement.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword) ?? false;
}

function rendererExports(path) {
  const names = [];
  const source = sourceFile(path);
  for (const statement of source.statements) {
    if (hasExportModifier(statement)) {
      const declared = declarationNames(statement);
      names.push(...(declared.length > 0 ? declared : ['default']));
    }
    if (ts.isExportDeclaration(statement) && statement.exportClause && ts.isNamedExports(statement.exportClause)) {
      names.push(...statement.exportClause.elements.map((element) => element.name.text));
    }
  }
  return [...new Set(names)].sort();
}

function rendererImportBindings(source) {
  const bindings = new Set();
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement) || !statement.importClause) continue;
    if (statement.importClause.name) bindings.add(statement.importClause.name.text);
    const named = statement.importClause.namedBindings;
    if (named && ts.isNamespaceImport(named)) bindings.add(named.name.text);
    if (named && ts.isNamedImports(named)) {
      for (const element of named.elements) bindings.add(element.name.text);
    }
  }
  return bindings;
}

function unwrapAdapterExpression(expression) {
  if (ts.isAwaitExpression(expression)
    || ts.isParenthesizedExpression(expression)
    || ts.isAsExpression(expression)
    || ts.isSatisfiesExpression(expression)) {
    return unwrapAdapterExpression(expression.expression);
  }
  return expression;
}

function adapterValueIsAllowed(expression, callable) {
  const value = unwrapAdapterExpression(expression);
  if (ts.isIdentifier(value)
    || ts.isStringLiteralLike(value)
    || ts.isNumericLiteral(value)
    || value.kind === ts.SyntaxKind.TrueKeyword
    || value.kind === ts.SyntaxKind.FalseKeyword
    || value.kind === ts.SyntaxKind.NullKeyword) return true;
  if (ts.isPropertyAccessExpression(value)) return adapterValueIsAllowed(value.expression, callable);
  if (ts.isElementAccessExpression(value)) {
    return adapterValueIsAllowed(value.expression, callable)
      && (!value.argumentExpression || adapterValueIsAllowed(value.argumentExpression, callable));
  }
  if (ts.isArrayLiteralExpression(value)) {
    return value.elements.every((element) => !ts.isSpreadElement(element)
      ? adapterValueIsAllowed(element, callable)
      : adapterValueIsAllowed(element.expression, callable));
  }
  if (ts.isObjectLiteralExpression(value)) {
    return value.properties.every((property) => {
      if (ts.isPropertyAssignment(property)) return adapterValueIsAllowed(property.initializer, callable);
      if (ts.isShorthandPropertyAssignment(property)) return true;
      if (ts.isSpreadAssignment(property)) return adapterValueIsAllowed(property.expression, callable);
      return false;
    });
  }
  if (ts.isCallExpression(value)) {
    const callee = unwrapAdapterExpression(value.expression);
    const callableName = ts.isIdentifier(callee)
      ? callee.text
      : ts.isPropertyAccessExpression(callee) && ts.isIdentifier(callee.expression)
        ? callee.expression.text
        : null;
    return callableName != null
      && callable.has(callableName)
      && value.arguments.every((argument) => adapterValueIsAllowed(argument, callable));
  }
  return false;
}

function adapterBodyIsAllowed(body, callable) {
  return body.statements.every((statement) => {
    if (ts.isVariableStatement(statement)) {
      return statement.declarationList.declarations.every((declaration) => (
        ts.isIdentifier(declaration.name)
        && (!declaration.initializer || adapterValueIsAllowed(declaration.initializer, callable))
      ));
    }
    if (ts.isExpressionStatement(statement)) return adapterValueIsAllowed(statement.expression, callable);
    if (ts.isReturnStatement(statement)) {
      return !statement.expression || adapterValueIsAllowed(statement.expression, callable);
    }
    return false;
  });
}

function assertFinalRendererAdapter(root) {
  const renderer = resolve(root, DOCX_SOURCE, 'renderer.ts');
  if (!existsSync(renderer)) fail('FINAL_ADAPTER_MISSING', `${DOCX_SOURCE}/renderer.ts`);
  const source = sourceFile(renderer);
  for (const statement of source.statements) {
    if (ts.isExportDeclaration(statement)
      && (!statement.exportClause || !ts.isNamedExports(statement.exportClause))) {
      fail('FINAL_ADAPTER_EXPORT', statement.getText(source));
    }
    for (const name of declarationNames(statement)) {
      if (!FINAL_RENDERER_DECLARATIONS.has(name)) fail('FINAL_ADAPTER_DECLARATION', name);
    }
  }
  for (const name of rendererExports(renderer)) {
    if (!FINAL_RENDERER_EXPORTS.has(name)) fail('FINAL_ADAPTER_EXPORT', name);
  }
  const callable = rendererImportBindings(source);
  callable.add('createLayoutServices');
  callable.add('normalizeRenderOptions');
  for (const statement of source.statements) {
    if (ts.isFunctionDeclaration(statement)
      && statement.name
      && statement.name.text === 'renderDocumentToCanvas'
      && statement.body
      && !adapterBodyIsAllowed(statement.body, callable)) {
      fail('FINAL_ADAPTER_BODY', statement.name.text);
    }
  }
  for (const edge of moduleEdges(renderer)) {
    if (!edge.literal) fail('FINAL_ADAPTER_IMPORT', '<dynamic>');
    if (edge.bare) fail('FINAL_ADAPTER_IMPORT', edge.specifier);
    if (!edge.specifier.startsWith('.')) continue;
    const target = resolveLocalImport(renderer, edge.specifier);
    if (!target) fail('FINAL_ADAPTER_IMPORT', edge.specifier);
    const rel = posixPath(relative(root, target));
    const allowed = rel.startsWith(`${LAYOUT_SOURCE}/`)
      || rel.startsWith(`${PAINT_SOURCE}/`)
      || rel === TEXT_RUN_PROJECTION_ADAPTER
      || (edge.typeOnly && rel === `${DOCX_SOURCE}/types.ts`);
    if (!allowed) fail('FINAL_ADAPTER_IMPORT', `${DOCX_SOURCE}/renderer.ts -> ${rel}`);
  }
}

function parseArguments(argv) {
  const options = {
    root: process.cwd(),
    baseRef: 'origin/main',
    write: false,
    final: false,
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--root') options.root = resolve(argv[++index]);
    else if (arg === '--base-ref') options.baseRef = argv[++index];
    else if (arg === '--write-transitional-baseline') options.write = true;
    else if (arg === '--final') options.final = true;
    else fail('UNKNOWN_ARGUMENT', arg);
  }
  return options;
}

export function checkDocxLayoutBoundaries(options) {
  const root = resolve(options.root);
  const baselinePath = resolve(root, BASELINE_PATH);
  const baselineExists = existsSync(baselinePath);
  const allowTransitionalParagraphAnchorAdapter = baselineExists && !options.final;
  assertNoProductionTestSupportImports(root);
  assertNoDeletedPageProducerIdentifiers(root);
  assertPaintBoundaries(root);
  assertCapabilityBoundaries(root);
  assertAffineRuntimeDependencies(root);
  assertTextRunProjectionAdapterBoundary(root);
  assertCoordinateSpaceRuntimeDependencies(root);
  assertOccurrenceProjectionRuntimeDependencies(root);
  assertBodyPaintConsumesRetainedLayout(root);
  assertLayoutParserModelBoundaries(root);
  assertBodyLayoutAdapterBoundary(root);
  assertParagraphAnchorFrameAdapterBoundary(root, allowTransitionalParagraphAnchorAdapter);
  assertBodyKernelServiceOwner(root);
  assertCanonicalCutoverBoundaries(root);

  if (options.write) {
    const baseBaseline = mergeBaseBaseline(root, options.baseRef);
    if (baseBaseline) fail('TRANSITIONAL_BASELINE_EXISTS', `${options.baseRef} already contains ${BASELINE_PATH}`);
    const actual = currentAllowances(root);
    mkdirSync(dirname(baselinePath), { recursive: true });
    writeFileSync(baselinePath, stableJson(actual));
    return;
  }

  if (options.final || !baselineExists) {
    if (options.final && baselineExists) fail('FINAL_BASELINE_PRESENT', BASELINE_PATH);
    assertFinalRendererAdapter(root);
    const actual = currentAllowances(root);
    if (Object.keys(actual.legacySymbolCounts).length > 0
      || Object.keys(actual.migrationIdentifierCounts).length > 0
      || Object.keys(actual.legacyDeclarationHashes).length > 0
      || actual.rendererImportEdges.length > 0) {
      fail('FINAL_LEGACY_BOUNDARY', stableJson(actual).trim());
    }
    return;
  }

  const baseBaseline = mergeBaseBaseline(root, options.baseRef);
  const headBaseline = readBaseline(baselinePath);
  const normalizedDeclarationKeys = [
    `${DOCX_SOURCE}/renderer.ts#FunctionDeclaration#computePages`,
    `${DOCX_SOURCE}/renderer.ts#FunctionDeclaration#computeTableLayout`,
    `${DOCX_SOURCE}/renderer.ts#FunctionDeclaration#isFragmentPaintableTable`,
    `${DOCX_SOURCE}/renderer.ts#FunctionDeclaration#renderShapeText`,
  ];
  for (const key of normalizedDeclarationKeys) {
    if (headBaseline.legacyDeclarationHashes[key]
      && baseBaseline?.legacyDeclarationHashes[key]) {
      // Treat the immutable merge-base declaration as the virtual baseline so
      // A2 can constrain exact dependency/route threading without rewriting the
      // committed A1 baseline.
      headBaseline.legacyDeclarationHashes[key]
        = baseBaseline.legacyDeclarationHashes[key];
    }
  }
  if (baseBaseline) {
    headBaseline.legacyDeclarationHashes = Object.fromEntries(
      Object.entries(headBaseline.legacyDeclarationHashes).sort(([left], [right]) => left.localeCompare(right)),
    );
    assertNoExpansion(headBaseline, baseBaseline);
  }
  const actual = currentAllowances(root, allowTransitionalParagraphAnchorAdapter);
  if (baseBaseline) assertNoExpansion(actual, baseBaseline);
  assertExactBaseline(headBaseline, actual);
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    checkDocxLayoutBoundaries(parseArguments(process.argv.slice(2)));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
