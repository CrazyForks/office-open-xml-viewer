#!/usr/bin/env node

import assert from 'node:assert/strict';
import { rmSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { API, SymbolFlags } from 'typescript/unstable/sync';

const typesDir = path.resolve(process.cwd(), 'dist/types');
const formats = ['docx', 'pptx', 'xlsx'];
const files = ['index', ...formats, 'math'].map((entry) => path.join(typesDir, `${entry}.d.ts`));
const configPath = path.join(typesDir, '.public-type-exports.tsconfig.json');

writeFileSync(configPath, JSON.stringify({
  compilerOptions: {
    lib: ['ES2022', 'DOM', 'DOM.Iterable'],
    module: 'ESNext',
    moduleResolution: 'Bundler',
    noEmit: true,
    strict: true,
    target: 'ES2022',
  },
  files,
}));

const api = new API({ cwd: typesDir });
const snapshot = api.updateSnapshot({ openProjects: [configPath] });

try {
  const project = snapshot.getProject(configPath);
  assert.ok(project, 'Cannot load the published declaration verification project.');
  const { checker, program } = project;
  const diagnostics = [
    ...program.getConfigFileParsingDiagnostics(),
    ...program.getProgramDiagnostics(),
    ...program.getGlobalDiagnostics(),
    ...program.getSyntacticDiagnostics(),
    ...program.getSemanticDiagnostics(),
  ];
  if (diagnostics.length > 0) {
    const formatDiagnostic = (diagnostic, indent = '') => {
      const location = diagnostic.fileName
        ? `${path.relative(process.cwd(), diagnostic.fileName)}:${diagnostic.pos}`
        : 'TypeScript';
      const chained = diagnostic.messageChain?.map(
        (child) => formatDiagnostic(child, `${indent}  `),
      ).join('\n');
      return `${indent}${location} TS${diagnostic.code}: ${diagnostic.text}${chained ? `\n${chained}` : ''}`;
    };
    throw new Error(diagnostics.map((diagnostic) => formatDiagnostic(diagnostic)).join('\n'));
  }

  const moduleExports = (file) => {
    const source = program.getSourceFile(file);
    const symbol = source && checker.getSymbolAtLocation(source);
    assert.ok(symbol, `Cannot resolve declaration module ${path.relative(process.cwd(), file)}.`);
    return new Map(checker.getExportsOfModule(symbol).map((entry) => [entry.name, entry]));
  };

  const rootExports = moduleExports(files[0]);
  const formatExports = [];
  for (const [index, format] of formats.entries()) {
    let namespace = rootExports.get(format);
    assert.ok(namespace, `Root declaration does not export the ${format} namespace.`);
    if (namespace.flags & SymbolFlags.Alias) namespace = checker.getAliasedSymbol(namespace);
    const namespaceNames = checker.getExportsOfModule(namespace).map((entry) => entry.name).sort();
    const directExports = moduleExports(files[index + 1]);
    formatExports.push(directExports);
    const directNames = [...directExports.keys()].sort();
    assert.deepEqual(
      namespaceNames,
      directNames,
      `Root ${format} namespace differs from the ./${format} entry point.`,
    );
  }

  const sharedOoxmlTypes = [
    'LoadOptions',
    'OoxmlError',
    'OoxmlErrorStage',
    'OoxmlFormat',
    'OoxmlResourceLimit',
    'OoxmlResourceLimits',
    'OoxmlResourceMetric',
    'OoxmlResourceName',
    'OoxmlResourceLimitError',
    'OoxmlResourceLimitErrorDetails',
    'OoxmlResourceUsageSnapshot',
    'OoxmlResourceViolation',
  ];

  function declaredType(exports, name, format) {
    let symbol = exports.get(name);
    assert.ok(symbol, `${format} does not export shared OOXML type ${name}.`);
    if (symbol.flags & SymbolFlags.Alias) symbol = checker.getAliasedSymbol(symbol);
    return checker.getDeclaredTypeOfSymbol(symbol);
  }

  for (const name of sharedOoxmlTypes) {
    const canonical = declaredType(formatExports[0], name, formats[0]);
    for (let index = 1; index < formatExports.length; index += 1) {
      const candidate = declaredType(formatExports[index], name, formats[index]);
      assert.ok(
        checker.isTypeAssignableTo(canonical, candidate)
          && checker.isTypeAssignableTo(candidate, canonical),
        `${name} differs between ${formats[0]} and ${formats[index]}.`,
      );
    }
  }

  process.stdout.write(
    'Published declaration entries compile; root namespace exports and shared OOXML contracts match.\n',
  );
} finally {
  snapshot.dispose();
  api.close();
  rmSync(configPath, { force: true });
}
