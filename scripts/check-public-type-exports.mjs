#!/usr/bin/env node

import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import path from 'node:path';
import process from 'node:process';

const require = createRequire(new URL('../package.json', import.meta.url));
const ts = require('typescript');
const typesDir = path.resolve(process.cwd(), 'dist/types');
const formats = ['docx', 'pptx', 'xlsx'];
const files = ['index', ...formats, 'math'].map((entry) => path.join(typesDir, `${entry}.d.ts`));

const program = ts.createProgram(files, {
  module: ts.ModuleKind.ESNext,
  moduleResolution: ts.ModuleResolutionKind.Bundler,
  target: ts.ScriptTarget.ES2022,
  lib: ['lib.es2022.d.ts', 'lib.dom.d.ts', 'lib.dom.iterable.d.ts'],
  strict: true,
  noEmit: true,
});
const diagnostics = ts.getPreEmitDiagnostics(program);
if (diagnostics.length > 0) {
  throw new Error(ts.formatDiagnosticsWithColorAndContext(diagnostics, {
    getCanonicalFileName: (file) => file,
    getCurrentDirectory: () => process.cwd(),
    getNewLine: () => '\n',
  }));
}

const checker = program.getTypeChecker();
const moduleExports = (file) => {
  const source = program.getSourceFile(file);
  const symbol = source && checker.getSymbolAtLocation(source);
  assert.ok(symbol, `Cannot resolve declaration module ${path.relative(process.cwd(), file)}.`);
  return new Map(checker.getExportsOfModule(symbol).map((entry) => [entry.name, entry]));
};

const rootExports = moduleExports(files[0]);
for (const [index, format] of formats.entries()) {
  let namespace = rootExports.get(format);
  assert.ok(namespace, `Root declaration does not export the ${format} namespace.`);
  if (namespace.flags & ts.SymbolFlags.Alias) namespace = checker.getAliasedSymbol(namespace);
  const namespaceNames = checker.getExportsOfModule(namespace).map((entry) => entry.name).sort();
  const directNames = [...moduleExports(files[index + 1]).keys()].sort();
  assert.deepEqual(
    namespaceNames,
    directNames,
    `Root ${format} namespace differs from the ./${format} entry point.`,
  );
}

process.stdout.write('Published declaration entries compile and root namespace exports match.\n');
