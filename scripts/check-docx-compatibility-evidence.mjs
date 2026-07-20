#!/usr/bin/env node

import {
  existsSync,
  readFileSync,
  readdirSync,
  statSync,
} from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { join, relative, resolve, sep } from 'node:path';

const require = createRequire(new URL('../packages/docx/package.json', import.meta.url));
const ts = require('typescript');

const DOCX_SOURCE = 'packages/docx/src';
const COMPATIBILITY_OWNER = `${DOCX_SOURCE}/layout/compatibility.ts`;
const OBSERVATION_BASELINE =
  'scripts/docx-compatibility-observation-baseline.json';
const ALLOWED_DECLARATION_FILES = new Set([
  COMPATIBILITY_OWNER,
  `${DOCX_SOURCE}/layout/anchor-compatibility.ts`,
  `${DOCX_SOURCE}/layout/body-pagination-compatibility.ts`,
  `${DOCX_SOURCE}/layout/page-flow-compatibility.ts`,
  `${DOCX_SOURCE}/layout/section-compatibility.ts`,
  `${DOCX_SOURCE}/layout/table-compatibility.ts`,
]);

function fail(code, detail) {
  throw new Error(`${code}: ${detail}`);
}

function posixPath(path) {
  return path.split(sep).join('/');
}

function listFiles(root) {
  if (!existsSync(root)) return [];
  return readdirSync(root).flatMap((entry) => {
    const path = join(root, entry);
    return statSync(path).isDirectory() ? listFiles(path) : [path];
  });
}

function isProductionTypeScript(path) {
  return /\.tsx?$/.test(path)
    && !path.endsWith('.d.ts')
    && !/\.(test|spec|stories|test-support)\.tsx?$/.test(path)
    && !path.includes(`${sep}wasm${sep}`);
}

function property(object, name) {
  return object.properties.find((entry) =>
    ts.isPropertyAssignment(entry)
    && ((ts.isIdentifier(entry.name) && entry.name.text === name)
      || (ts.isStringLiteral(entry.name) && entry.name.text === name)));
}

function stringValue(object, name, file) {
  const entry = property(object, name);
  if (!entry || !ts.isStringLiteralLike(entry.initializer)) {
    fail('COMPATIBILITY_LITERAL_SCHEMA', `${file} requires literal ${name}`);
  }
  if (entry.initializer.text.trim() === '') {
    fail('COMPATIBILITY_LITERAL_SCHEMA', `${file} has empty ${name}`);
  }
  return entry.initializer.text;
}

function objectValue(object, name, file) {
  const entry = property(object, name);
  if (!entry || !ts.isObjectLiteralExpression(entry.initializer)) {
    fail('COMPATIBILITY_LITERAL_SCHEMA', `${file} requires object ${name}`);
  }
  return entry.initializer;
}

function exportName(call, file) {
  const declaration = call.parent;
  if (!ts.isVariableDeclaration(declaration)
    || !ts.isIdentifier(declaration.name)
    || declaration.initializer !== call) {
    fail(
      'COMPATIBILITY_DECLARATION_SHAPE',
      `${file} must assign each rule directly to a named const`,
    );
  }
  const statement = declaration.parent?.parent;
  if (!ts.isVariableStatement(statement)
    || !statement.modifiers?.some((modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword)
    || (declaration.parent.flags & ts.NodeFlags.Const) === 0) {
    fail(
      'COMPATIBILITY_DECLARATION_SHAPE',
      `${file} must export each rule as a const`,
    );
  }
  return declaration.name.text;
}

function verifyFactoryImport(source, file) {
  if (file === COMPATIBILITY_OWNER) return;
  const imports = source.statements.filter(ts.isImportDeclaration).filter((entry) =>
    ts.isStringLiteral(entry.moduleSpecifier)
    && entry.moduleSpecifier.text === './compatibility.js');
  const exact = imports.flatMap((entry) => {
    const bindings = entry.importClause?.namedBindings;
    return bindings && ts.isNamedImports(bindings)
      ? bindings.elements.filter((element) => element.name.text === 'defineCompatibilityRule')
      : [];
  });
  if (exact.length !== 1 || exact[0].propertyName !== undefined) {
    fail(
      'COMPATIBILITY_FACTORY_IMPORT',
      `${file} must import defineCompatibilityRule exactly and without an alias`,
    );
  }
}

function verifyEvidence(root, file, ruleId, evidence) {
  const kind = stringValue(evidence, 'kind', file);
  if (kind === 'regression-test') {
    const reference = stringValue(evidence, 'reference', file);
    const separator = reference.indexOf('#');
    if (separator <= 0 || separator === reference.length - 1) {
      fail(
        'COMPATIBILITY_REGRESSION_REFERENCE',
        `${ruleId} must use path#test-title`,
      );
    }
    const path = reference.slice(0, separator);
    const title = reference.slice(separator + 1);
    if (!path.startsWith(`${DOCX_SOURCE}/`) || !/\.(test|spec)\.tsx?$/.test(path)) {
      fail(
        'COMPATIBILITY_REGRESSION_REFERENCE',
        `${ruleId} references non-DOCX test ${path}`,
      );
    }
    const absolute = resolve(root, path);
    if (!existsSync(absolute)) {
      fail('COMPATIBILITY_EVIDENCE_MISSING', `${ruleId} references missing ${path}`);
    }
    if (!readFileSync(absolute, 'utf8').includes(title)) {
      fail(
        'COMPATIBILITY_EVIDENCE_STALE',
        `${ruleId} test title is absent from ${path}: ${title}`,
      );
    }
    return;
  }
  if (kind === 'microsoft-note') {
    const reference = stringValue(evidence, 'reference', file);
    if (!/^\[MS-[A-Z0-9]+\] §§?\d/.test(reference)) {
      fail(
        'COMPATIBILITY_MICROSOFT_REFERENCE',
        `${ruleId} has invalid Microsoft note ${reference}`,
      );
    }
    return;
  }
  if (kind === 'office-observation') {
    const fixtureId = stringValue(evidence, 'syntheticFixtureId', file);
    stringValue(evidence, 'application', file);
    stringValue(evidence, 'version', file);
    stringValue(evidence, 'platform', file);
    if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(fixtureId)) {
      fail(
        'COMPATIBILITY_FIXTURE_ID',
        `${ruleId} has invalid synthetic fixture id ${fixtureId}`,
      );
    }
    return;
  }
  fail('COMPATIBILITY_EVIDENCE_KIND', `${ruleId} has unknown evidence kind ${kind}`);
}

function inspectSource(root, absolute, rules) {
  const file = posixPath(relative(root, absolute));
  const source = ts.createSourceFile(
    absolute,
    readFileSync(absolute, 'utf8'),
    ts.ScriptTarget.Latest,
    true,
  );
  const calls = [];
  const visit = (node) => {
    if (ts.isCallExpression(node)
      && ts.isIdentifier(node.expression)
      && node.expression.text === 'defineCompatibilityRule') {
      calls.push(node);
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  if (calls.length === 0) return;
  if (!ALLOWED_DECLARATION_FILES.has(file)) {
    fail(
      'COMPATIBILITY_DECLARATION_AUTHORITY',
      `${file} declares a compatibility rule`,
    );
  }
  verifyFactoryImport(source, file);
  for (const call of calls) {
    if (call.arguments.length !== 1 || !ts.isObjectLiteralExpression(call.arguments[0])) {
      fail(
        'COMPATIBILITY_LITERAL_SCHEMA',
        `${file} must pass one literal rule object`,
      );
    }
    const object = call.arguments[0];
    const name = exportName(call, file);
    if (!/^[A-Z][A-Z0-9_]+$/.test(name)) {
      fail('COMPATIBILITY_EXPORT_NAME', `${file} exports invalid rule name ${name}`);
    }
    const id = stringValue(object, 'id', file);
    stringValue(object, 'description', file);
    verifyEvidence(root, file, id, objectValue(object, 'evidence', file));
    const prior = rules.get(id);
    if (prior) {
      fail('COMPATIBILITY_DUPLICATE_ID', `${id} is declared by ${prior} and ${file}`);
    }
    rules.set(id, file);
  }
}

function options(argv) {
  let root = process.cwd();
  let printObservations = false;
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === '--root' && argv[index + 1]) {
      root = resolve(argv[index + 1]);
      index += 1;
      continue;
    }
    if (argv[index] === '--print-observations') {
      printObservations = true;
      continue;
    }
    fail('UNKNOWN_ARGUMENT', argv[index] ?? '<missing>');
  }
  return { root, printObservations };
}

const observationMarker =
  /\[MS-[A-Z0-9]+\]|\bobserved (?:Word|Office)\b|\bWord-compatible\b|\bWord(?:'s)? (?:uses|applies|adds|treats|admits|keeps|places|starts|does|doesn't|ignores|clips|measured|runtime)\b/i;

function observationKeys(root) {
  const sourceRoot = resolve(root, DOCX_SOURCE);
  return listFiles(sourceRoot)
    .filter(isProductionTypeScript)
    .flatMap((absolute) => {
      const file = posixPath(relative(root, absolute));
      if (file.endsWith('compatibility.ts')) return [];
      return readFileSync(absolute, 'utf8').split(/\r?\n/).flatMap((line) =>
        observationMarker.test(line)
          ? [`${file}::${line.trim().replace(/\s+/g, ' ')}`]
          : []);
    })
    .sort();
}

function verifyObservationBaseline(root, observations) {
  const path = resolve(root, OBSERVATION_BASELINE);
  const baseline = existsSync(path)
    ? JSON.parse(readFileSync(path, 'utf8'))
    : { version: 1, observations: [] };
  if (baseline?.version !== 1 || !Array.isArray(baseline.observations)
    || baseline.observations.some((entry) => typeof entry !== 'string')) {
    fail('COMPATIBILITY_OBSERVATION_BASELINE', OBSERVATION_BASELINE);
  }
  const allowed = new Set(baseline.observations);
  const added = observations.filter((entry) => !allowed.has(entry));
  if (added.length > 0) {
    fail(
      'INLINE_COMPATIBILITY_OBSERVATION',
      added.join('\n'),
    );
  }
}

export function checkDocxCompatibilityEvidence(root) {
  const rules = new Map();
  const sourceRoot = resolve(root, DOCX_SOURCE);
  for (const file of listFiles(sourceRoot).filter(isProductionTypeScript)) {
    inspectSource(root, file, rules);
  }
  if (rules.size === 0) fail('COMPATIBILITY_REGISTRY_EMPTY', DOCX_SOURCE);
  const observations = observationKeys(root);
  verifyObservationBaseline(root, observations);
  return Object.freeze([...rules].map(([id, file]) => Object.freeze({ id, file })));
}

if (process.argv[1]
  && resolve(process.argv[1]) === resolve(fileURLToPath(import.meta.url))) {
  try {
    const { root, printObservations } = options(process.argv.slice(2));
    if (printObservations) {
      process.stdout.write(`${JSON.stringify({
        version: 1,
        observations: observationKeys(root),
      }, null, 2)}\n`);
      process.exit(0);
    }
    const rules = checkDocxCompatibilityEvidence(root);
    process.stdout.write(`DOCX compatibility evidence verified (${rules.length} rules).\n`);
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  }
}
