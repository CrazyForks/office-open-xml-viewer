#!/usr/bin/env node

import { execFileSync } from 'node:child_process';
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';
import { API } from 'typescript/unstable/sync';
import {
  SyntaxKind,
  TokenFlags,
  createScanner,
  isClassDeclaration,
  isClassExpression,
  isConstructorDeclaration,
  isEnumDeclaration,
  isExportDeclaration,
  isFunctionDeclaration,
  isIdentifier,
  isImportEqualsDeclaration,
  isIntersectionTypeNode,
  isInterfaceDeclaration,
  isModuleDeclaration,
  isNamedExports,
  isParenthesizedTypeNode,
  isPrivateIdentifier,
  isSourceFile,
  isStringLiteral,
  isTypeAliasDeclaration,
  isUnionTypeNode,
  isVariableStatement,
  visitEachChild,
} from 'typescript/unstable/ast';
import {
  createConstructorDeclaration,
  createIdentifier,
  createIntersectionTypeNode,
  createNodeArray,
  createPrivateIdentifier,
  createPropertyDeclaration,
  createStringLiteral,
  createToken,
  createUnionTypeNode,
  updateClassDeclaration,
  updateClassExpression,
  updateEnumDeclaration,
  updateFunctionDeclaration,
  updateImportEqualsDeclaration,
  updateInterfaceDeclaration,
  updateModuleDeclaration,
  updateSourceFile,
  updateTypeAliasDeclaration,
  updateVariableStatement,
} from 'typescript/unstable/ast/factory';

function parseArgs(argv) {
  const result = {
    root: process.cwd(),
    entry: 'docx.d.ts',
    baseline: 'packages/docx/api/public-api-baseline.d.ts',
    label: 'DOCX',
    baseRef: undefined,
    writeBaseline: false,
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--root') result.root = path.resolve(argv[++index]);
    else if (arg === '--entry') result.entry = argv[++index];
    else if (arg === '--baseline') result.baseline = argv[++index];
    else if (arg === '--label') result.label = argv[++index];
    else if (arg === '--base-ref') result.baseRef = argv[++index];
    else if (arg === '--write-baseline') result.writeBaseline = true;
    else throw new Error(`Unknown argument: ${arg}`);
  }
  return result;
}

function normalizeText(source) {
  return source
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .filter((line) => !/^\/\/# sourceMappingURL=/.test(line))
    .map((line) => line.replace(/[ \t]+$/g, ''))
    .join('\n')
    .trimEnd();
}

function canonicalTokens(source) {
  const scanner = createScanner(false, undefined, source);
  const tokens = [];
  for (let kind = scanner.scan(); kind !== SyntaxKind.EndOfFile; kind = scanner.scan()) {
    if (
      kind !== SyntaxKind.WhitespaceTrivia
      && kind !== SyntaxKind.NewLineTrivia
      && kind !== SyntaxKind.SingleLineCommentTrivia
      && kind !== SyntaxKind.MultiLineCommentTrivia
    ) {
      tokens.push(`${kind}:${scanner.getTokenText()}`);
    }
  }
  return tokens.join('\n');
}

export function normalizeDeclaration(source, fileName) {
  const workDir = mkdtempSync(path.join(tmpdir(), 'ooxml-public-api-'));
  const sourcePath = path.join(workDir, path.basename(fileName).endsWith('.ts')
    ? path.basename(fileName)
    : 'declaration.d.ts');
  writeFileSync(sourcePath, source);
  const api = new API({ cwd: workDir });
  const snapshot = api.updateSnapshot({ openFiles: [sourcePath] });
  try {
    const project = snapshot.getDefaultProjectForFile(sourcePath);
    const parsed = project?.program.getSourceFile(sourcePath);
    if (!project || !parsed) throw new Error(`Could not parse declaration: ${fileName}`);

    const localExportNames = new Set();
    const exportAliases = new Map();
    for (const statement of parsed.statements) {
      if (!isExportDeclaration(statement) || statement.moduleSpecifier || !statement.exportClause
          || !isNamedExports(statement.exportClause)) continue;
      for (const element of statement.exportClause.elements) {
        const localName = element.propertyName?.text ?? element.name.text;
        localExportNames.add(element.name.text);
        if (localName !== element.name.text) exportAliases.set(localName, element.name.text);
      }
    }
    const collisionGroups = new Map();
    for (const statement of parsed.statements) {
      if (!('name' in statement) || !statement.name || !isIdentifier(statement.name)) continue;
      if (localExportNames.has(statement.name.text)) continue;
      const match = /^(.*?)(?:_|\$)\d+$/.exec(statement.name.text);
      if (!match) continue;
      const names = collisionGroups.get(match[1]) ?? [];
      names.push(statement.name.text);
      collisionGroups.set(match[1], names);
    }
    const collisionAliases = new Map();
    for (const [base, names] of collisionGroups) {
      names.sort((left, right) => {
        const leftOrdinal = Number(/\d+$/.exec(left)?.[0]);
        const rightOrdinal = Number(/\d+$/.exec(right)?.[0]);
        return leftOrdinal - rightOrdinal || left.localeCompare(right);
      });
      names.forEach((name, index) => {
        collisionAliases.set(name, `${base}__emitterCollision${index + 1}`);
      });
    }

    const hasModifier = (member, kind) => member.modifiers?.some(
      (modifier) => modifier.kind === kind,
    ) ?? false;
    const normalizeClassMembers = (members) => {
      let hasInstancePrivate = false;
      let hasStaticPrivate = false;
      let hasHardPrivate = false;
      let hasPrivateConstructor = false;
      const visibleMembers = members.filter((member) => {
        if (member.name && isPrivateIdentifier(member.name)) {
          hasHardPrivate = true;
          return false;
        }
        if (!hasModifier(member, SyntaxKind.PrivateKeyword)) return true;
        if (isConstructorDeclaration(member)) hasPrivateConstructor = true;
        else if (hasModifier(member, SyntaxKind.StaticKeyword)) hasStaticPrivate = true;
        else hasInstancePrivate = true;
        return false;
      });
      if (visibleMembers.length === members.length) return members;
      if (hasInstancePrivate) {
        visibleMembers.push(createPropertyDeclaration(
          [createToken(SyntaxKind.PrivateKeyword)],
          createIdentifier('__privatePresence'),
        ));
      }
      if (hasStaticPrivate) {
        visibleMembers.push(createPropertyDeclaration(
          [
            createToken(SyntaxKind.PrivateKeyword),
            createToken(SyntaxKind.StaticKeyword),
          ],
          createIdentifier('__staticPrivatePresence'),
        ));
      }
      if (hasHardPrivate) {
        visibleMembers.push(createPropertyDeclaration(
          undefined,
          createPrivateIdentifier('#private'),
        ));
      }
      if (hasPrivateConstructor) {
        visibleMembers.push(createConstructorDeclaration(
          [createToken(SyntaxKind.PrivateKeyword)],
          undefined,
          [],
        ));
      }
      return createNodeArray(visibleMembers);
    };
    const statementName = (statement) => {
      if ('name' in statement && statement.name && isIdentifier(statement.name)) {
        return statement.name.text;
      }
      if (isVariableStatement(statement)) {
        return statement.declarationList.declarations
          .map((declaration) => isIdentifier(declaration.name) ? declaration.name.text : '')
          .join(',');
      }
      return '';
    };
    const updateModifiers = (statement, modifiers) => {
      if (isVariableStatement(statement)) {
        return updateVariableStatement(statement, modifiers, statement.declarationList);
      }
      if (isFunctionDeclaration(statement)) {
        return updateFunctionDeclaration(
          statement,
          modifiers,
          statement.asteriskToken,
          statement.name,
          statement.typeParameters,
          statement.parameters,
          statement.type,
          statement.body,
        );
      }
      if (isClassDeclaration(statement)) {
        return updateClassDeclaration(
          statement,
          modifiers,
          statement.name,
          statement.typeParameters,
          statement.heritageClauses,
          statement.members,
        );
      }
      if (isInterfaceDeclaration(statement)) {
        return updateInterfaceDeclaration(
          statement,
          modifiers,
          statement.name,
          statement.typeParameters,
          statement.heritageClauses,
          statement.members,
        );
      }
      if (isTypeAliasDeclaration(statement)) {
        return updateTypeAliasDeclaration(
          statement,
          modifiers,
          statement.name,
          statement.typeParameters,
          statement.type,
        );
      }
      if (isEnumDeclaration(statement)) {
        return updateEnumDeclaration(statement, modifiers, statement.name, statement.members);
      }
      if (isModuleDeclaration(statement)) {
        return updateModuleDeclaration(statement, modifiers, statement.name, statement.body);
      }
      if (isImportEqualsDeclaration(statement)) {
        return updateImportEqualsDeclaration(
          statement,
          modifiers,
          statement.name,
          statement.moduleReference,
        );
      }
      return statement;
    };
    const normalizeSourceFile = (sourceFile) => {
      // API Extractor emits `export declare interface Foo`, whereas Rolldown's
      // declaration bundler emits `interface Foo` plus `export { Foo }`. Treat
      // those equivalent spellings alike so the guard protects the API rather
      // than coupling the project to one declaration emitter.
      const localExports = new Set();
      for (const statement of sourceFile.statements) {
        if (!isExportDeclaration(statement) || statement.moduleSpecifier || !statement.exportClause) continue;
        if (!isNamedExports(statement.exportClause)) continue;
        for (const element of statement.exportClause.elements) {
          const localName = element.propertyName?.text ?? element.name.text;
          if (localName === element.name.text) localExports.add(localName);
        }
      }

      const statements = [];
      for (const statement of sourceFile.statements) {
        if (isExportDeclaration(statement) && !statement.moduleSpecifier && statement.exportClause
            && isNamedExports(statement.exportClause)
            && statement.exportClause.elements.every((element) => {
              const localName = element.propertyName?.text ?? element.name.text;
              return localName === element.name.text;
            })) {
          continue;
        }
        const name = statementName(statement);
        if ('modifiers' in statement) {
          // `declare` is optional in an ambient .d.ts module and emitters differ
          // on whether they spell it explicitly.
          const modifiers = (statement.modifiers ?? [])
            .filter((modifier) => modifier.kind !== SyntaxKind.DeclareKeyword);
          if (name && localExports.has(name)
              && !modifiers.some((modifier) => modifier.kind === SyntaxKind.ExportKeyword)) {
            statements.push(updateModifiers(statement, [
              createToken(SyntaxKind.ExportKeyword),
              ...modifiers,
            ]));
            continue;
          }
          statements.push(updateModifiers(statement, modifiers));
          continue;
        }
        statements.push(statement);
      }

      // Declaration order is not part of a module's public surface. Keep
      // overloads and merged declarations with the same name stable while
      // canonicalizing the emitter-specific top-level ordering.
      const indexed = statements.map((statement, index) => ({ statement, index }));
      indexed.sort((left, right) => {
        const leftName = statementName(left.statement);
        const rightName = statementName(right.statement);
        const byName = leftName.localeCompare(rightName);
        return byName || left.index - right.index;
      });
      return updateSourceFile(
        sourceFile,
        createNodeArray(indexed.map(({ statement }) => statement)),
        sourceFile.endOfFileToken,
      );
    };
    const visit = (node) => {
      if (isIdentifier(node) && exportAliases.has(node.text)) {
        return createIdentifier(exportAliases.get(node.text));
      }
      if (isIdentifier(node) && collisionAliases.has(node.text)) {
        return createIdentifier(collisionAliases.get(node.text));
      }
      if (isParenthesizedTypeNode(node)) return visit(node.type);
      if (isStringLiteral(node)) return createStringLiteral(node.text, TokenFlags.SingleQuote);
      const visited = visitEachChild(node, visit);
      if (isUnionTypeNode(visited)) {
        return createUnionTypeNode(visited.types.flatMap(
          (type) => isUnionTypeNode(type) ? [...type.types] : [type],
        ));
      }
      if (isIntersectionTypeNode(visited)) {
        return createIntersectionTypeNode(visited.types.flatMap(
          (type) => isIntersectionTypeNode(type) ? [...type.types] : [type],
        ));
      }
      if (isClassDeclaration(visited)) {
        return updateClassDeclaration(
          visited,
          visited.modifiers,
          visited.name,
          visited.typeParameters,
          visited.heritageClauses,
          normalizeClassMembers(visited.members),
        );
      }
      if (isClassExpression(visited)) {
        return updateClassExpression(
          visited,
          visited.modifiers,
          visited.name,
          visited.typeParameters,
          visited.heritageClauses,
          normalizeClassMembers(visited.members),
        );
      }
      if (isSourceFile(visited)) return normalizeSourceFile(visited);
      return visited;
    };
    const printed = project.emitter.printNode(visit(parsed));
    const scanner = createScanner(false, undefined, printed);
    const uncommented = [];
    for (let kind = scanner.scan(); kind !== SyntaxKind.EndOfFile; kind = scanner.scan()) {
      if (kind !== SyntaxKind.SingleLineCommentTrivia && kind !== SyntaxKind.MultiLineCommentTrivia) {
        uncommented.push(scanner.getTokenText());
      }
    }
    return normalizeText(uncommented.join(''));
  } finally {
    snapshot.dispose();
    api.close();
    rmSync(workDir, { recursive: true, force: true });
  }
}

function localSpecifiers(source) {
  const values = [];
  for (const match of source.matchAll(/\b(?:from|import\s*\()\s*(['"])([^'"]+)\1/g)) {
    values.push(match[2]);
  }
  for (const match of source.matchAll(/^\s*import\s*(['"])([^'"]+)\1/gm)) {
    values.push(match[2]);
  }
  for (const match of source.matchAll(/^\s*\/\/\/\s*<reference\s+path=(['"])([^'"]+)\1/gm)) {
    values.push(match[2]);
  }
  return [...new Set(values.filter((specifier) => specifier.startsWith('.')))];
}

function resolveDeclaration(fromFile, specifier) {
  const absolute = path.resolve(path.dirname(fromFile), specifier);
  const withoutRuntimeExtension = absolute.replace(/\.(?:mjs|cjs|js|jsx|ts|tsx)$/, '');
  const candidates = [
    absolute,
    `${absolute}.d.ts`,
    `${withoutRuntimeExtension}.d.ts`,
    path.join(absolute, 'index.d.ts'),
    path.join(withoutRuntimeExtension, 'index.d.ts'),
  ];
  return candidates.find((candidate) => existsSync(candidate));
}

function collectDeclarations(typesRoot, entryName) {
  const entry = path.resolve(typesRoot, entryName);
  const entryRelative = path.relative(typesRoot, entry);
  if (entryRelative.startsWith('..') || path.isAbsolute(entryRelative)) {
    throw new Error(`Generated declaration entry must stay inside dist/types (${entryName}).`);
  }
  if (!existsSync(entry)) {
    throw new Error(`Generated declaration entry is missing (${entryName}); build the published package first.`);
  }

  const pending = [entry];
  const sources = new Map();
  while (pending.length > 0) {
    const file = pending.pop();
    if (sources.has(file)) continue;
    const source = readFileSync(file, 'utf8');
    sources.set(file, normalizeDeclaration(source, file));
    for (const specifier of localSpecifiers(source)) {
      const resolved = resolveDeclaration(file, specifier);
      if (!resolved) {
        const relative = path.relative(typesRoot, file);
        throw new Error(`Cannot resolve local declaration ${specifier} from ${relative}.`);
      }
      pending.push(resolved);
    }
  }

  return [...sources]
    .map(([file, source]) => ({ file: path.relative(typesRoot, file).split(path.sep).join('/'), source }))
    .sort((left, right) => left.file.localeCompare(right.file));
}

function renderBaseline(declarations, label) {
  const header = [
    '// Generated by scripts/check-public-api.mjs.',
    `// This file records every local declaration reachable from the ${label} public entry.`,
    '// Do not edit by hand.',
  ].join('\n');
  const modules = declarations.map(({ file, source }) => `// --- file: ${file} ---\n${source}`);
  return `${header}\n\n${modules.join('\n\n')}\n`;
}

function normalizeRenderedBaseline(source) {
  return normalizeText(source)
    .split(/\n\n(?=\/\/ --- file: )/)
    .map((section, index) => {
      if (index === 0) return section;
      const declarationStart = section.indexOf('\n');
      const heading = section.slice(0, declarationStart);
      const fileName = heading.slice('// --- file: '.length, -' ---'.length);
      const normalized = normalizeDeclaration(section.slice(declarationStart + 1), fileName);
      return `${heading}\n${canonicalTokens(normalized)}`;
    })
    .join('\n\n');
}

function resolveMergeBase(root, explicit) {
  for (const candidate of explicit ? [explicit] : ['origin/main', 'main']) {
    try {
      return execFileSync('git', ['merge-base', candidate, 'HEAD'], {
        cwd: root,
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'ignore'],
      }).trim();
    } catch {}
  }
  throw new Error('Cannot resolve the merge base; fetch origin/main or pass --base-ref.');
}

function refContains(root, ref, relativePath) {
  try {
    execFileSync('git', ['cat-file', '-e', `${ref}:${relativePath}`], {
      cwd: root,
      stdio: 'ignore',
    });
    return true;
  } catch {
    return false;
  }
}

export function checkPublicApi(options) {
  const typesRoot = path.join(options.root, 'dist/types');
  const baselineRelative = options.baseline ?? 'packages/docx/api/public-api-baseline.d.ts';
  const baselinePath = path.join(options.root, baselineRelative);
  const resolvedBaselineRelative = path.relative(options.root, baselinePath);
  if (resolvedBaselineRelative.startsWith('..') || path.isAbsolute(resolvedBaselineRelative)) {
    throw new Error(`Public API baseline must stay inside the repository (${baselineRelative}).`);
  }
  const label = options.label ?? 'DOCX';
  const actual = renderBaseline(collectDeclarations(typesRoot, options.entry), label);
  const mergeBase = resolveMergeBase(options.root, options.baseRef);

  if (options.writeBaseline) {
    if (refContains(options.root, mergeBase, baselineRelative)) {
      throw new Error('--write-baseline is only permitted before the merge base contains the public API baseline.');
    }
    mkdirSync(path.dirname(baselinePath), { recursive: true });
    writeFileSync(baselinePath, actual);
    process.stdout.write(`Wrote ${baselineRelative}.\n`);
    return;
  }

  if (!existsSync(baselinePath)) {
    throw new Error(`Public API baseline is missing (${baselineRelative}).`);
  }
  const expected = normalizeText(readFileSync(baselinePath, 'utf8'));
  if (normalizeRenderedBaseline(actual) !== normalizeRenderedBaseline(expected)) {
    throw new Error(
      `${label} public API declaration baseline differs. Rebuild, inspect the reachable declarations, and update the committed baseline only for an intentional compatible API change.`,
    );
  }
  process.stdout.write(`${label} public API declaration baseline matches.\n`);
}

function isMainModule() {
  return process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
}

if (isMainModule()) {
  try {
    checkPublicApi(parseArgs(process.argv.slice(2)));
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  }
}
