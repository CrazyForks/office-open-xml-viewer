#!/usr/bin/env node

import { mkdir, readdir, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { rolldown } from 'rolldown';
import { dts } from 'rolldown-plugin-dts';
import { API } from 'typescript/unstable/sync';
import {
  SyntaxKind,
  createScanner,
  getJSDocTags,
  isClassDeclaration,
  isClassExpression,
  isInterfaceDeclaration,
} from 'typescript/unstable/ast';

const entries = ['index', 'docx', 'xlsx', 'pptx', 'math', 'node'];
const dist = path.resolve(process.cwd(), 'dist');
const workDir = path.join(dist, '.types-work');
const outDir = path.join(dist, 'types');

function stripInternalMembers(sourceFile) {
  const isInternal = (node) => getJSDocTags(node)
    .some((tag) => tag.tagName.text === 'internal');
  const ranges = [];
  const visit = (node) => {
    if (isClassDeclaration(node) || isClassExpression(node) || isInterfaceDeclaration(node)) {
      for (const member of node.members) {
        if (isInternal(member)) ranges.push([member.getFullStart(), member.getEnd()]);
      }
    }
    node.forEachChild(visit);
  };
  visit(sourceFile);
  return ranges
    .sort((left, right) => right[0] - left[0])
    .reduce(
      (source, [start, end]) => `${source.slice(0, start)}${source.slice(end)}`,
      sourceFile.getFullText(),
    );
}

function stripComments(source) {
  // Rolldown's Oxc declaration resolver can reattach JSDoc between modifiers
  // and a member name, which changes the parsed public surface under tsgo.
  // Strip comments only after @internal members have been identified and
  // removed; the declarations themselves remain the compiler-owned source of
  // truth and are compiled again by check-public-type-exports.mjs.
  const scanner = createScanner(false, undefined, source);
  const tokens = [];
  for (let kind = scanner.scan(); kind !== SyntaxKind.EndOfFile; kind = scanner.scan()) {
    if (kind !== SyntaxKind.SingleLineCommentTrivia && kind !== SyntaxKind.MultiLineCommentTrivia) {
      tokens.push(scanner.getTokenText());
    }
  }
  return tokens.join('');
}

async function declarationFiles(root) {
  const entries = await readdir(root, { recursive: true, withFileTypes: true });
  return entries
    .filter((entry) => entry.isFile() && entry.name.endsWith('.d.ts'))
    .map((entry) => path.join(entry.parentPath ?? entry.path, entry.name));
}

async function prepareDeclarationInputs(files) {
  const api = new API({ cwd: workDir });
  const snapshot = api.updateSnapshot({ openFiles: files });
  try {
    for (const file of files) {
      const project = snapshot.getDefaultProjectForFile(file);
      const sourceFile = project?.program.getSourceFile(file);
      if (!project || !sourceFile) throw new Error(`Could not load declaration: ${file}`);
      const withoutInternals = stripInternalMembers(sourceFile);
      await writeFile(file, stripComments(withoutInternals));
    }
  } finally {
    snapshot.dispose();
    api.close();
  }
}

await mkdir(outDir, { recursive: true });

try {
  await prepareDeclarationInputs(await declarationFiles(workDir));

  await Promise.all(entries.map(async (entry) => {
    const build = await rolldown({
      input: path.join(workDir, `${entry}.d.ts`),
      plugins: [dts({ dtsInput: true })],
    });
    try {
      await build.write({
        file: path.join(outDir, `${entry}.d.ts`),
        format: 'es',
        codeSplitting: false,
      });
    } finally {
      await build.close();
    }
  }));
} finally {
  await rm(workDir, { recursive: true, force: true });
}
