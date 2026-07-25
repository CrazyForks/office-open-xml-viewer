#!/usr/bin/env node

import { createRequire } from 'node:module';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { rolldown } from 'rolldown';
import { dts } from 'rolldown-plugin-dts';

const require = createRequire(new URL('../package.json', import.meta.url));
const ts = require('typescript');

const entries = ['index', 'docx', 'xlsx', 'pptx', 'math'];
const dist = path.resolve(process.cwd(), 'dist');
const workDir = path.join(dist, '.types-work');
const outDir = path.join(dist, 'types');

function stripInternalMembers(source, fileName) {
  const parsed = ts.createSourceFile(
    fileName,
    source,
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TS,
  );
  const transformed = ts.transform(parsed, [(context) => {
    const isInternal = (node) => ts.getJSDocTags(node)
      .some((tag) => tag.tagName.text === 'internal');
    const visit = (node) => {
      const visited = ts.visitEachChild(node, visit, context);
      if (ts.isClassDeclaration(visited)) {
        return ts.factory.updateClassDeclaration(
          visited,
          visited.modifiers,
          visited.name,
          visited.typeParameters,
          visited.heritageClauses,
          visited.members.filter((member) => !isInternal(member)),
        );
      }
      if (ts.isClassExpression(visited)) {
        return ts.factory.updateClassExpression(
          visited,
          visited.modifiers,
          visited.name,
          visited.typeParameters,
          visited.heritageClauses,
          visited.members.filter((member) => !isInternal(member)),
        );
      }
      if (ts.isInterfaceDeclaration(visited)) {
        return ts.factory.updateInterfaceDeclaration(
          visited,
          visited.modifiers,
          visited.name,
          visited.typeParameters,
          visited.heritageClauses,
          visited.members.filter((member) => !isInternal(member)),
        );
      }
      return visited;
    };
    return (root) => ts.visitNode(root, visit);
  }]);
  const printer = ts.createPrinter({ newLine: ts.NewLineKind.LineFeed });
  const result = printer.printFile(transformed.transformed[0]);
  transformed.dispose();
  return result;
}

await mkdir(outDir, { recursive: true });

try {
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
      const output = path.join(outDir, `${entry}.d.ts`);
      await writeFile(
        output,
        stripInternalMembers(await readFile(output, 'utf8'), output),
      );
    } finally {
      await build.close();
    }
  }));
} finally {
  await rm(workDir, { recursive: true, force: true });
}
