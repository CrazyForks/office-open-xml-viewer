import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync, rmSync, statSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const coreDeclaration = path.join(root, 'packages/core/dist/types/index.d.ts');
const docxDeclaration = path.join(root, 'packages/docx/dist/types/index.d.ts');

function assertProductionDeclarations(packageName) {
  const typesDir = path.join(root, `packages/${packageName}/dist/types`);
  assert.equal(existsSync(path.join(typesDir, 'index.d.ts')), true);
  const emittedTests = readdirSync(typesDir, { recursive: true })
    .filter((file) => typeof file === 'string' && file.endsWith('.test.d.ts'));
  assert.deepEqual(emittedTests, []);
}

function buildPackage(name) {
  execFileSync('pnpm', ['--filter', name, 'build'], {
    cwd: root,
    stdio: 'pipe',
  });
}

test('building DOCX does not rebuild declaration outputs owned by core', () => {
  rmSync(path.join(root, 'packages/core/dist'), { recursive: true, force: true });
  rmSync(path.join(root, 'packages/core/tsconfig.tsbuildinfo'), { force: true });
  buildPackage('@silurus/ooxml-core');
  const before = {
    contents: readFileSync(coreDeclaration),
    modifiedMs: statSync(coreDeclaration).mtimeMs,
  };

  buildPackage('@silurus/ooxml-docx');

  assert.equal(existsSync(docxDeclaration), true);
  assertProductionDeclarations('docx');
  assert.deepEqual(readFileSync(coreDeclaration), before.contents);
  assert.equal(statSync(coreDeclaration).mtimeMs, before.modifiedMs);
});

test('standalone PPTX and XLSX builds retain production declarations', () => {
  buildPackage('@silurus/ooxml-pptx');
  assertProductionDeclarations('pptx');

  // XLSX must be independently buildable from a clean checkout, where the
  // core declaration output does not exist yet.
  rmSync(path.join(root, 'packages/core/dist'), { recursive: true, force: true });
  rmSync(path.join(root, 'packages/core/tsconfig.tsbuildinfo'), { force: true });
  rmSync(path.join(root, 'packages/xlsx/tsconfig.tsbuildinfo'), { force: true });
  buildPackage('@silurus/ooxml-xlsx');
  assert.equal(existsSync(coreDeclaration), true);
  assertProductionDeclarations('xlsx');
});
