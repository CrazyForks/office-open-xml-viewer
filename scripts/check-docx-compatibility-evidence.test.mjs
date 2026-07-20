import assert from 'node:assert/strict';
import { cpSync, mkdtempSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

const repositoryRoot = resolve(import.meta.dirname, '..');
const checker = resolve(import.meta.dirname, 'check-docx-compatibility-evidence.mjs');

function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'docx-compatibility-evidence-'));
  cpSync(
    resolve(repositoryRoot, 'packages/docx/src/layout/compatibility.ts'),
    resolve(root, 'packages/docx/src/layout/compatibility.ts'),
    { recursive: true },
  );
  const references = [
    ['packages/docx/src/layout/coordinate-space.test.ts', 'maps Transitional text direction %s to %s'],
    ['packages/docx/src/float-line-start-one-inch.test.ts', '(e) the boundary is identical across scales (absolute pt width)'],
    ['packages/docx/src/layout/floats.test.ts', 'keeps observed different-paragraph displacement on exclusion bounds'],
    ['packages/docx/src/float-table-page-fit.test.ts', '(g) DEFERS a page-anchored floating table when its raw band intersects an existing table float'],
  ];
  for (const [path, title] of references) {
    write(root, path, `it(${JSON.stringify(title)}, () => {});\n`);
  }
  write(root, 'scripts/docx-compatibility-observation-baseline.json',
    '{ "version": 1, "observations": [] }\n');
  return root;
}

function write(root, path, source) {
  const absolute = resolve(root, path);
  mkdirSync(dirname(absolute), { recursive: true });
  writeFileSync(absolute, source);
}

function run(root) {
  return spawnSync(process.execPath, [checker, '--root', root], {
    cwd: repositoryRoot,
    encoding: 'utf8',
  });
}

test('accepts literal rules with live regression evidence', () => {
  const result = run(fixture());
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /verified \(4 rules\)/);
});

test('rejects a stale regression test title', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/floats.test.ts', 'it("renamed", () => {});\n');
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_EVIDENCE_STALE/);
});

test('rejects a rule declaration outside the closed authority set', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table.ts', `
    import { defineCompatibilityRule } from './compatibility.js';
    export const WORD_INLINE = defineCompatibilityRule({
      id: 'word-inline',
      evidence: {
        kind: 'regression-test',
        reference: 'packages/docx/src/layout/floats.test.ts#keeps observed different-paragraph displacement on exclusion bounds',
      },
      description: 'Inline rule',
    });
  `);
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_DECLARATION_AUTHORITY/);
});

test('rejects duplicate rule ids across compatibility modules', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table-compatibility.ts', `
    import { defineCompatibilityRule } from './compatibility.js';
    export const WORD_DUPLICATE = defineCompatibilityRule({
      id: 'word-square-line-start-one-inch',
      evidence: { kind: 'microsoft-note', reference: '[MS-OI29500] §2.1.1' },
      description: 'Duplicate rule',
    });
  `);
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_DUPLICATE_ID/);
});

test('rejects an aliased factory import', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table-compatibility.ts', `
    import { defineCompatibilityRule as defineRule } from './compatibility.js';
    export const WORD_ALIAS = defineCompatibilityRule({
      id: 'word-alias',
      evidence: { kind: 'microsoft-note', reference: '[MS-OI29500] §2.1.1' },
      description: 'Aliased factory',
    });
  `);
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_FACTORY_IMPORT/);
});

test('rejects an unstructured Microsoft note reference', () => {
  const root = fixture();
  const source = readFileSync(
    resolve(root, 'packages/docx/src/layout/compatibility.ts'),
    'utf8',
  ).replace(
    "kind: 'regression-test',\n    reference: 'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s'",
    "kind: 'microsoft-note',\n    reference: 'MS note somewhere'",
  );
  write(root, 'packages/docx/src/layout/compatibility.ts', source);
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /COMPATIBILITY_MICROSOFT_REFERENCE/);
});

test('rejects a new inline Office observation outside compatibility modules', () => {
  const root = fixture();
  write(root, 'packages/docx/src/layout/table.ts',
    '// Word uses an undocumented magic branch here.\nexport const value = 1;\n');
  const result = run(root);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /INLINE_COMPATIBILITY_OBSERVATION/);
});

test('allows only an exact transitional observation-baseline entry', () => {
  const root = fixture();
  const marker = 'packages/docx/src/layout/table.ts::// Word uses an evidenced legacy branch.';
  write(root, 'packages/docx/src/layout/table.ts',
    '// Word uses an evidenced legacy branch.\nexport const value = 1;\n');
  write(root, 'scripts/docx-compatibility-observation-baseline.json',
    `${JSON.stringify({ version: 1, observations: [marker] }, null, 2)}\n`);
  assert.equal(run(root).status, 0);

  write(root, 'packages/docx/src/layout/table.ts',
    '// Word uses a changed legacy branch.\nexport const value = 1;\n');
  const changed = run(root);
  assert.equal(changed.status, 1);
  assert.match(changed.stderr, /INLINE_COMPATIBILITY_OBSERVATION/);
});
