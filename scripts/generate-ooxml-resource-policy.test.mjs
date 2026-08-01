import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { synchronizeResourcePolicy } from './generate-ooxml-resource-policy.mjs';

function fixture(policy = {
  defaults: {
    maxArchiveEntryBytes: 128,
    maxTotalInflatedBytes: 256,
  },
  hardCeilings: {
    maxArchiveEntryBytes: 512,
    maxTotalInflatedBytes: 1024,
    maxArchiveEntries: 20,
    maxCentralDirectoryBytes: 64,
    maxWorksheetRows: 100,
    maxWorksheetCells: 250,
    maxWorksheetCellContentUtf8Bytes: 320,
    maxWorksheetJsonBytes: 640,
    maxWorkbookCachedRows: 200,
    maxWorkbookCachedCells: 500,
    maxRendererCoordinateIndexEntries: 250,
  },
}) {
  const root = mkdtempSync(path.join(tmpdir(), 'ooxml-resource-policy-'));
  mkdirSync(path.join(root, 'packages/ooxml-common/src'), { recursive: true });
  mkdirSync(path.join(root, 'packages/core/src/worker'), { recursive: true });
  writeFileSync(
    path.join(root, 'packages/ooxml-common/resource-policy.json'),
    `${JSON.stringify(policy, null, 2)}\n`,
  );
  return root;
}

test('generates matching TypeScript and Rust constants from one policy source', (context) => {
  const root = fixture();
  context.after(() => rmSync(root, { recursive: true, force: true }));

  synchronizeResourcePolicy({ root, write: true });
  synchronizeResourcePolicy({ root, write: false });

  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /STANDARD_MAX_ARCHIVE_ENTRY_BYTES = 128/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /STANDARD_MAX_TOTAL_INFLATED_BYTES: u64 = 256/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'), 'utf8'),
    /HARD_MAX_XLSX_WORKSHEET_JSON_BYTES = 640/,
  );
  assert.match(
    readFileSync(path.join(root, 'packages/ooxml-common/src/resource-policy.generated.rs'), 'utf8'),
    /HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES: u64 = 250/,
  );
});

test('fails closed when a generated language file drifts', (context) => {
  const root = fixture();
  context.after(() => rmSync(root, { recursive: true, force: true }));
  synchronizeResourcePolicy({ root, write: true });
  writeFileSync(
    path.join(root, 'packages/core/src/worker/resource-policy.generated.ts'),
    '// stale\n',
  );

  assert.throws(
    () => synchronizeResourcePolicy({ root, write: false }),
    /Generated OOXML resource policy files are stale/,
  );
});

test('rejects invalid or internally inconsistent policy values', (context) => {
  const root = fixture({
    defaults: {
      maxArchiveEntryBytes: 513,
      maxTotalInflatedBytes: 1024,
    },
    hardCeilings: {
      maxArchiveEntryBytes: 512,
      maxTotalInflatedBytes: 1024,
      maxArchiveEntries: 20,
      maxCentralDirectoryBytes: 64,
      maxWorksheetRows: 100,
      maxWorksheetCells: 250,
      maxWorksheetCellContentUtf8Bytes: 320,
      maxWorksheetJsonBytes: 640,
      maxWorkbookCachedRows: 200,
      maxWorkbookCachedCells: 500,
      maxRendererCoordinateIndexEntries: 250,
    },
  });
  context.after(() => rmSync(root, { recursive: true, force: true }));

  assert.throws(
    () => synchronizeResourcePolicy({ root, write: true }),
    /must not exceed its hard ceiling/,
  );
});

test('accepts CRLF-normalized generated files on Windows-style checkouts', (context) => {
  const root = fixture();
  context.after(() => rmSync(root, { recursive: true, force: true }));
  synchronizeResourcePolicy({ root, write: true });
  for (const relativePath of [
    'packages/core/src/worker/resource-policy.generated.ts',
    'packages/ooxml-common/src/resource-policy.generated.rs',
  ]) {
    const file = path.join(root, relativePath);
    writeFileSync(file, readFileSync(file, 'utf8').replace(/\n/gu, '\r\n'));
  }

  assert.doesNotThrow(() => synchronizeResourcePolicy({ root, write: false }));
});
