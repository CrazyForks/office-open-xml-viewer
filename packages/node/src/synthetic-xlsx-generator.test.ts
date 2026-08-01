import { access, mkdtemp, readFile, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { inflateRawSync } from 'node:zlib';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { generateSyntheticXlsx } from '../scripts/generate-synthetic-xlsx.mjs';

let directory = '';

beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'ooxml-xlsx-generator-'));
});

afterEach(async () => {
  await rm(directory, { recursive: true, force: true });
});

describe('synthetic XLSX generator invariants', () => {
  it('preserves an exact target and rejects targets below natural length or with 1..6-byte padding', async () => {
    const naturalPath = join(directory, 'natural.xlsx');
    const natural = await generateSyntheticXlsx(naturalPath, { rows: 1, columns: 1 });

    for (const target of [natural.worksheetBytes - 1, ...Array.from(
      { length: 6 },
      (_, index) => natural.worksheetBytes + index + 1,
    )]) {
      const rejectedPath = join(directory, `rejected-${target}.xlsx`);
      await expect(generateSyntheticXlsx(rejectedPath, {
        rows: 1,
        columns: 1,
        targetWorksheetBytes: target,
      })).rejects.toThrow(/targetWorksheetBytes/);
      await expect(access(rejectedPath)).rejects.toThrow();
    }

    const exactPath = join(directory, 'exact.xlsx');
    const exact = await generateSyntheticXlsx(exactPath, {
      rows: 1,
      columns: 1,
      targetWorksheetBytes: natural.worksheetBytes + 7,
    });
    expect(exact.worksheetBytes).toBe(natural.worksheetBytes + 7);
    expect((await stat(exactPath)).size).toBeGreaterThan(0);
  });

  it('enforces the SpreadsheetML grid bounds before creating output', async () => {
    for (const options of [
      { rows: 1_048_577, columns: 1 },
      { rows: 1, columns: 16_385 },
    ]) {
      const output = join(directory, `${options.rows}-${options.columns}.xlsx`);
      await expect(generateSyntheticXlsx(output, options)).rejects.toThrow(/must be at most/);
      await expect(access(output)).rejects.toThrow();
    }
  });

  it('omits the optional shared-string occurrence count instead of publishing a false count', async () => {
    const output = join(directory, 'shared-count.xlsx');
    await generateSyntheticXlsx(output, { rows: 2, columns: 5 });
    const sharedStrings = extractZipEntry(await readFile(output), 'xl/sharedStrings.xml')
      .toString('utf8');
    const openingTag = sharedStrings.match(/<sst\b[^>]*>/u)?.[0];
    expect(openingTag).toContain('uniqueCount="2"');
    expect(openingTag).not.toMatch(/\scount=/u);
  });

  it('awaits close, removes partial output, and preserves an injected post-write error', async () => {
    const output = join(directory, 'partial.xlsx');
    const injected = new Error('injected after successful ZIP write');
    const generateWithFailure = generateSyntheticXlsx as unknown as (
      path: string,
      options: { rows: number; columns: number },
      testing: { failAfterWrites: number; failure: Error },
    ) => Promise<unknown>;
    await expect(generateWithFailure(
      output,
      { rows: 1, columns: 1 },
      { failAfterWrites: 2, failure: injected },
    )).rejects.toBe(injected);
    await expect(access(output)).rejects.toThrow();
  });
});

function extractZipEntry(zip: Buffer, expectedName: string): Buffer {
  let eocd = zip.length - 22;
  while (eocd >= 0 && zip.readUInt32LE(eocd) !== 0x06054b50) eocd--;
  if (eocd < 0) throw new Error('ZIP end-of-central-directory is missing');
  const entries = zip.readUInt16LE(eocd + 10);
  let cursor = zip.readUInt32LE(eocd + 16);
  for (let index = 0; index < entries; index++) {
    if (zip.readUInt32LE(cursor) !== 0x02014b50) throw new Error('invalid ZIP central entry');
    const method = zip.readUInt16LE(cursor + 10);
    const compressedSize = zip.readUInt32LE(cursor + 20);
    const nameLength = zip.readUInt16LE(cursor + 28);
    const extraLength = zip.readUInt16LE(cursor + 30);
    const commentLength = zip.readUInt16LE(cursor + 32);
    const localOffset = zip.readUInt32LE(cursor + 42);
    const name = zip.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8');
    if (name === expectedName) {
      if (zip.readUInt32LE(localOffset) !== 0x04034b50) throw new Error('invalid ZIP local entry');
      const localNameLength = zip.readUInt16LE(localOffset + 26);
      const localExtraLength = zip.readUInt16LE(localOffset + 28);
      const start = localOffset + 30 + localNameLength + localExtraLength;
      const compressed = zip.subarray(start, start + compressedSize);
      if (method === 0) return Buffer.from(compressed);
      if (method === 8) return inflateRawSync(compressed);
      throw new Error(`unsupported ZIP compression method ${method}`);
    }
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  throw new Error(`ZIP entry not found: ${expectedName}`);
}
