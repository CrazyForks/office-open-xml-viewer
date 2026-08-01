#!/usr/bin/env node
import { once } from 'node:events';
import { createWriteStream } from 'node:fs';
import { rm } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Readable, Transform } from 'node:stream';
import { createDeflateRaw, crc32 } from 'node:zlib';
import { pathToFileURL } from 'node:url';

const encoder = new TextEncoder();

/**
 * Generate a deterministic, standards-valid single-sheet XLSX without ever
 * retaining the worksheet XML or ZIP in memory. `targetWorksheetBytes` pads
 * with individually small XML comments, avoiding one giant lexical event.
 */
export async function generateSyntheticXlsx(outputPath, options = {}, testing = {}) {
  const rows = boundedPositiveInteger(options.rows ?? 2049, 'rows', 1_048_576);
  const columns = boundedPositiveInteger(options.columns ?? 32, 'columns', 16_384);
  const targetWorksheetBytes = options.targetWorksheetBytes === undefined
    ? undefined
    : positiveInteger(options.targetWorksheetBytes, 'targetWorksheetBytes');
  const malformedWorksheetTail = options.malformedWorksheetTail === true;
  const naturalWorksheetBytes = await measureWorksheetXml(rows, columns, malformedWorksheetTail);
  if (targetWorksheetBytes !== undefined) {
    if (targetWorksheetBytes < naturalWorksheetBytes) {
      throw new RangeError(
        `targetWorksheetBytes ${targetWorksheetBytes} is smaller than natural worksheet XML length ${naturalWorksheetBytes}`,
      );
    }
    const padding = targetWorksheetBytes - naturalWorksheetBytes;
    if (padding > 0 && padding < 7) {
      throw new RangeError('targetWorksheetBytes requires an impossible 1..6-byte XML comment remainder');
    }
  }
  const output = createWriteStream(outputPath);
  let outputError;
  let successfulWrites = 0;
  output.on('error', (error) => {
    outputError = error;
  });
  let offset = 0;
  const central = [];

  try {

  const write = async (chunk) => {
    if (outputError) throw outputError;
    if (!output.write(chunk)) await once(output, 'drain');
    if (outputError) throw outputError;
    offset += chunk.length;
    successfulWrites++;
    if (successfulWrites === testing.failAfterWrites) {
      throw testing.failure ?? new Error('injected synthetic XLSX writer failure');
    }
  };

  const addEntry = async (name, source) => {
    const nameBytes = Buffer.from(name);
    const localOffset = offset;
    const local = Buffer.alloc(30 + nameBytes.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x08, 6); // sizes and CRC follow in a data descriptor
    local.writeUInt16LE(8, 8); // deflate
    local.writeUInt16LE(nameBytes.length, 26);
    nameBytes.copy(local, 30);
    await write(local);

    let checksum = 0;
    let uncompressedSize = 0;
    let compressedSize = 0;
    const meter = new Transform({
      transform(chunk, _encoding, callback) {
        checksum = crc32(chunk, checksum) >>> 0;
        uncompressedSize += chunk.length;
        callback(null, chunk);
      },
    });
    const compressed = Readable.from(source).pipe(meter).pipe(createDeflateRaw({ level: 6 }));
    for await (const chunk of compressed) {
      compressedSize += chunk.length;
      await write(chunk);
    }
    const descriptor = Buffer.alloc(16);
    descriptor.writeUInt32LE(0x08074b50, 0);
    descriptor.writeUInt32LE(checksum, 4);
    descriptor.writeUInt32LE(compressedSize, 8);
    descriptor.writeUInt32LE(uncompressedSize, 12);
    await write(descriptor);
    central.push({ nameBytes, checksum, compressedSize, uncompressedSize, localOffset });
    return uncompressedSize;
  };

  await addEntry('[Content_Types].xml', strings(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`));
  await addEntry('_rels/.rels', strings(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`));
  await addEntry('xl/workbook.xml', strings(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Synthetic" sheetId="1" r:id="rId1"/></sheets></workbook>`));
  await addEntry('xl/_rels/workbook.xml.rels', strings(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`));
  await addEntry('xl/styles.xml', strings(`<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`));
  await addEntry('xl/sharedStrings.xml', strings(`<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="2"><si><t>shared-alpha</t></si><si><t>shared-beta</t></si></sst>`));
  const worksheetBytes = await addEntry(
    'xl/worksheets/sheet1.xml',
    worksheetXml(rows, columns, targetWorksheetBytes, malformedWorksheetTail),
  );

  const centralOffset = offset;
  for (const entry of central) {
    const header = Buffer.alloc(46 + entry.nameBytes.length);
    header.writeUInt32LE(0x02014b50, 0);
    header.writeUInt16LE(20, 4);
    header.writeUInt16LE(20, 6);
    header.writeUInt16LE(0x08, 8);
    header.writeUInt16LE(8, 10);
    header.writeUInt32LE(entry.checksum, 16);
    header.writeUInt32LE(entry.compressedSize, 20);
    header.writeUInt32LE(entry.uncompressedSize, 24);
    header.writeUInt16LE(entry.nameBytes.length, 28);
    header.writeUInt32LE(entry.localOffset, 42);
    entry.nameBytes.copy(header, 46);
    await write(header);
  }
  const centralSize = offset - centralOffset;
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(central.length, 8);
  end.writeUInt16LE(central.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(centralOffset, 16);
  await write(end);
  output.end();
  await once(output, 'close');
  if (outputError) throw outputError;
  return { outputPath, rows, columns, cells: rows * columns, worksheetBytes, zipBytes: offset };
  } catch (error) {
    output.destroy();
    if (!output.closed) await once(output, 'close').catch(() => undefined);
    await rm(outputPath, { force: true }).catch(() => undefined);
    throw error;
  }
}

async function* worksheetXml(rows, columns, targetBytes, malformedTail = false) {
  let emitted = 0;
  const emit = (text) => {
    const bytes = encoder.encode(text);
    emitted += bytes.length;
    return bytes;
  };
  yield emit(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${columnName(columns)}${rows}"/><sheetData>`);
  for (let row = 1; row <= rows; row++) {
    let xml = `<row r="${row}">`;
    for (let column = 1; column <= columns; column++) {
      const ref = `${columnName(column)}${row}`;
      switch ((row + column) % 5) {
        case 0: xml += `<c r="${ref}"><v>${row * columns + column}</v></c>`; break;
        case 1: xml += `<c r="${ref}" t="s"><v>${(row + column) % 2}</v></c>`; break;
        case 2: xml += `<c r="${ref}" t="inlineStr"><is><t>r${row}c${column}</t></is></c>`; break;
        case 3: xml += `<c r="${ref}" t="b"><v>${(row + column) % 2}</v></c>`; break;
        default: xml += `<c r="${ref}"><f>${ref}+1</f><v>${row + column}</v></c>`; break;
      }
    }
    yield emit(`${xml}</row>`);
  }
  const suffix = malformedTail ? '</sheetData><broken></worksheet>' : '</sheetData></worksheet>';
  const naturalBytes = emitted + Buffer.byteLength(suffix);
  const padding = targetBytes === undefined ? 0 : targetBytes - naturalBytes;
  if (padding < 0) throw new RangeError('targetWorksheetBytes is below the natural worksheet length');
  for (const size of commentSizes(padding)) yield emit(`<!--${'p'.repeat(size - 7)}-->`);
  yield emit(suffix);
}

async function measureWorksheetXml(rows, columns, malformedTail) {
  let bytes = 0;
  for await (const chunk of worksheetXml(rows, columns, undefined, malformedTail)) {
    bytes += chunk.byteLength;
  }
  return bytes;
}

function commentSizes(total) {
  if (total === 0) return [];
  if (total < 7) throw new RangeError('targetWorksheetBytes leaves less than seven bytes of padding');
  const sizes = [];
  let remaining = total;
  while (remaining > 4096) {
    const size = remaining - 4096 < 7 ? 4089 : 4096;
    sizes.push(size);
    remaining -= size;
  }
  if (remaining) sizes.push(remaining);
  return sizes;
}

function columnName(index) {
  let name = '';
  while (index > 0) {
    index--;
    name = String.fromCharCode(65 + (index % 26)) + name;
    index = Math.floor(index / 26);
  }
  return name;
}

async function* strings(value) {
  yield encoder.encode(value);
}

function positiveInteger(value, name) {
  if (!Number.isSafeInteger(value) || value <= 0) throw new RangeError(`${name} must be a positive safe integer`);
  return value;
}

function boundedPositiveInteger(value, name, maximum) {
  const integer = positiveInteger(value, name);
  if (integer > maximum) throw new RangeError(`${name} must be at most ${maximum}`);
  return integer;
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  const args = Object.fromEntries(process.argv.slice(2).map((arg) => {
    const [key, value] = arg.replace(/^--/, '').split('=', 2);
    return [key, value];
  }));
  if (!args.output) throw new Error('usage: generate-synthetic-xlsx.mjs --output=FILE [--rows=2049] [--columns=32] [--target-worksheet-bytes=N]');
  const result = await generateSyntheticXlsx(resolve(args.output), {
    rows: args.rows === undefined ? undefined : Number(args.rows),
    columns: args.columns === undefined ? undefined : Number(args.columns),
    targetWorksheetBytes: args['target-worksheet-bytes'] === undefined
      ? undefined
      : Number(args['target-worksheet-bytes']),
  });
  process.stdout.write(`${JSON.stringify(result)}\n`);
}
