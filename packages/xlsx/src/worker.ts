import init, { parse_xlsx, parse_sheet } from './wasm/xlsx_parser.js';
import { decodeDataUrl } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types.js';

let initPromise: Promise<unknown> | null = null;

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.type === 'init') {
    initPromise = init(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  // Every non-init request carries a correlation id that must be echoed back so
  // the client can route the response to the right pending promise.
  const id = req.id;
  try {
    await initPromise;
    if (req.type === 'parse') {
      const maxBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      const json = parse_xlsx(new Uint8Array(req.data), maxBytes);
      const workbook = JSON.parse(json);
      const res: WorkerResponse = { type: 'parsed', id, workbook };
      self.postMessage(res);
    } else if (req.type === 'parseSheet') {
      const maxBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      const json = parse_sheet(new Uint8Array(req.data), req.sheetIndex, req.sheetName, maxBytes);
      const worksheet = JSON.parse(json);
      const res: WorkerResponse = { type: 'parsedSheet', id, worksheet };
      self.postMessage(res);
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', id, message: String(err) };
    self.postMessage(res);
  }
};
