import init, { parse_docx } from './wasm/docx_parser.js';
import { decodeDataUrl } from '@silurus/ooxml-core';
import type { WorkerRequest, WorkerResponse } from './types';

let initPromise: Promise<unknown> | null = null;

self.onmessage = async (e: MessageEvent<WorkerRequest>) => {
  const req = e.data;

  if (req.type === 'init') {
    initPromise = init(decodeDataUrl(req.wasmUrl) ?? req.wasmUrl);
    return;
  }

  // Echo the correlation id so the client routes the response to the right
  // pending promise (id correlation, not response-type matching).
  const id = req.id;
  try {
    await initPromise;
    if (req.type === 'parse') {
      const maxBytes =
        typeof req.maxZipEntryBytes === 'number' && req.maxZipEntryBytes > 0
          ? BigInt(req.maxZipEntryBytes)
          : undefined;
      const json = parse_docx(new Uint8Array(req.data), maxBytes);
      const document = JSON.parse(json);
      if (document.error) throw new Error(`Parse error: ${document.error}`);
      const res: WorkerResponse = { type: 'parsed', id, document };
      self.postMessage(res);
    }
  } catch (err) {
    const res: WorkerResponse = { type: 'error', id, message: String(err) };
    self.postMessage(res);
  }
};
