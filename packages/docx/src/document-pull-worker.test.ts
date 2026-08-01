import { describe, expect, it } from 'vitest';
import type {
  PullSessionIdentity,
} from '@silurus/ooxml-core/worker';
import {
  materializeDocumentPullSession,
} from './document-pull-client.js';
import {
  createLocalDocumentPullTransport,
  DocumentPullWorker,
  MaterializedDocumentCursorArchive,
  type DocxDocumentCursorArchive,
} from './document-pull-worker.js';
import type { DocxDocumentModel } from './types.js';

const encoder = new TextEncoder();

class FakeArchive implements DocxDocumentCursorArchive {
  private sequence = 0;
  private delivered = false;
  private currentDone = false;
  opened = false;
  canceled = false;
  readonly units: Uint8Array[];

  constructor(units?: unknown[]) {
    this.units = (units ?? [
      { kind: 'body', body: [{ type: 'pageBreak' }] },
      { kind: 'body', body: [{ type: 'columnBreak' }] },
      { kind: 'complete', document: { body: [] } },
    ]).map((unit) => encoder.encode(JSON.stringify(unit)));
  }

  open_document_cursor(): void { this.opened = true; }

  pull_document_chunk(sequence: number, _operation: number, _generation: number, credit: number): Uint8Array {
    if (!this.opened) throw new Error('not open');
    if (sequence !== this.sequence) throw new Error('bad sequence');
    if (this.delivered) throw new Error('document unit must be acknowledged before another pull');
    const bytes = this.units[sequence]!;
    if (bytes.byteLength > credit) {
      throw new Error(`document unit requires ${bytes.byteLength} bytes but credit is ${credit}`);
    }
    this.delivered = true;
    this.currentDone = sequence + 1 === this.units.length;
    return bytes.slice();
  }

  document_chunk_done(): boolean { return this.currentDone; }

  acknowledge_document_chunk(sequence: number): void {
    if (!this.delivered || sequence !== this.sequence) throw new Error('bad acknowledgement');
    this.delivered = false;
    this.sequence += 1;
  }

  cancel_document_cursor(): void { this.canceled = true; }
  close_document_session(): void { this.canceled = true; }
}

const identity: PullSessionIdentity<number> = {
  sessionId: 1,
  operationId: 1,
  generation: 1,
};

describe('DOCX document pull integration', () => {
  it('materializes acknowledged body units and a body-free terminal envelope', async () => {
    const archive = new FakeArchive();
    const worker = new DocumentPullWorker(() => archive);
    worker.open(identity);

    const document = await materializeDocumentPullSession(
      createLocalDocumentPullTransport(worker),
      identity,
    );

    expect(document.body.map((element) => element.type)).toEqual(['pageBreak', 'columnBreak']);
    expect(archive.canceled).toBe(false);
  });

  it('cancels producer ownership when consumer validation rejects a unit', async () => {
    const archive = new FakeArchive([{ kind: 'unexpected' }]);
    const worker = new DocumentPullWorker(() => archive);
    worker.open(identity);

    await expect(
      materializeDocumentPullSession(createLocalDocumentPullTransport(worker), identity),
    ).rejects.toThrow('unknown shape');
    expect(archive.canceled).toBe(true);
  });

  it('streams an already-materialized fallback model without a whole-model envelope', async () => {
    const source = {
      section: {},
      body: [{ type: 'pageBreak' }, { type: 'columnBreak' }],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
    } as unknown as DocxDocumentModel;
    const archive = new MaterializedDocumentCursorArchive(source);
    const worker = new DocumentPullWorker(() => archive);
    worker.open(identity);

    // Ownership moves into the bounded cursor immediately; the compatibility
    // model returned to the consumer is reconstructed from transferred units.
    expect(source.body).toEqual([]);
    const document = await materializeDocumentPullSession(
      createLocalDocumentPullTransport(worker),
      identity,
    );

    expect(document).not.toBe(source);
    expect(document.body.map((element) => element.type)).toEqual(['pageBreak', 'columnBreak']);
  });
});
