import type { OoxmlResourceUsageSnapshot, WorkerBridgeTransport } from '@silurus/ooxml-core';
import {
  BoundedPullSession,
  HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
  HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
  PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
  PULL_SESSION_PROTOCOL,
  type PullSessionIdentity,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import type { BodyElement, DocxDocumentModel } from './types.js';

export const DOCX_INITIAL_BODY_PULL_BYTES = 1024 * 1024;
const MAX_DOCUMENT_UNIT_BYTES = Math.max(
  HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
  HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
);

type DocumentUnit =
  | { readonly kind: 'body'; readonly body: BodyElement[] }
  | { readonly kind: 'complete'; readonly document: DocxDocumentModel };

export interface MaterializeDocumentPullOptions {
  readonly timeoutMs?: number;
  readonly signal?: AbortSignal;
  readonly onUsage?: (usage: OoxmlResourceUsageSnapshot) => void;
}

/** Drain one sequential DOCX operation into the backward-compatible public
 * model. Each transferred JSON unit is decoded, accepted, and ACKed before the
 * worker may produce another; no monolithic document JSON crosses realms. */
export async function materializeDocumentPullSession(
  transport: WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>,
  identity: PullSessionIdentity<number>,
  options: MaterializeDocumentPullOptions = {},
): Promise<DocxDocumentModel> {
  const session = new BoundedPullSession(transport, {
    ...identity,
    maxByteCredit: MAX_DOCUMENT_UNIT_BYTES,
    timeoutMs: options.timeoutMs,
  });
  const body: BodyElement[] = [];
  try {
    for (;;) {
      const chunk = await pullWithCreditRetry(session, options.signal);
      try {
        const usage = chunk.usage ?? session.usageCheckpoint;
        if (usage) options.onUsage?.(usage);
        const unit = parseDocumentUnit(chunk.payload);
        if (chunk.done !== (unit.kind === 'complete')) {
          throw new TypeError('DOCX document unit terminal flag does not match its payload');
        }
        if (unit.kind === 'body') {
          body.push(...unit.body);
          await chunk.ack({ signal: options.signal });
          continue;
        }
        if (!Array.isArray(unit.document.body) || unit.document.body.length !== 0) {
          throw new TypeError('DOCX terminal document must not duplicate streamed body blocks');
        }
        unit.document.body = body;
        await chunk.ack({ signal: options.signal });
        return unit.document;
      } finally {
        chunk.disposeTransferred();
      }
    }
  } catch (error) {
    await session.cancel('request-error').catch(() => undefined);
    throw error;
  }
}

export function isDocumentPullResponse(
  value: unknown,
): value is PullSessionResponse<ArrayBuffer, number> {
  return !!value && typeof value === 'object'
    && (value as { protocol?: unknown }).protocol === PULL_SESSION_PROTOCOL;
}

function parseDocumentUnit(payload: ArrayBuffer): DocumentUnit {
  const value = JSON.parse(new TextDecoder().decode(new Uint8Array(payload))) as unknown;
  if (!value || typeof value !== 'object') throw new TypeError('DOCX document unit must be an object');
  const record = value as Record<string, unknown>;
  if (record.kind === 'body' && Array.isArray(record.body)) {
    return record as unknown as Extract<DocumentUnit, { kind: 'body' }>;
  }
  if (record.kind === 'complete' && record.document && typeof record.document === 'object') {
    return record as unknown as Extract<DocumentUnit, { kind: 'complete' }>;
  }
  throw new TypeError('DOCX document unit has an unknown shape');
}

async function pullWithCreditRetry(
  session: BoundedPullSession<ArrayBuffer, number>,
  signal?: AbortSignal,
) {
  try {
    return await session.pull(DOCX_INITIAL_BODY_PULL_BYTES, { signal });
  } catch (error) {
    const required = requiredCredit(error);
    if (required === undefined) throw error;
    return session.pull(required, { signal });
  }
}

const REQUIRED_CREDIT = /^document unit requires ([0-9]+) bytes but credit is ([0-9]+)$/u;

function requiredCredit(error: unknown): number | undefined {
  if (!error || typeof error !== 'object'
    || (error as { code?: unknown }).code !== PULL_SESSION_INSUFFICIENT_CREDIT_CODE) {
    return undefined;
  }
  const match = REQUIRED_CREDIT.exec(error instanceof Error ? error.message : String(error));
  if (!match) return undefined;
  const required = Number(match[1]);
  const offered = Number(match[2]);
  if (
    !Number.isSafeInteger(required) || required <= offered
    || required > MAX_DOCUMENT_UNIT_BYTES
    || offered !== DOCX_INITIAL_BODY_PULL_BYTES
    || String(required) !== match[1] || String(offered) !== match[2]
  ) return undefined;
  return required;
}
