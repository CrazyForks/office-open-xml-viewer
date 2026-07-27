import {
  OoxmlError,
  ParserResourceLimitError,
  type OoxmlErrorCode,
  type OoxmlErrorSource,
  type OoxmlErrorStage,
} from '../errors/ooxml-error';

const RESOURCE_LIMIT_PREFIX = 'OOXML_RESOURCE_LIMIT:';

/** Structured-clone-safe error payload shared by all OOXML workers. */
export interface WorkerErrorPayload {
  message: string;
  errorName?: string;
  code?: string;
  stage?: OoxmlErrorStage;
  source?: OoxmlErrorSource;
  part?: string;
  metric?: string;
  limit?: number;
  observed?: number;
}

interface RustResourceLimitPayload {
  code: 'parser-resource-limit';
  stage: OoxmlErrorStage;
  source: OoxmlErrorSource;
  part: string;
  metric: string;
  limit: number;
  observed: number;
}

function isFiniteNonNegativeNumber(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0;
}

const OOXML_ERROR_STAGES = new Set<OoxmlErrorStage>([
  'container',
  'decompression',
  'parsing',
  'serialization',
  'layout',
  'rendering',
  'worker',
]);

const OOXML_ERROR_SOURCES = new Set<OoxmlErrorSource>([
  'container',
  'zip-part',
  'parser',
  'serializer',
  'layout',
  'renderer',
  'worker',
]);

function isOoxmlErrorStage(value: unknown): value is OoxmlErrorStage {
  return typeof value === 'string' && OOXML_ERROR_STAGES.has(value as OoxmlErrorStage);
}

function isOoxmlErrorSource(value: unknown): value is OoxmlErrorSource {
  return typeof value === 'string' && OOXML_ERROR_SOURCES.has(value as OoxmlErrorSource);
}

/** Parse the Rust ZIP reader's deterministic resource-limit envelope. */
export function parseResourceLimitError(error: unknown): ParserResourceLimitError | undefined {
  const text = error instanceof Error ? error.message : String(error);
  const start = text.indexOf(RESOURCE_LIMIT_PREFIX);
  if (start < 0) return undefined;
  const json = text.slice(start + RESOURCE_LIMIT_PREFIX.length);
  let value: unknown;
  try {
    value = JSON.parse(json);
  } catch {
    return undefined;
  }
  if (!value || typeof value !== 'object') return undefined;
  const data = value as Partial<RustResourceLimitPayload>;
  if (
    data.code !== 'parser-resource-limit' ||
    !isOoxmlErrorStage(data.stage) ||
    !isOoxmlErrorSource(data.source) ||
    typeof data.part !== 'string' ||
    typeof data.metric !== 'string' ||
    !isFiniteNonNegativeNumber(data.limit) ||
    !isFiniteNonNegativeNumber(data.observed)
  ) {
    return undefined;
  }
  return new ParserResourceLimitError(
    `Parser resource limit exceeded for ${data.part}: ${data.metric} ${data.observed} > ${data.limit}`,
    {
      stage: data.stage,
      source: data.source,
      part: data.part,
      metric: data.metric,
      limit: data.limit,
      observed: data.observed,
    },
  );
}

/** Convert an arbitrary worker-side error to a structured-clone-safe payload. */
export function serializeWorkerError(error: unknown): WorkerErrorPayload {
  const typed =
    error instanceof OoxmlError || error instanceof ParserResourceLimitError
      ? error
      : parseResourceLimitError(error);
  if (typed) {
    return {
      message: typed.message,
      errorName: typed.name,
      code: typed.code,
      ...(typed instanceof ParserResourceLimitError
        ? {
            stage: typed.stage,
            source: typed.source,
            part: typed.part,
            metric: typed.metric,
            limit: typed.limit,
            observed: typed.observed,
          }
        : {}),
    };
  }
  const ordinary = error instanceof Error ? error : new Error(String(error));
  const details = ordinary as Error & Partial<WorkerErrorPayload>;
  return {
    message: ordinary.message,
    errorName: ordinary.name,
    ...(details.code !== undefined ? { code: details.code } : {}),
    ...(details.stage !== undefined ? { stage: details.stage } : {}),
    ...(details.source !== undefined ? { source: details.source } : {}),
    ...(details.part !== undefined ? { part: details.part } : {}),
    ...(details.metric !== undefined ? { metric: details.metric } : {}),
    ...(details.limit !== undefined ? { limit: details.limit } : {}),
    ...(details.observed !== undefined ? { observed: details.observed } : {}),
  };
}

const OOXML_ERROR_CODES = new Set<OoxmlErrorCode>([
  'encrypted',
  'invalid-password',
  'unsupported-encryption',
  'legacy-binary-format',
  'not-ooxml',
]);

/** Reconstruct a real Error subclass after a payload crosses a worker boundary. */
export function deserializeWorkerError(payload: WorkerErrorPayload): Error {
  if (
    payload.code === 'parser-resource-limit' &&
    isOoxmlErrorStage(payload.stage) &&
    isOoxmlErrorSource(payload.source) &&
    payload.part !== undefined &&
    payload.metric !== undefined &&
    payload.limit !== undefined &&
    payload.observed !== undefined
  ) {
    return new ParserResourceLimitError(payload.message, {
      stage: payload.stage,
      source: payload.source,
      part: payload.part,
      metric: payload.metric,
      limit: payload.limit,
      observed: payload.observed,
    });
  }
  if (payload.code && OOXML_ERROR_CODES.has(payload.code as OoxmlErrorCode)) {
    return new OoxmlError(payload.code as OoxmlErrorCode, payload.message);
  }
  const error = new Error(payload.message);
  if (payload.errorName) error.name = payload.errorName;
  if (payload.code !== undefined) Object.assign(error, { code: payload.code });
  return error;
}
