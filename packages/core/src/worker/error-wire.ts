import {
  OoxmlError,
  OoxmlResourceLimitError,
  type OoxmlErrorCode,
  type OoxmlFormat,
  type OoxmlResourceLimitErrorDetails,
  type OoxmlResourceUsageSnapshot,
  type OoxmlResourceViolation,
} from '../errors/ooxml-error.js';

const RESOURCE_LIMIT_PREFIX = 'OOXML_RESOURCE_LIMIT:';

/** Structured-clone-safe error payload shared by all OOXML workers. */
export interface WorkerErrorPayload {
  message: string;
  errorName?: string;
  code?: string;
  resourceLimit?: OoxmlResourceLimitErrorDetails;
}

interface RustResourceLimitPayload {
  code: 'ooxml-resource-limit';
  details: OoxmlResourceLimitErrorDetails;
}

function isNonNegativeSafeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function isUsage(value: unknown): value is OoxmlResourceUsageSnapshot {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const usage = value as Partial<OoxmlResourceUsageSnapshot>;
  return (
    isNonNegativeSafeInteger(usage.archiveEntryCount) &&
    isNonNegativeSafeInteger(usage.declaredInflatedBytes) &&
    isNonNegativeSafeInteger(usage.distinctInflatedBytes) &&
    isNonNegativeSafeInteger(usage.operationInflatedBytes)
  );
}

function isFormat(value: unknown): value is OoxmlFormat {
  return value === 'docx' || value === 'xlsx' || value === 'pptx';
}

function isViolation(value: unknown): value is OoxmlResourceViolation {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const data = value as Partial<OoxmlResourceViolation>;
  if (
    !isFormat(data.format) ||
    typeof data.operation !== 'string' ||
    data.operation.length === 0 ||
    !isNonNegativeSafeInteger(data.limit) ||
    !isNonNegativeSafeInteger(data.observed) ||
    typeof data.configurable !== 'boolean' ||
    !isUsage(data.usage)
  ) {
    return false;
  }
  if (data.resource === 'archive-entry') {
    return (
      (data.metric === 'declared-inflated-bytes' ||
        data.metric === 'actual-inflated-bytes') &&
      typeof data.part === 'string' &&
      data.part.length > 0
    );
  }
  if (data.resource !== 'archive') return false;
  if (data.metric === 'entry-count') {
    return data.configurable === false && !('part' in data);
  }
  return (
    data.metric === 'distinct-inflated-bytes' &&
    typeof data.part === 'string' &&
    data.part.length > 0
  );
}

function isResourceLimitDetails(value: unknown): value is OoxmlResourceLimitErrorDetails {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const details = value as Partial<OoxmlResourceLimitErrorDetails>;
  if (!isViolation(details.violation)) return false;
  const metric = details.violation.metric;
  const expectedStage =
    metric === 'actual-inflated-bytes' || metric === 'distinct-inflated-bytes'
      ? 'decompression'
      : 'container';
  return details.stage === expectedStage;
}

function resourceLimitMessage(details: OoxmlResourceLimitErrorDetails): string {
  const violation = details.violation;
  const location = 'part' in violation && violation.part ? ` for ${violation.part}` : '';
  return `OOXML resource limit exceeded${location}: ${violation.metric} ${violation.observed} > ${violation.limit}`;
}

/** Parse only an exact Rust resource-limit envelope, never a wrapped substring. */
export function parseResourceLimitError(error: unknown): OoxmlResourceLimitError | undefined {
  const text = error instanceof Error ? error.message : String(error);
  if (!text.startsWith(RESOURCE_LIMIT_PREFIX)) return undefined;
  let value: unknown;
  try {
    value = JSON.parse(text.slice(RESOURCE_LIMIT_PREFIX.length));
  } catch {
    return undefined;
  }
  if (!value || typeof value !== 'object') return undefined;
  const data = value as Partial<RustResourceLimitPayload>;
  if (data.code !== 'ooxml-resource-limit' || !isResourceLimitDetails(data.details)) {
    return undefined;
  }
  return new OoxmlResourceLimitError(resourceLimitMessage(data.details), data.details);
}

/** Convert an arbitrary worker-side error to a structured-clone-safe payload. */
export function serializeWorkerError(error: unknown): WorkerErrorPayload {
  const typed =
    error instanceof OoxmlError || error instanceof OoxmlResourceLimitError
      ? error
      : parseResourceLimitError(error);
  if (typed instanceof OoxmlResourceLimitError) {
    return {
      message: typed.message,
      errorName: typed.name,
      code: typed.code,
      resourceLimit: typed.details,
    };
  }
  if (typed instanceof OoxmlError) {
    return { message: typed.message, errorName: typed.name, code: typed.code };
  }
  const ordinary = error instanceof Error ? error : new Error(String(error));
  const details = ordinary as Error & { code?: string };
  return {
    message: ordinary.message,
    errorName: ordinary.name,
    ...(details.code !== undefined ? { code: details.code } : {}),
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
    payload.code === 'ooxml-resource-limit' &&
    isResourceLimitDetails(payload.resourceLimit)
  ) {
    return new OoxmlResourceLimitError(payload.message, payload.resourceLimit);
  }
  if (payload.code && OOXML_ERROR_CODES.has(payload.code as OoxmlErrorCode)) {
    return new OoxmlError(payload.code as OoxmlErrorCode, payload.message);
  }
  const error =
    payload.errorName === 'TypeError'
      ? new TypeError(payload.message)
      : payload.errorName === 'RangeError'
        ? new RangeError(payload.message)
        : new Error(payload.message);
  if (payload.errorName) error.name = payload.errorName;
  if (payload.code !== undefined) Object.assign(error, { code: payload.code });
  return error;
}
