/**
 * Machine-readable code for a typed load-time failure.
 *
 * The container-level failures the `load()` factories detect on the main thread
 * before handing bytes to the parser worker (see `sniffCfb` / `decryptOoxml`).
 * This is the seed of the broader typed-error surface tracked as PD4 (OoxmlError
 * typed errors). Add codes here rather than throwing bare `Error(string)`, so
 * callers can `switch` on `err.code` instead of matching message text.
 *
 *   - `'encrypted'`             — password-protected, but no `password` was
 *     supplied (pass `LoadOptions.password` to decrypt).
 *   - `'invalid-password'`      — a `password` was supplied but did not match.
 *   - `'unsupported-encryption'`— encrypted with a scheme other than Agile
 *     (Standard / Extensible / a legacy binary encryptor), which this library
 *     cannot decrypt (PD8 implements Agile only).
 *   - `'legacy-binary-format'`  — a raw .doc / .xls / .ppt (not OOXML).
 *   - `'not-ooxml'`             — a CFB of an unrecognised kind, or otherwise
 *     not an OOXML ZIP.
 */
export type OoxmlErrorCode =
  | 'encrypted'
  | 'invalid-password'
  | 'unsupported-encryption'
  | 'legacy-binary-format'
  | 'not-ooxml';

export type OoxmlErrorStage =
  | 'container'
  | 'decompression'
  | 'parsing'
  | 'serialization'
  | 'layout'
  | 'rendering'
  | 'worker';

export type OoxmlErrorSource =
  | 'container'
  | 'zip-part'
  | 'parser'
  | 'serializer'
  | 'layout'
  | 'renderer'
  | 'worker';

/**
 * Typed error thrown by the docx / pptx / xlsx `load()` factories for failures
 * that carry a stable, programmatic {@link OoxmlErrorCode} (e.g. a
 * password-protected or legacy-binary file detected from its container magic).
 *
 * Note on workers: `instanceof OoxmlError` does not survive a structured-clone
 * across the worker boundary. Detection that needs a typed error is therefore
 * done on the main thread (before the worker is involved) so a genuine
 * `OoxmlError` instance is thrown to the caller. Errors that must cross the
 * worker boundary should carry the `code` string and be reconstructed on the
 * main side.
 */
export class OoxmlError extends Error {
  readonly code: OoxmlErrorCode;

  constructor(code: OoxmlErrorCode, message: string) {
    super(message);
    this.name = 'OoxmlError';
    this.code = code;
    // Restore the prototype chain for environments that down-level `extends
    // Error` (e.g. older TS `target`), so `instanceof OoxmlError` holds.
    Object.setPrototypeOf(this, OoxmlError.prototype);
  }
}

export type OoxmlFormat = 'docx' | 'xlsx' | 'pptx';

export interface OoxmlResourceUsageSnapshot {
  readonly archiveEntryCount: number;
  readonly declaredInflatedBytes: number;
  readonly distinctInflatedBytes: number;
  readonly operationInflatedBytes: number;
}

interface OoxmlResourceViolationBase {
  format: OoxmlFormat;
  operation: string;
  limit: number;
  observed: number;
  configurable: boolean;
  usage: OoxmlResourceUsageSnapshot;
}

export type OoxmlResourceViolation =
  | (OoxmlResourceViolationBase & {
      resource: 'archive-entry';
      metric: 'declared-inflated-bytes' | 'actual-inflated-bytes';
      part: string;
    })
  | (OoxmlResourceViolationBase & {
      resource: 'archive';
      metric: 'entry-count';
      configurable: false;
    })
  | (OoxmlResourceViolationBase & {
      resource: 'archive';
      metric: 'distinct-inflated-bytes';
      part: string;
    });

export interface OoxmlResourceLimitErrorDetails {
  readonly stage: 'container' | 'decompression';
  readonly violation: OoxmlResourceViolation;
}

/** Deterministic rejection caused by a measured OOXML resource-policy breach. */
export class OoxmlResourceLimitError extends Error {
  readonly code = 'ooxml-resource-limit' as const;
  readonly details: OoxmlResourceLimitErrorDetails;

  constructor(message: string, details: OoxmlResourceLimitErrorDetails) {
    super(message);
    this.name = 'OoxmlResourceLimitError';
    const violation = {
      ...details.violation,
      usage: Object.freeze({ ...details.violation.usage }),
    } as OoxmlResourceViolation;
    this.details = Object.freeze({
      ...details,
      violation: Object.freeze(violation),
    });
    Object.setPrototypeOf(this, OoxmlResourceLimitError.prototype);
  }
}
