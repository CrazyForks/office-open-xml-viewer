import type {
  OoxmlResourceLimits,
  OoxmlResourceMetrics,
} from '@silurus/ooxml-core';

/** Resource policy, diagnostics, and cancellation shared by every Node session. */
export interface OoxmlNodeSessionOptions {
  /** Package-level inflated ZIP admission limits. */
  resourceLimits?: OoxmlResourceLimits;
  /** @deprecated Use `resourceLimits.maxArchiveEntryBytes`. */
  maxZipEntryBytes?: number;
  /** Emit one content-free resource report for the terminal session outcome. */
  debug?: boolean;
  /** Receive the same terminal report without enabling console output. */
  onResourceMetrics?: (metrics: OoxmlResourceMetrics) => void;
  /** Cooperatively abort initialization or active work; synchronous WASM cannot be preempted. */
  signal?: AbortSignal;
}
