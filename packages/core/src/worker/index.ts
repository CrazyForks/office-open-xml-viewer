export {
  WorkerBridge,
  type WorkerLike,
  type WorkerBridgeOptions,
  type WorkerBridgeTransport,
  type WorkerRequestOptions,
} from './bridge.js';
export { decodeDataUrl } from './decode-data-url.js';
export {
  WasmParserHost,
  WasmTrapError,
  isWasmTrap,
  type WasmTrapErrorCode,
  type WasmInit,
  type WasmReinit,
  type WasmInitInput,
  type WasmParserHostOptions,
} from './wasm-guard.js';
export {
  decodeOoxmlResourceUsage,
  deserializeWorkerError,
  parseResourceLimitError,
  serializeWorkerError,
  type WorkerErrorPayload,
} from './error-wire.js';
export {
  DEFAULT_OOXML_RESOURCE_LIMITS,
  normalizeLoadResourceOptions,
  normalizeResourcePolicy,
  resourcePolicyForWasm,
  type NormalizedOoxmlResourceOptions,
  type NormalizedOoxmlResourcePolicy,
} from './resource-policy.js';
export {
  OoxmlResourceMetricsSession,
  type OoxmlResourceMetricsSessionOptions,
} from './resource-debug.js';
export type {
  OoxmlResourceMetrics,
  OoxmlResourceMetricsCheckpoint,
  OoxmlResourcePolicySnapshot,
} from '../types/resource-metrics.js';
export {
  emitOoxmlResourceDebugReport,
  formatOoxmlResourceDebugReport,
} from './resource-debug-view.js';
export {
  HARD_MAX_DOCX_BODY_BLOCK_XML_BYTES,
  HARD_MAX_DOCX_BODY_CHUNK_JSON_BYTES,
  HARD_MAX_DOCX_BOOTSTRAP_JSON_BYTES,
  HARD_MAX_DOCX_RETAINED_MODEL_JSON_BYTES,
  HARD_MAX_PPTX_CACHED_SLIDES,
  HARD_MAX_PPTX_CACHED_SLIDE_PROJECTION_BYTES,
  HARD_MAX_PPTX_MARKDOWN_BYTES,
  HARD_MAX_PPTX_RAW_PART_CACHE_BYTES,
  HARD_MAX_PPTX_RAW_PART_CACHE_ENTRIES,
  HARD_MAX_PPTX_SLIDE_JSON_BYTES,
  HARD_MAX_PPTX_PREFLIGHT_PROJECTION_BYTES,
  HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES,
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS,
  HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS,
  HARD_MAX_XLSX_WORKSHEET_CELLS,
  HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
  HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
  HARD_MAX_XLSX_WORKSHEET_ROWS,
} from './resource-policy.generated.js';
export { disposeRejectedLoad } from './rejected-load.js';
export {
  BoundedPullSession,
  DEFAULT_PULL_CANCEL_GRACE_MS,
  PULL_SESSION_PROTOCOL,
  PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
  PullSessionHost,
  PullSessionHostCoordinator,
  type PullCancelReason,
  type PullChunk,
  type PullRequestOptions,
  type PullSessionClientOptions,
  type PullSessionCommand,
  type PullSessionHostChunk,
  type PullSessionHostDriver,
  type PullSessionHostOptions,
  type PullSessionIdentity,
  type PullSessionKey,
  type PullSessionPost,
  type PullSessionResponse,
} from './pull-session.js';
