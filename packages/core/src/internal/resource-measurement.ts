/** Exact measurements without constructing monolithic JSON text. */
export interface StructuralJsonMeasurement {
  jsonBytes: number;
  stringValueUtf8Bytes: number;
}

function assertMeasurementInteger(value: number, name: string): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new Error(`${name} must be a non-negative safe integer`);
  }
}

/**
 * Adds resource counts, saturating at the canonical `limit + 1` observation.
 * At `Number.MAX_SAFE_INTEGER`, where no safe +1 exists, it saturates at the
 * limit itself.
 */
export function cappedAdd(left: number, right: number, limit: number): number {
  assertMeasurementInteger(left, 'resource measurement');
  assertMeasurementInteger(right, 'resource measurement');
  assertMeasurementInteger(limit, 'resource measurement limit');
  if (left > limit || right > limit || right > limit - left) {
    return limit === Number.MAX_SAFE_INTEGER ? limit : limit + 1;
  }
  return left + right;
}

/** Matches `TextEncoder`'s UTF-8 handling, including lone UTF-16 surrogates. */
export function utf8Bytes(value: string, limit = Number.MAX_SAFE_INTEGER): number {
  assertMeasurementInteger(limit, 'resource measurement limit');
  let bytes = 0;
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index);
    let width: number;
    if (unit <= 0x7f) width = 1;
    else if (unit <= 0x7ff) width = 2;
    else if (unit >= 0xd800 && unit <= 0xdbff && index + 1 < value.length) {
      const low = value.charCodeAt(index + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        width = 4;
        index += 1;
      } else width = 3;
    } else width = 3;
    bytes = cappedAdd(bytes, width, limit);
    if (bytes > limit) return bytes;
  }
  return bytes;
}

/** Matches the UTF-8 byte length of `JSON.stringify(value)` for a string. */
export function jsonStringBytes(value: string, limit = Number.MAX_SAFE_INTEGER): number {
  assertMeasurementInteger(limit, 'resource measurement limit');
  let bytes = cappedAdd(0, 2, limit);
  if (bytes > limit) return bytes;
  for (let index = 0; index < value.length; index += 1) {
    const unit = value.charCodeAt(index);
    let width: number;
    if (
      unit === 0x22 || unit === 0x5c || unit === 0x08 || unit === 0x09 ||
      unit === 0x0a || unit === 0x0c || unit === 0x0d
    ) width = 2;
    else if (unit <= 0x1f) width = 6;
    else if (unit <= 0x7f) width = 1;
    else if (unit <= 0x7ff) width = 2;
    else if (unit >= 0xd800 && unit <= 0xdbff && index + 1 < value.length) {
      const low = value.charCodeAt(index + 1);
      if (low >= 0xdc00 && low <= 0xdfff) {
        width = 4;
        index += 1;
      } else width = 6;
    } else if (unit >= 0xd800 && unit <= 0xdfff) width = 6;
    else width = 3;
    bytes = cappedAdd(bytes, width, limit);
    if (bytes > limit) return bytes;
  }
  return bytes;
}

function cappedLiteralBytes(bytes: number, limit: number): number {
  return cappedAdd(0, bytes, limit);
}

/**
 * Measures a JSON-compatible value without creating its monolithic JSON text.
 * It follows `JSON.stringify` for structural primitives, arrays, and ordinary
 * enumerable object properties. Object keys do not contribute to retained
 * string-value UTF-8 bytes.
 */
export function measureStructuralJson(
  value: unknown,
  limit = Number.MAX_SAFE_INTEGER,
  inArray = false,
): StructuralJsonMeasurement {
  assertMeasurementInteger(limit, 'resource measurement limit');
  if (value === null) {
    return { jsonBytes: cappedLiteralBytes(4, limit), stringValueUtf8Bytes: 0 };
  }
  if (typeof value === 'string') {
    return {
      jsonBytes: jsonStringBytes(value, limit),
      stringValueUtf8Bytes: utf8Bytes(value, limit),
    };
  }
  if (typeof value === 'boolean') {
    return { jsonBytes: cappedLiteralBytes(value ? 4 : 5, limit), stringValueUtf8Bytes: 0 };
  }
  if (typeof value === 'number') {
    const serialized = Number.isFinite(value) ? String(Object.is(value, -0) ? 0 : value) : 'null';
    return { jsonBytes: cappedLiteralBytes(serialized.length, limit), stringValueUtf8Bytes: 0 };
  }
  if (typeof value === 'bigint') {
    throw new TypeError('BigInt values cannot be serialized to JSON');
  }
  if (Array.isArray(value)) {
    let jsonBytes = cappedLiteralBytes(2, limit);
    let stringValueUtf8Bytes = 0;
    for (let index = 0; index < value.length; index += 1) {
      if (index !== 0) jsonBytes = cappedAdd(jsonBytes, 1, limit);
      const item = measureStructuralJson(value[index], limit, true);
      jsonBytes = cappedAdd(jsonBytes, item.jsonBytes, limit);
      stringValueUtf8Bytes = cappedAdd(
        stringValueUtf8Bytes,
        item.stringValueUtf8Bytes,
        limit,
      );
    }
    return { jsonBytes, stringValueUtf8Bytes };
  }
  if (typeof value === 'object') {
    let jsonBytes = cappedLiteralBytes(2, limit);
    let stringValueUtf8Bytes = 0;
    let emitted = 0;
    for (const [key, entry] of Object.entries(value)) {
      if (entry === undefined || typeof entry === 'function' || typeof entry === 'symbol') continue;
      if (emitted++ !== 0) jsonBytes = cappedAdd(jsonBytes, 1, limit);
      jsonBytes = cappedAdd(jsonBytes, jsonStringBytes(key, limit), limit);
      jsonBytes = cappedAdd(jsonBytes, 1, limit);
      const child = measureStructuralJson(entry, limit);
      jsonBytes = cappedAdd(jsonBytes, child.jsonBytes, limit);
      stringValueUtf8Bytes = cappedAdd(
        stringValueUtf8Bytes,
        child.stringValueUtf8Bytes,
        limit,
      );
    }
    return { jsonBytes, stringValueUtf8Bytes };
  }
  return {
    jsonBytes: inArray ? cappedLiteralBytes(4, limit) : 0,
    stringValueUtf8Bytes: 0,
  };
}
