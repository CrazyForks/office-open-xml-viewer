import type { ParserResourceLimits } from '../types/load-options';

function optionalPositiveByteLimit(
  name: keyof ParserResourceLimits,
  value: number | undefined,
): bigint | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(`${name} must be a positive safe integer number of bytes`);
  }
  return BigInt(value);
}

/** Convert the shared JS options to the scalar optional-u64 WASM ABI. */
export function parserResourceLimitsForWasm(
  limits: ParserResourceLimits | undefined,
): readonly [bigint | undefined, bigint | undefined] {
  return [
    optionalPositiveByteLimit(
      'maxParsedPartInflatedBytes',
      limits?.maxParsedPartInflatedBytes,
    ),
    optionalPositiveByteLimit(
      'maxOperationInflatedBytes',
      limits?.maxOperationInflatedBytes,
    ),
  ];
}
