import { describe, expect, it } from 'vitest';
import {
  PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
  PullSessionInsufficientCreditError,
  parsePullSessionInsufficientCreditError,
  requiredPullCredit,
} from './pull-credit-error.js';
import { deserializeWorkerError, serializeWorkerError } from './error-wire.js';

function envelope(requiredBytes: unknown, offeredBytes: unknown): Error {
  return new Error(`OOXML_INSUFFICIENT_CREDIT:${JSON.stringify({
    code: PULL_SESSION_INSUFFICIENT_CREDIT_CODE,
    requiredBytes,
    offeredBytes,
  })}`);
}

describe('pull insufficient-credit wire contract', () => {
  it('decodes the exact Rust envelope and preserves details across worker wire', () => {
    const parsed = parsePullSessionInsufficientCreditError(envelope(2048, 1024));
    expect(parsed).toMatchObject({ requiredBytes: 2048, offeredBytes: 1024 });
    const roundTrip = deserializeWorkerError(serializeWorkerError(parsed));
    expect(roundTrip).toBeInstanceOf(PullSessionInsufficientCreditError);
    expect(roundTrip).toMatchObject({ requiredBytes: 2048, offeredBytes: 1024 });
  });

  it.each([
    new Error('wrapped OOXML_INSUFFICIENT_CREDIT:{}'),
    new Error('OOXML_INSUFFICIENT_CREDIT:{'),
    envelope(1024, 1024),
    envelope(1024.5, 1),
    envelope(Number.MAX_SAFE_INTEGER + 1, 1),
    envelope(2048, 0),
  ])('rejects malformed or unsafe envelopes', (error) => {
    expect(parsePullSessionInsufficientCreditError(error)).toBeUndefined();
  });

  it('returns retry credit only for the exact offer and hard ceiling', () => {
    const error = new PullSessionInsufficientCreditError({
      requiredBytes: 2048,
      offeredBytes: 1024,
    });
    expect(requiredPullCredit(error, 1024, 4096)).toBe(2048);
    expect(requiredPullCredit(error, 512, 4096)).toBeUndefined();
    expect(requiredPullCredit(error, 1024, 1500)).toBeUndefined();
  });
});
