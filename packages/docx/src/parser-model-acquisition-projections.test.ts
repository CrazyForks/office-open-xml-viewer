import { describe, expect, it } from 'vitest';
import {
  bodyAcquisitionInputProjections,
  numberingMarkerShapeInput,
  paragraphAcquisitionInput,
  paragraphMarkShapeInput,
  tableColumnLayoutInput,
  tableFormatInput,
  tableParticipatesInOrdinaryFlow,
} from './parser-model.js';

describe('parser-to-body-acquisition projection capability', () => {
  it('is one frozen identity-preserving record without compatibility wrappers', () => {
    expect(Object.isFrozen(bodyAcquisitionInputProjections)).toBe(true);
    expect(Object.keys(bodyAcquisitionInputProjections).sort()).toEqual([
      'numberingMarkerShapeInput',
      'paragraphAcquisitionInput',
      'paragraphMarkShapeInput',
      'tableColumnLayoutInput',
      'tableFormatInput',
      'tableParticipatesInOrdinaryFlow',
    ]);
    expect(bodyAcquisitionInputProjections.numberingMarkerShapeInput)
      .toBe(numberingMarkerShapeInput);
    expect(bodyAcquisitionInputProjections.paragraphAcquisitionInput)
      .toBe(paragraphAcquisitionInput);
    expect(bodyAcquisitionInputProjections.paragraphMarkShapeInput)
      .toBe(paragraphMarkShapeInput);
    expect(bodyAcquisitionInputProjections.tableColumnLayoutInput)
      .toBe(tableColumnLayoutInput);
    expect(bodyAcquisitionInputProjections.tableFormatInput).toBe(tableFormatInput);
    expect(bodyAcquisitionInputProjections.tableParticipatesInOrdinaryFlow)
      .toBe(tableParticipatesInOrdinaryFlow);
  });
});
