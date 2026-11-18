import { describe, expect, it } from 'vitest';
import { formatNumber } from '../../src/core/formats/cell-format';
import { temporalEntry } from '../../src/core/formats/temporal';
import desktop from '../fixtures/formats/S10-desktop.json';

describe('desktop numeric display and temporal entry', () => {
  it.each(desktop.cases)('$code with $value', ({ code, value, display, color, entry }) => {
    expect.soft(formatNumber(value, code)).toEqual({ text: display, ...(color && { color }) });
    expect.soft(temporalEntry(value, code)).toBe(entry ?? undefined);
  });
});
