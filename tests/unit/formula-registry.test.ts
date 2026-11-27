import { describe, expect, it } from 'vitest';
import {
  functionRegistry,
  getFunctionSpec,
  isVolatileFunction,
  nativeFunctions,
  supportedFunctions,
} from '../../src/calculation/formula-functions';
import { calculateNative } from '../../src/calculation/native-calculator';
import { parseFormula } from '../../src/core/formula-references';
import { emptySheet, formulaErrors, ownDto, type WorkbookDocument } from '../../src/core/workbook';
import { formulaHint } from '../../src/features/editing/formula-hints';

function makeWorkbook(sources: readonly string[]): WorkbookDocument {
  return ownDto({
    id: 'registry-test',
    formatVersion: 1,
    dateSystem: '1900',
    sheets: [{
      ...emptySheet('s1', 'Sheet1'),
      cells: sources.map((source, index) => ({
        id: `c-${index}`,
        row: index + 1,
        column: 1,
        content: { kind: 'formula', source },
      })),
    }],
    names: [],
  });
}

describe('formula function registry and error catalog', () => {
  it('registers metadata, signatures and volatility for every native function', () => {
    expect(functionRegistry.size).toBe(nativeFunctions.length);
    for (const name of nativeFunctions) {
      expect(supportedFunctions.has(name)).toBe(true);
      const spec = getFunctionSpec(name);
      expect(spec, name).toBeDefined();
      expect(spec!.minArgs).toBeGreaterThanOrEqual(0);
      expect(spec!.maxArgs).toBeGreaterThanOrEqual(spec!.minArgs);
      expect(spec!.category.length).toBeGreaterThan(0);
      const hint = formulaHint(`=${name.toLowerCase()}(`, name.length + 2);
      expect(hint?.name).toBe(name);
      expect(hint?.parameters).toEqual(spec!.signature[0]);
    }
    expect(isVolatileFunction('TODAY')).toBe(true);
    expect(isVolatileFunction('today')).toBe(true);
    expect(isVolatileFunction('SUM')).toBe(false);
    expect(isVolatileFunction('UNKNOWN')).toBe(false);
  });

  it('includes #SPILL! and #CALC! in formulaErrors, parser, and IFERROR evaluation', () => {
    expect(formulaErrors).toContain('#SPILL!');
    expect(formulaErrors).toContain('#CALC!');
    expect(parseFormula('=#SPILL!').expression).toEqual({ kind: 'error', value: '#SPILL!' });
    expect(parseFormula('=#CALC!').expression).toEqual({ kind: 'error', value: '#CALC!' });

    const { updates } = calculateNative(makeWorkbook([
      '=#SPILL!',
      '=#CALC!',
      '=IFERROR(#SPILL!,"caught-spill")',
      '=IFERROR(#CALC!,"caught-calc")',
    ]));
    expect(updates.map(update => update.cache)).toEqual([
      { kind: 'error', value: '#SPILL!' },
      { kind: 'error', value: '#CALC!' },
      { kind: 'text', value: 'caught-spill' },
      { kind: 'text', value: 'caught-calc' },
    ]);
  });
});
