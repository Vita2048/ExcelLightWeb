import { describe, expect, it } from 'vitest';
import { nativeFunctions } from '../../src/calculation/formula-functions';
import { previewFormula } from '../../src/calculation/native-calculator';
import { newWorkbook } from '../../src/core/workbook';
import {
  allLibraryCategories,
  buildFunctionFormula,
  defaultRecentFunctions,
  dialogCategoryOptions,
  formatArgumentReference,
  functionCategory,
  functionDescription,
  functionSignatureText,
  functionsByCategory,
  initialFunctionArguments,
  isVariadicFunction,
  maxRecentFunctions,
  moreLibraryCategories,
  nextVariadicArgument,
  recordRecentFunction,
  sanitizeRecentFunctions,
  searchFunctions,
  topLibraryCategories,
} from '../../src/features/formulas/function-library';

describe('function library catalog', () => {
  it('maps every native function once across Excel categories with no empty categories', () => {
    const categorized = allLibraryCategories.flatMap(category => functionsByCategory[category]);
    expect(new Set(categorized).size).toBe(nativeFunctions.length);
    expect(categorized).toHaveLength(nativeFunctions.length);
    for (const name of nativeFunctions) {
      expect(categorized).toContain(name);
      expect(allLibraryCategories).toContain(functionCategory(name));
      expect(functionDescription(name).trim().length).toBeGreaterThan(10);
      expect(functionSignatureText(name, ',')).toMatch(new RegExp(`^${name.replace('.', '\\.')}\\(`));
      expect(functionSignatureText(name, ';')).toMatch(new RegExp(`^${name.replace('.', '\\.')}\\(`));
    }
    for (const category of [...topLibraryCategories, ...moreLibraryCategories]) {
      expect(functionsByCategory[category].length).toBeGreaterThan(0);
    }
    expect(dialogCategoryOptions).toEqual(['Most Recently Used', 'All', 'Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Text', 'Logical', 'Information']);
  });

  it('searches across names and descriptions, and respects filters when query is empty', () => {
    expect(searchFunctions('', 'Most Recently Used', defaultRecentFunctions)).toEqual(defaultRecentFunctions);
    expect(searchFunctions('', 'All')).toHaveLength(nativeFunctions.length);
    expect(searchFunctions('', 'Logical')).toEqual(functionsByCategory.Logical);
    expect(searchFunctions('xlookup', 'Financial')).toContain('XLOOKUP');
    expect(searchFunctions('depreciation', 'Most Recently Used')).toEqual(expect.arrayContaining(['DB', 'DDB', 'SLN']));
  });

  it('maintains a bounded deduplicated recently used list', () => {
    const updated = recordRecentFunction(defaultRecentFunctions, 'XLOOKUP');
    expect(updated[0]).toBe('XLOOKUP');
    expect(updated.filter(name => name === 'XLOOKUP')).toHaveLength(1);
    expect(updated).toHaveLength(maxRecentFunctions);

    const added = recordRecentFunction(defaultRecentFunctions, 'PMT');
    expect(added[0]).toBe('PMT');
    expect(added).toHaveLength(maxRecentFunctions);
    expect(added).not.toContain(defaultRecentFunctions[maxRecentFunctions - 1]);

    expect(sanitizeRecentFunctions(undefined)).toEqual(defaultRecentFunctions);
    expect(sanitizeRecentFunctions(['UNKNOWN'])).toEqual(defaultRecentFunctions);
    expect(sanitizeRecentFunctions(['PMT', 'UNKNOWN', 'PV'])).toEqual(['PMT', 'PV']);
  });

  it('extracts parameter specs, expands variadic arguments, and builds localized formulas', () => {
    expect(initialFunctionArguments('IF')).toEqual([
      { name: 'logical_test', required: true },
      { name: 'value_if_true', required: false },
      { name: 'value_if_false', required: false },
    ]);
    expect(initialFunctionArguments('TODAY')).toEqual([]);
    expect(isVariadicFunction('SUM')).toBe(true);
    expect(isVariadicFunction('IF')).toBe(false);

    const sumSpecs = [...initialFunctionArguments('SUM')];
    expect(nextVariadicArgument('SUM', sumSpecs)).toEqual({ name: 'number3', required: false });

    const sumifsSpecs = [...initialFunctionArguments('SUMIFS')];
    const extraRange = nextVariadicArgument('SUMIFS', sumifsSpecs)!;
    expect(extraRange).toEqual({ name: 'criteria_range3', required: false });
    expect(nextVariadicArgument('SUMIFS', [...sumifsSpecs, extraRange])).toEqual({ name: 'criteria3', required: false });

    expect(buildFunctionFormula('IF', ['TRUE', '"a"', ''], ';')).toBe('=IF(TRUE;"a")');
    expect(buildFunctionFormula('INDEX', ['A1:C5', '', '2'], ',')).toBe('=INDEX(A1:C5,,2)');
    expect(formatArgumentReference('=Sheet1!$A$1:$B$2', 'Sheet1')).toBe('A1:B2');
    expect(formatArgumentReference('=Sheet2!$A$1:$B$2', 'Sheet1')).toBe('Sheet2!A1:B2');
    expect(formatArgumentReference("='Sales Data'!$C$3", 'Sales Data')).toBe('C3');
    expect(formatArgumentReference("='$Sales'!$B$2:$D$5", 'Sheet1')).toBe("'$Sales'!B2:D5");
  });

  it('previews scalar, range, error, and zero-arg formulas without mutating the workbook', () => {
    let counter = 0;
    const book = newWorkbook(() => `id-${++counter}`);
    const sheet = book.sheets[0]!;
    const populated = {
      ...book,
      sheets: [{
        ...sheet,
        cells: [
          { id: 'c1', row: 1, column: 1, content: { kind: 'number' as const, value: 10 } },
          { id: 'c2', row: 2, column: 1, content: { kind: 'number' as const, value: 25 } },
        ],
      }],
    };
    const target = { sheetId: sheet.id, row: 1, column: 2 };
    expect(previewFormula(populated, target, '=SUM(A1:A2)')).toEqual({ text: '35', isError: false });
    expect(previewFormula(populated, target, '=A1:A2')).toEqual({ text: '{10; 25}', isError: false });
    expect(previewFormula(populated, target, '=1/0')).toEqual({ text: '#DIV/0!', isError: true });
    expect(previewFormula(populated, target, '=IF(TRUE,"ok","no")', 'en-US', false)).toEqual({ text: 'ok', isError: false });
    expect(previewFormula(populated, target, '=TODAY()')?.isError).toBe(false);
    expect(previewFormula(populated, target, '=SUM(')).toBeUndefined();
    expect(populated.sheets[0]!.cells).toHaveLength(2);
  });
});
