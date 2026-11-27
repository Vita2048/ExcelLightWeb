import { describe, expect, it } from 'vitest';
import { scanFormulaReferences } from '../../src/core/formula-references';
import { newWorkbook } from '../../src/core/workbook';
import { formulaHighlightPalette, formulaReferenceHighlights } from '../../src/features/formulas/formula-highlights';

describe('formula range highlights', () => {
  it('scans cell and range references across complete, incomplete, and localized formulas while ignoring literals', () => {
    const source = '=SUM(G4:G11; H7:H15; "A1:B2"; $C$3';
    const refs = scanFormulaReferences(source);
    expect(refs.map(ref => source.slice(ref.start, ref.end))).toEqual(['G4:G11', 'H7:H15', '$C$3']);
    expect(scanFormulaReferences('plain text')).toEqual([]);
  });

  it('assigns distinct palette colors per range and reuses color for repeated references', () => {
    let id = 0;
    const base = newWorkbook(() => `s-${++id}`);
    const sheet1 = base.sheets[0]!;
    const doc = {
      ...base,
      sheets: [sheet1, { ...sheet1, id: 'sheet-2', name: 'Rates', cells: [] }],
    };

    const source = '=SUM(A1:A3; B2:B4; $A$1:$A$3; Rates!C1:D2; Missing!E1)';
    const highlights = formulaReferenceHighlights(source, doc, sheet1.id);
    expect(highlights).toHaveLength(4);
    expect(highlights[0]).toEqual({
      start: 5, end: 10, sheetId: sheet1.id,
      area: { top: 1, left: 1, bottom: 3, right: 1 },
      color: formulaHighlightPalette[0],
    });
    expect(highlights[1]).toEqual({
      start: 12, end: 17, sheetId: sheet1.id,
      area: { top: 2, left: 2, bottom: 4, right: 2 },
      color: formulaHighlightPalette[1],
    });
    expect(highlights[2]).toEqual({
      start: 19, end: 28, sheetId: sheet1.id,
      area: { top: 1, left: 1, bottom: 3, right: 1 },
      color: formulaHighlightPalette[0],
    });
    expect(highlights[3]).toEqual({
      start: 30, end: 41, sheetId: 'sheet-2',
      area: { top: 1, left: 3, bottom: 2, right: 4 },
      color: formulaHighlightPalette[2],
    });
  });
});
