import { referenceBox, scanFormulaReferences } from '../../core/formula-references';
import { documentLimits, type WorkbookDocument } from '../../core/workbook';
import { ordinalNameKey } from '../../core/workbook-validation';
import type { Area } from '../../grid/grid-selection';

export const formulaHighlightPalette = [
  '#5ca8ff',
  '#f46b6b',
  '#b47cff',
  '#4cd47e',
  '#f5a642',
  '#45d0d8',
  '#f071b8',
  '#d4d84c',
] as const;

export interface FormulaReferenceHighlight {
  readonly start: number;
  readonly end: number;
  readonly sheetId: string;
  readonly area: Area;
  readonly color: string;
}

export function formulaReferenceHighlights(
  source: string,
  document: WorkbookDocument,
  targetSheetId: string,
): readonly FormulaReferenceHighlight[] {
  const refs = scanFormulaReferences(source);
  if (!refs.length) return [];

  const sheetsByName = new Map(document.sheets.map(sheet => [ordinalNameKey(sheet.name), sheet.id]));
  const colorByKey = new Map<string, string>();
  const highlights: FormulaReferenceHighlight[] = [];

  for (const ref of refs) {
    if (ref.endSheet !== undefined) continue;
    const sheetId = ref.sheet === undefined ? targetSheetId : sheetsByName.get(ordinalNameKey(ref.sheet));
    if (!sheetId) continue;
    const area = referenceBox(ref);
    if (area.bottom > documentLimits.rows || area.right > documentLimits.columns) continue;
    const key = `${sheetId}:${area.top}:${area.left}:${area.bottom}:${area.right}`;
    let color = colorByKey.get(key);
    if (!color) {
      color = formulaHighlightPalette[colorByKey.size % formulaHighlightPalette.length]!;
      colorByKey.set(key, color);
    }
    highlights.push({ start: ref.start, end: ref.end, sheetId, area, color });
  }
  return highlights;
}
