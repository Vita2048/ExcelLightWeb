import { cellKey, documentLimits, findCell, ownDto, type CellStyle } from '../core/workbook';
import { equalDto } from '../core/workbook-delta';
import { requireDocument } from '../core/workbook-validation';
import type { AlignmentPatch } from './alignment-commands';
import type { BorderDraft } from '../core/formats/borders';
import { applyBorderDraft, mirrorBorderDraft } from './border-commands';
import { isNumberFormat } from '../core/formats/format-code';
import { updateAppearance, type AppearancePatch } from '../core/formats/appearance';
import { checkedSheet, expandMerges, type EditRange } from './editing-commands';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { CellEdit, SetCellsCommand } from './workbook-session';

export function changeStyles(basis: WorkbookSnapshot, ranges: readonly EditRange[], label: string,
  update: (style: CellStyle | undefined, row: number, column: number, area: EditRange) => CellStyle | undefined,
  includeMerges = true): SetCellsCommand {
  const edits: CellEdit[] = [];
  const visited = new Map<string, Set<number>>();
  let count = 0;
  for (const range of ranges) {
    const sheet = checkedSheet(basis, range);
    const area = includeMerges ? expandMerges(range, sheet) : range;
    requireDocument(area.rows * area.columns <= documentLimits.cells, 'Format selection is too large.');
    const seen = visited.get(sheet.id) ?? new Set<number>();
    visited.set(sheet.id, seen);
    for (let row = area.row; row < area.row + area.rows; row++) {
      for (let column = area.column; column < area.column + area.columns; column++) {
        const target = { sheetId: sheet.id, row, column };
        const key = cellKey(target);
        if (seen.has(key)) continue;
        seen.add(key);
        requireDocument(++count <= documentLimits.cells, 'Format selection is too large.');
        const cell = findCell(sheet, target);
        const changed = update(cell?.style, row, column, area);
        if (changed === cell?.style) continue;
        const style = changed && Object.keys(changed).length ? changed : undefined;
        if (equalDto(style, cell?.style)) continue;
        edits.push({ target, content: cell?.content ?? { kind: 'blank' }, preserveStyle: false, ...(style && { style }) });
      }
    }
  }
  return { kind: 'set-cells', label, edits };
}

export function formatStyles(basis: WorkbookSnapshot, ranges: readonly EditRange[], patch: AppearancePatch,
  numberCode?: string, borders: BorderDraft = {}, alignment: AlignmentPatch = {}): SetCellsCommand {
  requireDocument(numberCode === undefined || isNumberFormat(numberCode), 'Enter a valid Excel number format.');
  const withBorder = Object.keys(borders).length > 0;
  const sheets = new Map(basis.document.sheets.map(sheet => [sheet.id, sheet]));
  const command = changeStyles(basis, ranges, 'Format Cells', (style, row, column, area) => {
    const appearance = updateAppearance(style, patch);
    const aligned = Object.keys(alignment).length ? Object.fromEntries(Object.entries({ ...appearance, ...alignment })
      .filter(([, value]) => value !== undefined)) as CellStyle : appearance;
    const next = (withBorder ? applyBorderDraft(sheets.get(area.sheetId)!, area, aligned, row, column, borders) : aligned) ?? {};
    if (numberCode === undefined || numberCode === style?.numberFormat) return next;
    const { numberFormat: _prior, ...rest } = next;
    return /^general$/iu.test(numberCode) ? rest : { ...rest, numberFormat: numberCode };
  });
  if (!withBorder) return command;

  const bySheet = new Map<string, EditRange[]>();
  for (const range of ranges) {
    const areas = bySheet.get(range.sheetId) ?? [];
    areas.push(expandMerges(range, sheets.get(range.sheetId)!));
    bySheet.set(range.sheetId, areas);
  }
  const mirrors = Array.from(bySheet, ([id, areas]) => mirrorBorderDraft(sheets.get(id)!, areas, borders)).flat();
  requireDocument(command.edits.length + mirrors.length <= documentLimits.cells, 'Format selection is too large.');
  return { ...command, edits: [...command.edits, ...mirrors] };
}

export interface FormatBrush {
  readonly sessionId: string;
  readonly rows: number;
  readonly columns: number;
  readonly styles: readonly (CellStyle | null)[];
}

export function captureFormatBrush(basis: WorkbookSnapshot, source: EditRange): FormatBrush {
  const sheet = checkedSheet(basis, source);
  const range = expandMerges(source, sheet);
  requireDocument(range.rows * range.columns <= documentLimits.cells, 'Format selection is too large.');
  const styles: (CellStyle | null)[] = [];
  for (let row = range.row; row < range.row + range.rows; row++) {
    for (let column = range.column; column < range.column + range.columns; column++) {
      styles.push(findCell(sheet, { row, column })?.style ?? null);
    }
  }
  return ownDto({ sessionId: basis.sessionId, rows: range.rows, columns: range.columns, styles });
}

export function activeFormatBrush<T extends FormatBrush>(basis: Pick<WorkbookSnapshot, 'sessionId'>, brush?: T): T | undefined {
  return brush?.sessionId === basis.sessionId ? brush : undefined;
}

export function paintStyles(basis: WorkbookSnapshot, target: EditRange, brush: FormatBrush): SetCellsCommand {
  requireDocument(activeFormatBrush(basis, brush), 'Format Painter belongs to a different workbook.');
  const range = target.rows === 1 && target.columns === 1 ? { ...target, rows: brush.rows, columns: brush.columns } : target;
  return changeStyles(basis, [range], 'Format Painter', (_style, row, column) =>
    brush.styles[((row - range.row) % brush.rows) * brush.columns + (column - range.column) % brush.columns] ?? undefined, false);
}
