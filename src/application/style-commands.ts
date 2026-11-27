import { cellKey, documentLimits, findCell, findColumnStyle, findRowStyle, ownDto, resolvedCellStyle,
  type AxisDimension, type CellDocument, type CellStyle, type SheetDocument,
} from '../core/workbook';
import { equalDto } from '../core/workbook-delta';
import { requireDocument } from '../core/workbook-validation';
import type { AlignmentPatch } from './alignment-commands';
import type { BorderDraft } from '../core/formats/borders';
import { applyAxisBorderChanges, applyBorderDraft, draftBorderOps, mirrorBorderDraft } from './border-commands';
import { isNumberFormat } from '../core/formats/format-code';
import { updateAppearance, type AppearancePatch } from '../core/formats/appearance';
import { checkedSheet, expandMerges, rangeSpanKind, type EditRange, type SelectionSpanKind } from './editing-commands';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { CellEdit, SetCellsCommand, StructureCommand, WorkbookCommand } from './workbook-session';

export { rangeSpanKind, type SelectionSpanKind };

function normalizeStyle(style: CellStyle | undefined): CellStyle | undefined {
  return style && Object.keys(style).length ? style : undefined;
}

function reconcileCellStyle(style: CellStyle | undefined, inherited: CellStyle | undefined): CellStyle | undefined {
  const normalized = normalizeStyle(style);
  if (normalized) return equalDto(normalized, inherited) ? undefined : normalized;
  return inherited ? {} : undefined;
}

export function changeStyles(basis: WorkbookSnapshot, ranges: readonly EditRange[], label: string,
  update: (style: CellStyle | undefined, row: number, column: number, area: EditRange) => CellStyle | undefined,
  includeMerges = true): WorkbookCommand {
  const hasAxisRange = ranges.some(range => {
    checkedSheet(basis, range);
    return rangeSpanKind(range) !== 'cells';
  });
  if (!hasAxisRange) {
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
          const current = resolvedCellStyle(sheet, target, cell);
          const changed = update(current, row, column, area);
          const inherited = findRowStyle(sheet, row) ?? findColumnStyle(sheet, column);
          const style = reconcileCellStyle(changed, inherited);
          if (equalDto(style, cell?.style)) continue;
          edits.push({ target, content: cell?.content ?? { kind: 'blank' }, preserveStyle: false,
            ...(style !== undefined && { style }) });
        }
      }
    }
    return { kind: 'set-cells', label, edits };
  }

  const bySheet = new Map<string, EditRange[]>();
  for (const range of ranges) {
    checkedSheet(basis, range);
    const list = bySheet.get(range.sheetId) ?? [];
    list.push(range);
    bySheet.set(range.sheetId, list);
  }

  const allEdits: CellEdit[] = [];
  const nextSheets = new Map<string, SheetDocument>();

  for (const [sheetId, sheetRanges] of bySheet) {
    const sheet = basis.document.sheets.find(s => s.id === sheetId)!;
    const rowDims = new Map<number, AxisDimension>(sheet.rows.map(dim => [dim.index, dim]));
    const colDims = new Map<number, AxisDimension>(sheet.columns.map(dim => [dim.index, dim]));
    const pending = new Map<number, { row: number; column: number; prior: CellDocument | undefined; style: CellStyle | undefined }>();

    function getCellStyle(row: number, column: number): CellStyle | undefined {
      const key = cellKey({ row, column });
      if (pending.has(key)) return pending.get(key)!.style;
      const prior = findCell(sheet, { row, column });
      if (prior?.style !== undefined) return prior.style;
      return rowDims.get(row)?.style ?? colDims.get(column)?.style;
    }

    function setCellStyle(row: number, column: number, style: CellStyle | undefined) {
      const key = cellKey({ row, column });
      const prior = pending.has(key) ? pending.get(key)!.prior : findCell(sheet, { row, column });
      pending.set(key, { row, column, prior, style: normalizeStyle(style) });
      requireDocument(pending.size <= documentLimits.cells, 'Format selection is too large.');
    }

    function setDim(map: Map<number, AxisDimension>, index: number, nextStyle: CellStyle | undefined) {
      const normalized = normalizeStyle(nextStyle);
      const prior = map.get(index);
      const next: AxisDimension = {
        index,
        ...(prior?.size !== undefined && { size: prior.size }),
        ...(prior?.hidden && { hidden: true }),
        ...(normalized !== undefined && { style: normalized }),
      };
      if (next.size === undefined && !next.hidden && !next.style) map.delete(index);
      else map.set(index, next);
    }

    for (const range of sheetRanges) {
      const area = includeMerges ? expandMerges(range, sheet) : range;
      const span = rangeSpanKind(area);
      if (span === 'cells') {
        requireDocument(area.rows * area.columns <= documentLimits.cells, 'Format selection is too large.');
        for (let r = area.row; r < area.row + area.rows; r++) {
          for (let c = area.column; c < area.column + area.columns; c++) {
            setCellStyle(r, c, update(getCellStyle(r, c), r, c, area));
          }
        }
      } else if (span === 'columns') {
        const c1 = area.column, c2 = area.column + area.columns - 1;
        const coords = new Map<number, { row: number; column: number; style: CellStyle | undefined }>();
        const track = (row: number, column: number) => {
          const key = cellKey({ row, column });
          if (!coords.has(key)) coords.set(key, { row, column, style: getCellStyle(row, column) });
        };
        for (const cell of sheet.cells) {
          if (cell.column >= c1 && cell.column <= c2) track(cell.row, cell.column);
        }
        for (const entry of pending.values()) {
          if (entry.column >= c1 && entry.column <= c2) track(entry.row, entry.column);
        }
        for (const dim of rowDims.values()) {
          if (dim.style === undefined) continue;
          for (let c = c1; c <= c2; c++) track(dim.index, c);
        }
        for (const { row, column, style } of coords.values()) {
          setCellStyle(row, column, update(style, row, column, area));
        }
        for (let c = c1; c <= c2; c++) {
          setDim(colDims, c, update(colDims.get(c)?.style, 1, c, area));
        }
      } else if (span === 'rows') {
        const r1 = area.row, r2 = area.row + area.rows - 1;
        const coords = new Map<number, { row: number; column: number; style: CellStyle | undefined }>();
        const track = (row: number, column: number) => {
          const key = cellKey({ row, column });
          if (!coords.has(key)) coords.set(key, { row, column, style: getCellStyle(row, column) });
        };
        for (const cell of sheet.cells) {
          if (cell.row >= r1 && cell.row <= r2) track(cell.row, cell.column);
        }
        for (const entry of pending.values()) {
          if (entry.row >= r1 && entry.row <= r2) track(entry.row, entry.column);
        }
        for (const dim of colDims.values()) {
          if (dim.style === undefined) continue;
          for (let r = r1; r <= r2; r++) track(r, dim.index);
        }
        for (const { row, column, style } of coords.values()) {
          setCellStyle(row, column, update(style, row, column, area));
        }
        for (let r = r1; r <= r2; r++) {
          setDim(rowDims, r, update(rowDims.get(r)?.style, r, 1, area));
        }
      } else {
        const coords = new Map<number, { row: number; column: number; style: CellStyle | undefined }>();
        const track = (row: number, column: number) => {
          const key = cellKey({ row, column });
          if (!coords.has(key)) coords.set(key, { row, column, style: getCellStyle(row, column) });
        };
        for (const cell of sheet.cells) track(cell.row, cell.column);
        for (const entry of pending.values()) track(entry.row, entry.column);
        for (const { row, column, style } of coords.values()) {
          setCellStyle(row, column, update(style, row, column, area));
        }
        for (const [r, dim] of [...rowDims.entries()]) {
          if (dim.style !== undefined) setDim(rowDims, r, update(dim.style, r, 1, area));
        }
        for (let c = 1; c <= documentLimits.columns; c++) {
          setDim(colDims, c, update(colDims.get(c)?.style, 1, c, area));
        }
      }
    }

    for (const { row, column, prior, style: rawStyle } of pending.values()) {
      const inherited = rowDims.get(row)?.style ?? colDims.get(column)?.style;
      const style = reconcileCellStyle(rawStyle, inherited);
      if (equalDto(style, prior?.style)) continue;
      allEdits.push({
        target: { sheetId, row, column },
        content: prior?.content ?? { kind: 'blank' },
        preserveStyle: false,
        ...(style !== undefined && { style }),
      });
    }

    nextSheets.set(sheetId, {
      ...sheet,
      rows: [...rowDims.values()].sort((a, b) => a.index - b.index),
      columns: [...colDims.values()].sort((a, b) => a.index - b.index),
    });
  }

  return {
    kind: 'structure',
    label,
    document: { ...basis.document, sheets: basis.document.sheets.map(s => nextSheets.get(s.id) ?? s) },
    edits: allEdits,
  } satisfies StructureCommand;
}

export function formatStyles(basis: WorkbookSnapshot, ranges: readonly EditRange[], patch: AppearancePatch,
  numberCode?: string, borders: BorderDraft = {}, alignment: AlignmentPatch = {}): WorkbookCommand {
  requireDocument(numberCode === undefined || isNumberFormat(numberCode), 'Enter a valid Excel number format.');
  const withBorder = Object.keys(borders).length > 0;
  const hasAxisRange = ranges.some(range => rangeSpanKind(range) !== 'cells');
  const sheets = new Map(basis.document.sheets.map(sheet => [sheet.id, sheet]));
  const applyNonBorders = (style: CellStyle | undefined): CellStyle => {
    const appearance = updateAppearance(style, patch);
    const aligned = Object.keys(alignment).length ? Object.fromEntries(Object.entries({ ...appearance, ...alignment })
      .filter(([, value]) => value !== undefined)) as CellStyle : appearance;
    if (numberCode === undefined || numberCode === style?.numberFormat) return aligned;
    const { numberFormat: _prior, ...rest } = aligned;
    return /^general$/iu.test(numberCode) ? rest : { ...rest, numberFormat: numberCode };
  };

  if (!withBorder) {
    return changeStyles(basis, ranges, 'Format Cells', applyNonBorders);
  }

  if (!hasAxisRange) {
    const command = changeStyles(basis, ranges, 'Format Cells', (style, row, column, area) => {
      const aligned = applyNonBorders(style);
      return applyBorderDraft(sheets.get(area.sheetId)!, area, aligned, row, column, borders);
    }) as SetCellsCommand;
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

  const baseCmd = changeStyles(basis, ranges, 'Format Cells', applyNonBorders) as StructureCommand;
  const bySheetRanges = new Map<string, EditRange[]>();
  for (const range of ranges) {
    const sheet = baseCmd.document.sheets.find(s => s.id === range.sheetId)!;
    const areas = bySheetRanges.get(range.sheetId) ?? [];
    areas.push(expandMerges(range, sheet));
    bySheetRanges.set(range.sheetId, areas);
  }
  const editsBySheet = new Map<string, Map<number, CellEdit>>();
  for (const edit of baseCmd.edits ?? []) {
    let map = editsBySheet.get(edit.target.sheetId);
    if (!map) editsBySheet.set(edit.target.sheetId, map = new Map());
    map.set(cellKey(edit.target), edit);
  }

  const finalEdits = new Map<string, CellEdit>();
  for (const edit of baseCmd.edits ?? []) {
    finalEdits.set(`${edit.target.sheetId}:${cellKey(edit.target)}`, edit);
  }
  const finalSheets = new Map<string, SheetDocument>();

  for (const [sheetId, areas] of bySheetRanges) {
    const baseSheet = baseCmd.document.sheets.find(s => s.id === sheetId)!;
    const baseEdits = editsBySheet.get(sheetId);
    let sheetWithBaseCells = baseSheet;
    if (baseEdits?.size) {
      const cellMap = new Map(baseSheet.cells.map(c => [cellKey(c), c]));
      for (const [key, edit] of baseEdits) {
        if (edit.content.kind === 'blank' && edit.style === undefined) cellMap.delete(key);
        else {
          const prior = cellMap.get(key);
          cellMap.set(key, {
            id: prior?.id ?? `pending-${key}`,
            row: edit.target.row,
            column: edit.target.column,
            content: edit.content,
            ...(edit.style !== undefined && { style: edit.style }),
          });
        }
      }
      sheetWithBaseCells = { ...baseSheet, cells: [...cellMap.values()] };
    }
    const { rowOps, colOps, cellEdges } = draftBorderOps(sheetWithBaseCells, areas, borders);
    const borderCmd = applyAxisBorderChanges(
      { ...basis, document: { ...baseCmd.document, sheets: baseCmd.document.sheets.map(s => s.id === sheetId ? sheetWithBaseCells : s) } },
      sheetWithBaseCells,
      rowOps,
      colOps,
      cellEdges,
      'Format Cells',
    );
    finalSheets.set(sheetId, borderCmd.document.sheets.find(s => s.id === sheetId)!);
    const origSheet = sheets.get(sheetId)!;
    for (const edit of borderCmd.edits ?? []) {
      const mapKey = `${sheetId}:${cellKey(edit.target)}`;
      const orig = findCell(origSheet, edit.target);
      const nextStyle = edit.style;
      if (equalDto(orig?.style, nextStyle) && equalDto(orig?.content ?? { kind: 'blank' }, edit.content)) {
        finalEdits.delete(mapKey);
      } else {
        finalEdits.set(mapKey, edit);
      }
    }
  }

  return {
    kind: 'structure',
    label: 'Format Cells',
    document: { ...baseCmd.document, sheets: baseCmd.document.sheets.map(s => finalSheets.get(s.id) ?? s) },
    edits: [...finalEdits.values()],
  };
}

export interface FormatBrush {
  readonly sessionId: string;
  readonly rows: number;
  readonly columns: number;
  readonly kind?: SelectionSpanKind;
  readonly styles: readonly (CellStyle | null)[];
}

export function captureFormatBrush(basis: WorkbookSnapshot, source: EditRange): FormatBrush {
  const sheet = checkedSheet(basis, source);
  const span = rangeSpanKind(source);
  if (span === 'columns') {
    const styles = Array.from({ length: source.columns }, (_, i) => findColumnStyle(sheet, source.column + i) ?? null);
    return ownDto({ sessionId: basis.sessionId, rows: 1, columns: source.columns, kind: 'columns', styles });
  }
  if (span === 'rows') {
    const styles = Array.from({ length: source.rows }, (_, i) => findRowStyle(sheet, source.row + i) ?? null);
    return ownDto({ sessionId: basis.sessionId, rows: source.rows, columns: 1, kind: 'rows', styles });
  }
  if (span === 'sheet') {
    const styles = [findColumnStyle(sheet, 1) ?? null];
    return ownDto({ sessionId: basis.sessionId, rows: 1, columns: 1, kind: 'sheet', styles });
  }
  const range = expandMerges(source, sheet);
  requireDocument(range.rows * range.columns <= documentLimits.cells, 'Format selection is too large.');
  const styles: (CellStyle | null)[] = [];
  for (let row = range.row; row < range.row + range.rows; row++) {
    for (let column = range.column; column < range.column + range.columns; column++) {
      styles.push(resolvedCellStyle(sheet, { row, column }) ?? null);
    }
  }
  return ownDto({ sessionId: basis.sessionId, rows: range.rows, columns: range.columns, styles });
}

export function activeFormatBrush<T extends FormatBrush>(basis: Pick<WorkbookSnapshot, 'sessionId'>, brush?: T): T | undefined {
  return brush?.sessionId === basis.sessionId ? brush : undefined;
}

export function paintStyles(basis: WorkbookSnapshot, target: EditRange, brush: FormatBrush): WorkbookCommand {
  requireDocument(activeFormatBrush(basis, brush), 'Format Painter belongs to a different workbook.');
  const targetSpan = rangeSpanKind(target);
  if (targetSpan === 'columns') {
    requireDocument(brush.rows === 1, 'Format selection is too large.');
    return changeStyles(basis, [target], 'Format Painter', (_style, _row, column) =>
      brush.styles[(column - target.column) % brush.columns] ?? undefined, false);
  }
  if (targetSpan === 'rows') {
    requireDocument(brush.columns === 1, 'Format selection is too large.');
    return changeStyles(basis, [target], 'Format Painter', (_style, row) =>
      brush.styles[((row - target.row) % brush.rows) * brush.columns] ?? undefined, false);
  }
  if (targetSpan === 'sheet') {
    requireDocument(brush.rows === 1 && brush.columns === 1, 'Format selection is too large.');
    return changeStyles(basis, [target], 'Format Painter', () => brush.styles[0] ?? undefined, false);
  }
  const range = target.rows === 1 && target.columns === 1 ? { ...target, rows: brush.rows, columns: brush.columns } : target;
  return changeStyles(basis, [range], 'Format Painter', (_style, row, column) =>
    brush.styles[((row - range.row) % brush.rows) * brush.columns + (column - range.column) % brush.columns] ?? undefined, false);
}
