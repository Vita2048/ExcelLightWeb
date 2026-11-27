import { cellKey, defaultPythonSource, documentLimits as limits, findCell, findColumnStyle, findRowStyle, ownDto, resolvedCellStyle, worksheetViewport,
  type AxisDimension, type CellContent, type CellDocument, type CellExtent, type CellReference, type CellStyle, type Coordinate, type PythonBlock, type SheetDocument,
} from '../core/workbook';
import { equalDto } from '../core/workbook-delta';
import { parseTemporalEntry, temporalEntry } from '../core/formats/temporal';
import { requireDocument } from '../core/workbook-validation';
import { copyReferences, formulaAddress, intersection, moveReferences, type ReferenceBox, type ReferenceMove } from '../core/formula-references';
import { transformPythonBlock } from '../python/references';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { CellEdit, SetCellsCommand, WorkbookCommand } from './workbook-session';

export type PasteMode = 'all' | 'values' | 'formulas';
export type ClearMode = 'all' | 'contents' | 'formats';
export type FillDirection = 'down' | 'right' | 'up' | 'left';
export type Aggregate = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN';
export interface EditRange extends CellExtent { readonly sheetId: string }
export interface ClipboardBlock {
  readonly sessionId: string;
  readonly authoredRevision: number;
  readonly range: EditRange;
  readonly cells: readonly (CellDocument | null)[];
  readonly merges: readonly CellExtent[];
  readonly cut: boolean;
}
export const clipboardTextLimit = 4 * 1024 * 1024;
const blank = Object.freeze({ kind: 'blank' } as const);
const numberPattern = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/iu;

export function isPythonEntry(text: string): boolean {
  return /^=PY(?:\(\))?$/iu.test(text.trim());
}

function pythonContent(makeId: () => string): CellContent {
  const block: PythonBlock = { id: makeId(), source: defaultPythonSource, bindings: [], lastValue: null,
    hasResult: false, inputFingerprint: null, runtimeFingerprint: null, status: 'never', error: null,
    lastTable: null, spillActive: false, lastFigure: null, placements: [] };
  return { kind: 'python', blockId: block.id, source: block.source, block };
}

/** Parse typed text. Python entry needs identity creation; an existing anchor keeps its block. */
export function parseEntry(text: string, style?: CellStyle, locale?: string,
  options?: { readonly makeId: () => string; readonly existing?: CellContent | undefined }): CellContent {
  if (text === '') return blank;
  if (style?.numberFormat === '@') return { kind: 'text', value: text };
  if (isPythonEntry(text)) {
    if (options?.existing?.kind === 'python') return options.existing;
    requireDocument(options?.makeId, 'Python entry needs a workbook identity factory.');
    return pythonContent(options.makeId);
  }
  if (text.startsWith('=')) return { kind: 'formula', source: text };
  const trimmed = text.trim();
  if (numberPattern.test(trimmed) && Number.isFinite(Number(trimmed))) return { kind: 'number', value: Number(trimmed) };
  const serial = parseTemporalEntry(trimmed, style?.numberFormat, locale);
  if (serial !== undefined) return { kind: 'number', value: serial };
  if (/^(?:true|false)$/iu.test(trimmed)) return { kind: 'boolean', value: trimmed.toLowerCase() === 'true' };
  return { kind: 'text', value: text };
}

export type SelectionSpanKind = 'cells' | 'rows' | 'columns' | 'sheet';

export function rangeSpanKind(range: CellExtent): SelectionSpanKind {
  const allRows = range.row === 1 && range.rows === limits.rows;
  const allColumns = range.column === 1 && range.columns === limits.columns;
  return allRows && allColumns ? 'sheet' : allRows ? 'columns' : allColumns ? 'rows' : 'cells';
}

export function editorEntryText(cell: CellDocument | undefined, locale?: string, style: CellStyle | undefined = cell?.style): string {
  return cell?.content.kind === 'number' ?
    temporalEntry(cell.content.value, style?.numberFormat, locale) ?? entryText(cell) : entryText(cell);
}

export function entryText(cell: CellDocument | undefined): string {
  switch (cell?.content.kind) {
    case 'number':
    case 'boolean': return String(cell.content.value);
    case 'text': return cell.content.value;
    case 'formula': return cell.content.source;
    case 'python': return '=PY()';
    default: return '';
  }
}

function box(range: CellExtent): ReferenceBox {
  return { top: range.row, left: range.column, bottom: range.row + range.rows - 1, right: range.column + range.columns - 1 };
}
function extent({ top, left, bottom, right }: ReferenceBox): CellExtent {
  return { row: top, column: left, rows: bottom - top + 1, columns: right - left + 1 };
}
function includes(range: ReferenceBox, point: Coordinate): boolean {
  return point.row >= range.top && point.row <= range.bottom && point.column >= range.left && point.column <= range.right;
}
function contains(a: ReferenceBox, b: ReferenceBox): boolean {
  return includes(a, { row: b.top, column: b.left }) && includes(a, { row: b.bottom, column: b.right });
}
export function checkedSheet(basis: WorkbookSnapshot, range: EditRange): SheetDocument {
  const sheet = basis.document.sheets.find(sheet => sheet.id === range.sheetId);
  const { top, left, bottom, right } = box(range);
  requireDocument(sheet && [range.row, range.column, range.rows, range.columns].every(Number.isSafeInteger) &&
    top >= 1 && left >= 1 && range.rows >= 1 && range.columns >= 1 && bottom <= limits.rows && right <= limits.columns,
  'Selection is outside the worksheet.');
  return sheet;
}
function point(sheetId: string, row: number, column: number): CellReference { return { sheetId, row, column }; }
function command(label: string, edits: readonly CellEdit[]): SetCellsCommand { return { kind: 'set-cells', edits, label }; }

export function expandMerges(range: EditRange, sheet: SheetDocument): EditRange {
  let expanded = box(range);
  for (;;) {
    const previous = expanded;
    for (const merge of sheet.merges) {
      const area = box(merge);
      if (!intersection(expanded, area) || contains(expanded, area)) continue;
      expanded = { top: Math.min(expanded.top, area.top), left: Math.min(expanded.left, area.left),
        bottom: Math.max(expanded.bottom, area.bottom), right: Math.max(expanded.right, area.right) };
    }
    if (previous === expanded) return { sheetId: range.sheetId, ...extent(expanded) };
  }
}

export function captureClipboard(basis: WorkbookSnapshot, selection: EditRange, cut: boolean): ClipboardBlock {
  const sheet = checkedSheet(basis, selection);
  const range = expandMerges(selection, sheet);
  requireDocument(range.rows * range.columns <= limits.cells, 'Clipboard selection is too large.');
  const cells: (CellDocument | null)[] = [];
  for (let row = range.row; row < range.row + range.rows; row++) {
    for (let column = range.column; column < range.column + range.columns; column++) {
      const cell = findCell(sheet, { row, column });
      requireDocument(!cell?.spillOwner && !(cell?.content.kind === 'python' && cell.content.block?.spillActive),
        'Copy or move the complete Python spill.');
      cells.push(cell ?? null);
    }
  }
  return ownDto({ sessionId: basis.sessionId, authoredRevision: basis.authoredRevision, range, cells, cut,
    merges: sheet.merges.filter(merge => intersection(box(merge), box(range))) });
}

export function clipboardText(block: ClipboardBlock): string {
  const rows: string[] = [];
  let length = 0;
  for (let row = 0; row < block.range.rows; row++) {
    const fields = block.cells.slice(row * block.range.columns, (row + 1) * block.range.columns).map(cell => {
      const text = entryText(cell ?? undefined);
      return cell?.content.kind === 'text' || /[\t\r\n"]/u.test(text) ? '"' + text.replaceAll('"', '""') + '"' : text;
    });
    const text = fields.join('\t');
    length += text.length + (row === 0 ? 0 : 2);
    requireDocument(length <= clipboardTextLimit, 'Clipboard text exceeds 4 MiB.');
    rows.push(text);
  }
  return rows.join('\r\n');
}

interface Field { text: string; quoted: boolean }

function parseTsv(text: string): Field[][] {
  requireDocument(text.length <= clipboardTextLimit && !text.includes('\0'), 'Clipboard text is invalid or exceeds 4 MiB.');
  const rows: Field[][] = [];
  let row: Field[] = [];
  let value = '';
  let quoted = false;
  let state: 'start' | 'quoted' | 'closed' | 'plain' = 'start';
  function addCell() {
    row.push({ text: value, quoted });
    requireDocument(row.length <= limits.columns, 'Clipboard has too many columns.');
    value = ''; quoted = false; state = 'start';
  }
  function addRow() {
    addCell();
    requireDocument(rows.length < limits.rows && (!rows.length || rows[0]!.length === row.length) &&
      (rows.length + 1) * row.length <= limits.cells, 'Clipboard must be a bounded rectangle.');
    rows.push(row);
    row = [];
  }
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (state === 'quoted') {
      if (ch === '"' && text[i + 1] === '"') { value += '"'; i++; }
      else if (ch === '"') state = 'closed';
      else value += ch;
    } else if (state === 'start' && ch === '"') {
      quoted = true; state = 'quoted';
    } else if (ch === '\t') {
      addCell();
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      addRow();
    } else {
      requireDocument(state !== 'closed', 'Unexpected characters after quoted clipboard field.');
      value += ch;
      state = 'plain';
    }
  }
  requireDocument(state !== 'quoted', 'Unclosed quoted clipboard field.');
  if (row.length || value.length || quoted || !rows.length) addRow();
  return rows;
}

function destination(basis: WorkbookSnapshot, target: EditRange, rows: number, columns: number) {
  const sheet = checkedSheet(basis, target);
  const range = { ...target, rows, columns };
  checkedSheet(basis, range);
  const scalarInMerge = rows === 1 && columns === 1 && sheet.merges.some(merge =>
    merge.row === target.row && merge.column === target.column &&
    (target.rows === 1 && target.columns === 1 || target.rows === merge.rows && target.columns === merge.columns));
  requireDocument(scalarInMerge || target.rows === 1 && target.columns === 1 ||
    target.rows === rows && target.columns === columns, 'Paste selection must match clipboard dimensions or be one cell.');
  return { sheet, range, scalarInMerge };
}

export function pasteText(basis: WorkbookSnapshot, target: EditRange, text: string, mode: PasteMode, locale: string,
  makeId: () => string): SetCellsCommand {
  const fields = parseTsv(text);
  const { sheet, range } = destination(basis, target, fields.length, fields[0]!.length);
  const edits: CellEdit[] = [];
  fields.forEach((row, y) => row.forEach((field, x) => {
    const formula = !field.quoted && field.text.startsWith('=');
    if (mode === 'formulas' && !formula) return;
    const at = point(range.sheetId, range.row + y, range.column + x);
    const content: CellContent = field.quoted || mode === 'values' && formula ? { kind: 'text', value: field.text } :
      parseEntry(field.text, mode === 'formulas' ? undefined : resolvedCellStyle(sheet, at), locale, { makeId });
    edits.push({ target: at, content, resetCache: true });
  }));
  return command('Paste text', edits);
}

function valueOf(cell: CellDocument | null): CellContent {
  if (cell?.content.kind === 'python') {
    const value = cell.content.block?.lastValue;
    if (!value) return blank;
    requireDocument(value.kind !== 'error', 'Paste Values requires a valid Python result.');
    if (value.kind === 'number') return { kind: 'number', value: value.value };
    if (value.kind === 'boolean') return { kind: 'boolean', value: value.value };
    return { kind: 'text', value: value.value };
  }
  if (cell?.content.kind !== 'formula') return cell?.content ?? blank;
  requireDocument(cell.cache && cell.cache.kind !== 'error', 'Paste Values requires a valid scalar formula result.');
  return cell.cache;
}

interface Transfer {
  readonly rewrite: (source: string) => string;
  readonly rewritePython: (block: PythonBlock) => PythonBlock;
  readonly cut: boolean;
}

function transferredCell(original: CellDocument | null, target: CellReference, mode: PasteMode,
  transfer: Transfer): CellEdit | undefined {
  if (mode === 'formulas' && original?.content.kind !== 'formula') return;
  let content = mode === 'values' ? valueOf(original) : original?.content ?? blank;
  if (content.kind === 'formula') content = { kind: 'formula', source: transfer.rewrite(content.source) };
  if (content.kind === 'python') {
    const block = transfer.rewritePython(content.block ?? { id: content.blockId, source: content.source, bindings: [] });
    content = { kind: 'python', blockId: block.id, source: block.source, block };
  }
  return {
    target, content, preserveStyle: mode !== 'all', resetCache: true,
    ...(mode === 'all' && original?.style && { style: original.style }),
    ...(transfer.cut && original && { cellId: original.id }),
    ...(content.kind === 'formula' && original?.requiresExternalCalculation && { requiresExternalCalculation: true }),
  };
}

export function pasteBlock(basis: WorkbookSnapshot, target: EditRange, block: ClipboardBlock, mode: PasteMode,
  makeId: () => string): SetCellsCommand {
  requireDocument(basis.sessionId === block.sessionId, 'The internal clipboard belongs to another workbook session.');
  requireDocument(!block.cut || mode === 'all' && block.authoredRevision === basis.authoredRevision,
    'Cut source changed or paste mode is not All. Cut the selection again.');
  const { sheet, range, scalarInMerge } = destination(basis, target, block.range.rows, block.range.columns);
  const sourceSheet = checkedSheet(basis, block.range);
  const from = box(block.range), into = box(range);
  const dr = range.row - block.range.row, dc = range.column - block.range.column;
  if (block.cut && sourceSheet === sheet && dr === 0 && dc === 0) return command('Move cells', []);
  const move: ReferenceMove = { sourceSheet: sourceSheet.name, destinationSheet: sheet.name, from, dr, dc };
  const transfer: Transfer = {
    cut: block.cut,
    rewrite: source => block.cut ? moveReferences(source, sourceSheet.name, sheet.name, move) : copyReferences(source, dr, dc),
    rewritePython: prior => {
      if (block.cut) return transformPythonBlock(prior, reference =>
        moveReferences(reference, sourceSheet.name, sheet.name, move));
      const id = makeId();
      const translated = transformPythonBlock(prior, reference => copyReferences(reference, dr, dc));
      return { ...translated, id, status: 'never', placements: [] };
    },
  };
  const edits = new Map<string, CellEdit>();
  const add = (edit: CellEdit) => edits.set(edit.target.sheetId + ':' + cellKey(edit.target), edit);

  if (block.cut) {
    for (const original of block.cells) if (original) {
      add({ target: point(sourceSheet.id, original.row, original.column), content: blank, preserveStyle: false });
    }
  }
  block.cells.forEach((original, index) => {
    const at = point(sheet.id, range.row + Math.floor(index / range.columns), range.column + index % range.columns);
    const edit = transferredCell(original, at, mode, transfer);
    if (edit) add(edit);
  });

  if (block.cut) {
    for (const current of basis.document.sheets) for (const cell of current.cells) {
      if (cell.content.kind !== 'formula' || current.id === sourceSheet.id && includes(from, cell) ||
        current.id === sheet.id && includes(into, cell)) continue;
      const source = moveReferences(cell.content.source, current.name, current.name, move);
      if (source !== cell.content.source) add({
        target: point(current.id, cell.row, cell.column), content: { kind: 'formula', source }, resetCache: true,
        ...(cell.requiresExternalCalculation && { requiresExternalCalculation: true }),
      });
    }
  }

  const merges = new Map<string, readonly CellExtent[]>();
  if (mode === 'all' && !scalarInMerge) {
    if (block.cut && block.merges.length) {
      merges.set(sourceSheet.id, sourceSheet.merges.filter(merge => !intersection(box(merge), from)));
    }
    const previous = merges.get(sheet.id) ?? sheet.merges;
    const remaining = previous.filter(merge => {
      if (!intersection(box(merge), into)) return true;
      requireDocument(contains(into, box(merge)), 'Select the complete merged area before pasting.');
      return false;
    });
    if (remaining.length !== previous.length || block.merges.length) {
      merges.set(sheet.id, [...remaining, ...block.merges.map(merge => ({ ...merge, row: merge.row + dr, column: merge.column + dc }))]);
    }
  }
  const names = block.cut ? basis.document.names.map(name => {
    if (name.area.invalid || name.area.sheetId !== sourceSheet.id || !intersection(box(name.area), from)) return name;
    if (!contains(from, box(name.area))) return { ...name, area: { ...name.area, invalid: true } };
    return { ...name, area: { ...name.area, sheetId: sheet.id, row: name.area.row + dr, column: name.area.column + dc } };
  }) : undefined;
  return { ...command(block.cut ? 'Move cells' : 'Paste cells', [...edits.values()]),
    ...(names && { names }), ...(merges.size > 0 && { merges: [...merges].map(([sheetId, ranges]) => ({ sheetId, ranges })) }),
  };
}

export function fill(basis: WorkbookSnapshot, range: EditRange, direction: FillDirection, makeId: () => string): SetCellsCommand {
  const sheet = checkedSheet(basis, range);
  requireDocument(!sheet.merges.some(merge => intersection(box(merge), box(range))), 'Unmerge cells before filling this range.');
  const vertical = direction === 'down' || direction === 'up';
  const reverse = direction === 'up' || direction === 'left';
  const edgeRow = range.row + (direction === 'up' ? range.rows - 1 : 0);
  const edgeColumn = range.column + (direction === 'left' ? range.columns - 1 : 0);
  const edits: CellEdit[] = [];
  for (let distance = 1; distance < (vertical ? range.rows : range.columns); distance++) {
    const shift = reverse ? -distance : distance;
    const dr = vertical ? shift : 0, dc = vertical ? 0 : shift;
    for (let i = 0; i < (vertical ? range.columns : range.rows); i++) {
      const row = edgeRow + (vertical ? 0 : i), column = edgeColumn + (vertical ? i : 0);
      const original = findCell(sheet, { row, column });
      requireDocument(!original?.spillOwner && !(original?.content.kind === 'python' && original.content.block?.spillActive),
        'Fill the complete Python spill.');
      const at = point(sheet.id, row + dr, column + dc);
      edits.push(transferredCell(original ?? null, at, 'all', {
        cut: false, rewrite: source => copyReferences(source, dr, dc),
        rewritePython: prior => ({ ...transformPythonBlock(prior, reference => copyReferences(reference, dr, dc)),
          id: makeId(), status: 'never', placements: [] }),
      })!);
    }
  }
  return command('Fill ' + direction, edits);
}

function stripBorderSides(style: CellStyle, sides: readonly ('top' | 'bottom' | 'left' | 'right')[]): CellStyle | undefined {
  if (!style.borderDetails && !style.borders) return style;
  const details = { ...style.borderDetails };
  let legacy = style.borders ?? '';
  const originalLegacy = legacy;
  let changed = false;
  for (const side of sides) {
    if (details[side]) { delete details[side]; changed = true; }
    legacy = legacy.replace(side[0]!, '');
  }
  if (!changed && legacy === originalLegacy) return style;
  const { borders: _legacy, borderDetails: _details, ...remaining } = style;
  const result = { ...remaining, ...(legacy && { borders: legacy }),
    ...(Object.keys(details).length > 0 && { borderDetails: details }) };
  return Object.keys(result).length ? result : undefined;
}

function reconcileClearedStyle(style: CellStyle | undefined, inherited: CellStyle | undefined): CellStyle | undefined {
  const normalized = style && Object.keys(style).length ? style : undefined;
  if (normalized) return equalDto(normalized, inherited) ? undefined : normalized;
  return inherited ? {} : undefined;
}

export function clear(basis: WorkbookSnapshot, selections: readonly EditRange[], mode: ClearMode): WorkbookCommand {
  const label = mode === 'all' ? 'Clear All' : mode === 'contents' ? 'Clear Contents' : 'Clear Formats';
  const bySheet = new Map<SheetDocument, EditRange[]>();
  let hasAxisSpan = false;
  for (const selection of selections) {
    const sheet = checkedSheet(basis, selection);
    const area = expandMerges(selection, sheet);
    if (mode !== 'contents' && rangeSpanKind(area) !== 'cells') hasAxisSpan = true;
    const areas = bySheet.get(sheet) ?? [];
    areas.push(area);
    bySheet.set(sheet, areas);
  }
  if (!hasAxisSpan) {
    const edits: CellEdit[] = [];
    for (const [sheet, sheetAreas] of bySheet) {
      const areas = sheetAreas.map(box);
      const cellMap = new Map(sheet.cells.map(cell => [cellKey(cell), cell]));
      const touched = new Set<number>();
      for (const cell of sheet.cells) {
        const target = point(sheet.id, cell.row, cell.column);
        if (areas.some(area => includes(area, cell))) {
          touched.add(cellKey(cell));
          if (mode === 'contents') {
            edits.push({ target, content: blank, preserveStyle: true });
          } else {
            const inherited = findRowStyle(sheet, cell.row) ?? findColumnStyle(sheet, cell.column);
            edits.push({ target, content: mode === 'formats' ? cell.content : blank, preserveStyle: false,
              ...(inherited && { style: {} }) });
          }
        } else if (mode !== 'contents') {
          const effective = resolvedCellStyle(sheet, cell, cell);
          if (!effective) continue;
          const merge = sheet.merges.find(merge => merge.row === cell.row && merge.column === cell.column);
          const stripped = clearAdjacentBorders(effective, box(merge ?? { ...cell, rows: 1, columns: 1 }), areas);
          const inherited = findRowStyle(sheet, cell.row) ?? findColumnStyle(sheet, cell.column);
          const next = reconcileClearedStyle(stripped, inherited);
          if (!equalDto(next, cell.style)) {
            touched.add(cellKey(cell));
            edits.push({ target, content: cell.content, preserveStyle: false, ...(next !== undefined && { style: next }) });
          }
        }
      }
      if (mode !== 'contents' && (sheet.rows.some(r => r.style) || sheet.columns.some(c => c.style))) {
        for (const area of areas) {
          const hasStyledAxis = (() => {
            for (const r of sheet.rows) if (r.style && r.index >= area.top && r.index <= area.bottom) return true;
            for (const c of sheet.columns) if (c.style && c.index >= area.left && c.index <= area.right) return true;
            return false;
          })();
          if (hasStyledAxis) {
            const areaRows = area.bottom - area.top + 1;
            const areaCols = area.right - area.left + 1;
            requireDocument(areaRows * areaCols <= limits.cells, 'Format selection is too large.');
            for (let row = area.top; row <= area.bottom; row++) {
              for (let column = area.left; column <= area.right; column++) {
                const key = cellKey({ row, column });
                if (touched.has(key)) continue;
                const inherited = findRowStyle(sheet, row) ?? findColumnStyle(sheet, column);
                if (!inherited) continue;
                touched.add(key);
                edits.push({ target: point(sheet.id, row, column), content: blank, preserveStyle: false, style: {} });
              }
            }
          }
          // Clear adjacent borders on unpopulated cells inheriting an axis border along the outer edges.
          const perimeterEdges: { row: number; column: number; side: 'top' | 'bottom' | 'left' | 'right' }[] = [];
          if (area.top > 1 && area.right - area.left + 1 <= limits.cells) {
            for (let c = area.left; c <= area.right; c++) perimeterEdges.push({ row: area.top - 1, column: c, side: 'bottom' });
          }
          if (area.bottom < limits.rows && area.right - area.left + 1 <= limits.cells) {
            for (let c = area.left; c <= area.right; c++) perimeterEdges.push({ row: area.bottom + 1, column: c, side: 'top' });
          }
          if (area.left > 1 && area.bottom - area.top + 1 <= limits.cells) {
            for (let r = area.top; r <= area.bottom; r++) perimeterEdges.push({ row: r, column: area.left - 1, side: 'right' });
          }
          if (area.right < limits.columns && area.bottom - area.top + 1 <= limits.cells) {
            for (let r = area.top; r <= area.bottom; r++) perimeterEdges.push({ row: r, column: area.right + 1, side: 'left' });
          }
          for (const { row, column, side } of perimeterEdges) {
            const key = cellKey({ row, column });
            if (touched.has(key) || cellMap.has(key)) continue;
            const inherited = findRowStyle(sheet, row) ?? findColumnStyle(sheet, column);
            if (!inherited?.borderDetails && !inherited?.borders) continue;
            const stripped = stripBorderSides(inherited, [side]);
            if (equalDto(stripped, inherited)) continue;
            touched.add(key);
            const next = reconcileClearedStyle(stripped, inherited);
            edits.push({ target: point(sheet.id, row, column), content: blank, preserveStyle: false,
              ...(next !== undefined && { style: next }) });
          }
        }
      }
    }
    return command(label, edits);
  }

  const edits: CellEdit[] = [];
  const nextSheets = new Map<string, SheetDocument>();
  for (const [sheet, sheetAreas] of bySheet) {
    const areas = sheetAreas.map(box);
    const rowDims = new Map(sheet.rows.map(dim => [dim.index, dim]));
    const colDims = new Map(sheet.columns.map(dim => [dim.index, dim]));
    const clearedCellKeys = new Set<number>();
    const pending = new Map<number, { row: number; column: number; prior: CellDocument | undefined; content: CellContent; style: CellStyle | undefined }>();

    function recordPending(row: number, column: number, content: CellContent, style: CellStyle | undefined) {
      const key = cellKey({ row, column });
      const existing = pending.get(key);
      const prior = existing ? existing.prior : findCell(sheet, { row, column });
      pending.set(key, { row, column, prior, content, style });
      requireDocument(pending.size <= limits.cells, 'Format selection is too large.');
    }

    for (const area of areas) {
      const span = rangeSpanKind({ row: area.top, column: area.left, rows: area.bottom - area.top + 1, columns: area.right - area.left + 1 });
      if (span === 'sheet') {
        for (const [idx, dim] of rowDims) {
          if (dim.style === undefined) continue;
          const { style: _removed, ...rest } = dim;
          if (rest.size === undefined && !rest.hidden) rowDims.delete(idx);
          else rowDims.set(idx, rest);
        }
        for (const [idx, dim] of colDims) {
          if (dim.style === undefined) continue;
          const { style: _removed, ...rest } = dim;
          if (rest.size === undefined && !rest.hidden) colDims.delete(idx);
          else colDims.set(idx, rest);
        }
      } else if (span === 'columns') {
        for (let c = area.left; c <= area.right; c++) {
          const dim = colDims.get(c);
          if (dim?.style !== undefined) {
            const { style: _removed, ...rest } = dim;
            if (rest.size === undefined && !rest.hidden) colDims.delete(c);
            else colDims.set(c, rest);
          }
        }
        for (const [r] of rowDims) {
          if (rowDims.get(r)?.style === undefined) continue;
          for (let c = area.left; c <= area.right; c++) {
            const key = cellKey({ row: r, column: c });
            clearedCellKeys.add(key);
            const existing = pending.get(key)?.prior ?? findCell(sheet, { row: r, column: c });
            const content = mode === 'formats' ? existing?.content ?? blank : blank;
            recordPending(r, c, content, undefined);
          }
        }
        if (area.left > 1) {
          const prevDim = colDims.get(area.left - 1);
          if (prevDim?.style) {
            const stripped = stripBorderSides(prevDim.style, ['right']);
            const { style: _old, ...rest } = prevDim;
            const nextDim: AxisDimension = { ...rest, ...(stripped && { style: stripped }) };
            if (nextDim.size === undefined && !nextDim.hidden && !nextDim.style) colDims.delete(prevDim.index);
            else colDims.set(prevDim.index, nextDim);
          }
        }
        if (area.right < limits.columns) {
          const nextAxis = colDims.get(area.right + 1);
          if (nextAxis?.style) {
            const stripped = stripBorderSides(nextAxis.style, ['left']);
            const { style: _old, ...rest } = nextAxis;
            const nextDim: AxisDimension = { ...rest, ...(stripped && { style: stripped }) };
            if (nextDim.size === undefined && !nextDim.hidden && !nextDim.style) colDims.delete(nextAxis.index);
            else colDims.set(nextAxis.index, nextDim);
          }
        }
      } else if (span === 'rows') {
        for (let r = area.top; r <= area.bottom; r++) {
          const dim = rowDims.get(r);
          if (dim?.style !== undefined) {
            const { style: _removed, ...rest } = dim;
            if (rest.size === undefined && !rest.hidden) rowDims.delete(r);
            else rowDims.set(r, rest);
          }
        }
        for (const [c] of colDims) {
          if (colDims.get(c)?.style === undefined) continue;
          for (let r = area.top; r <= area.bottom; r++) {
            const key = cellKey({ row: r, column: c });
            clearedCellKeys.add(key);
            const existing = pending.get(key)?.prior ?? findCell(sheet, { row: r, column: c });
            const content = mode === 'formats' ? existing?.content ?? blank : blank;
            recordPending(r, c, content, undefined);
          }
        }
        if (area.top > 1) {
          const prevDim = rowDims.get(area.top - 1);
          if (prevDim?.style) {
            const stripped = stripBorderSides(prevDim.style, ['bottom']);
            const { style: _old, ...rest } = prevDim;
            const nextDim: AxisDimension = { ...rest, ...(stripped && { style: stripped }) };
            if (nextDim.size === undefined && !nextDim.hidden && !nextDim.style) rowDims.delete(prevDim.index);
            else rowDims.set(prevDim.index, nextDim);
          }
        }
        if (area.bottom < limits.rows) {
          const nextAxis = rowDims.get(area.bottom + 1);
          if (nextAxis?.style) {
            const stripped = stripBorderSides(nextAxis.style, ['top']);
            const { style: _old, ...rest } = nextAxis;
            const nextDim: AxisDimension = { ...rest, ...(stripped && { style: stripped }) };
            if (nextDim.size === undefined && !nextDim.hidden && !nextDim.style) rowDims.delete(nextAxis.index);
            else rowDims.set(nextAxis.index, nextDim);
          }
        }
      } else {
        const areaRows = area.bottom - area.top + 1;
        const areaCols = area.right - area.left + 1;
        requireDocument(areaRows * areaCols <= limits.cells, 'Format selection is too large.');
        for (let r = area.top; r <= area.bottom; r++) {
          for (let c = area.left; c <= area.right; c++) {
            const key = cellKey({ row: r, column: c });
            clearedCellKeys.add(key);
            const existing = pending.get(key)?.prior ?? findCell(sheet, { row: r, column: c });
            const content = mode === 'formats' ? existing?.content ?? blank : blank;
            recordPending(r, c, content, undefined);
          }
        }
      }
    }

    for (const cell of sheet.cells) {
      const key = cellKey(cell);
      if (areas.some(area => includes(area, cell))) {
        clearedCellKeys.add(key);
        const content = mode === 'formats' ? cell.content : blank;
        recordPending(cell.row, cell.column, content, undefined);
      } else if (!clearedCellKeys.has(key) && cell.style) {
        const merge = sheet.merges.find(merge => merge.row === cell.row && merge.column === cell.column);
        const style = clearAdjacentBorders(cell.style, box(merge ?? { ...cell, rows: 1, columns: 1 }), areas);
        if (style !== cell.style) recordPending(cell.row, cell.column, cell.content, style);
      }
    }

    for (const { row, column, prior, content, style } of pending.values()) {
      const inherited = rowDims.get(row)?.style ?? colDims.get(column)?.style;
      const finalStyle = reconcileClearedStyle(style, inherited);
      if (equalDto(content, prior?.content ?? blank) && equalDto(finalStyle, prior?.style)) continue;
      edits.push({ target: point(sheet.id, row, column), content, preserveStyle: false,
        ...(finalStyle !== undefined && { style: finalStyle }) });
    }

    nextSheets.set(sheet.id, {
      ...sheet,
      rows: [...rowDims.values()].sort((a, b) => a.index - b.index),
      columns: [...colDims.values()].sort((a, b) => a.index - b.index),
    });
  }
  return {
    kind: 'structure',
    label,
    document: { ...basis.document, sheets: basis.document.sheets.map(s => nextSheets.get(s.id) ?? s) },
    edits,
  };
}

function clearAdjacentBorders(style: CellStyle, cell: ReferenceBox, areas: readonly ReferenceBox[]): CellStyle | undefined {
  if (!style.borderDetails && !style.borders) return style;
  const details = { ...style.borderDetails };
  let legacy = style.borders ?? '';
  const originalLegacy = legacy;
  let changed = false;
  function remove(side: 'top' | 'bottom' | 'left' | 'right', letter: string) {
    if (details[side]) { delete details[side]; changed = true; }
    legacy = legacy.replace(letter, '');
  }
  for (const area of areas) {
    if (area.left <= cell.right && area.right >= cell.left) {
      if (area.bottom === cell.top - 1) remove('top', 't');
      if (area.top === cell.bottom + 1) remove('bottom', 'b');
    }
    if (area.top <= cell.bottom && area.bottom >= cell.top) {
      if (area.right === cell.left - 1) remove('left', 'l');
      if (area.left === cell.right + 1) remove('right', 'r');
    }
  }
  if (!changed && legacy === originalLegacy) return style;
  const { borders: _legacy, borderDetails: _details, ...remaining } = style;
  const result = { ...remaining, ...(legacy && { borders: legacy }),
    ...(Object.keys(details).length > 0 && { borderDetails: details }) };
  return Object.keys(result).length ? result : undefined;
}

export function autoSum(basis: WorkbookSnapshot, range: EditRange, aggregate: Aggregate): { command: SetCellsCommand; target: CellReference } {
  const sheet = checkedSheet(basis, range);
  requireDocument(range.rows === 1 || range.columns === 1, 'AutoSum accepts one row or column.');
  const { bottom, right } = box(range);
  let row = bottom, column = right;
  let first = { row: range.row, column: range.column };
  let last: Coordinate;
  const at = (row: number, column: number) => findCell(sheet, { row, column });
  const occupied = (row: number, column: number) => !!at(row, column) && at(row, column)!.content.kind !== 'blank';
  const numeric = (row: number, column: number) => {
    const cell = at(row, column);
    return cell?.content.kind === 'number' || cell?.content.kind === 'formula' && cell.cache?.kind === 'number';
  };
  if (range.rows > 1) {
    if (occupied(row, column)) row++;
    last = { row: row - 1, column };
  } else if (range.columns > 1) {
    if (occupied(row, column)) column++;
    last = { row, column: column - 1 };
  } else if (row > 1 && numeric(row - 1, column)) {
    first.row = row - 1;
    while (first.row > 1 && numeric(first.row - 1, column)) first.row--;
    last = { row: row - 1, column };
  } else {
    requireDocument(column > 1 && numeric(row, column - 1), 'Select a numeric range or a blank cell below or beside contiguous numbers.');
    first.column = column - 1;
    while (first.column > 1 && numeric(row, first.column - 1)) first.column--;
    last = { row, column: column - 1 };
  }
  requireDocument(row <= worksheetViewport.rows && column <= worksheetViewport.columns, 'AutoSum destination is outside the current grid view.');
  requireDocument(!occupied(row, column), 'AutoSum destination must be blank.');
  const target = point(sheet.id, row, column);
  return { target, command: command('AutoSum', [{ target, content: {
    kind: 'formula', source: '=' + aggregate + '(' + formulaAddress(first) + ':' + formulaAddress(last) + ')',
  } }]) };
}
