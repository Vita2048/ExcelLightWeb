import { cellKey, documentLimits as limits, findCell, ownDto, worksheetViewport,
  type CellContent, type CellDocument, type CellExtent, type CellReference, type CellStyle, type Coordinate, type SheetDocument,
} from '../core/workbook';
import { requireDocument } from '../core/workbook-validation';
import { copyReferences, formulaAddress, intersection, moveReferences, type ReferenceBox, type ReferenceMove } from '../core/formula-references';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { CellEdit, SetCellsCommand } from './workbook-session';

export type PasteMode = 'all' | 'values' | 'formulas';
export type ClearMode = 'all' | 'contents' | 'formats';
export type FillDirection = 'down' | 'right' | 'up' | 'left';
export type Aggregate = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN';
export interface EditRange extends CellExtent { readonly sheetId: string }
export interface ClipboardBlock {
  readonly sessionId: string;
  readonly revision: number;
  readonly range: EditRange;
  readonly cells: readonly (CellDocument | null)[];
  readonly merges: readonly CellExtent[];
  readonly cut: boolean;
}
export const clipboardTextLimit = 4 * 1024 * 1024;
const blank = Object.freeze({ kind: 'blank' } as const);
const numberPattern = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/iu;

export function parseEntry(text: string, style?: CellStyle): CellContent {
  if (text === '') return blank;
  if (style?.numberFormat === '@') return { kind: 'text', value: text };
  requireDocument(!/^=PY(?:\(\))?$/iu.test(text.trim()), 'Python authoring is unavailable.');
  if (text.startsWith('=')) return { kind: 'formula', source: text };
  const trimmed = text.trim();
  if (numberPattern.test(trimmed) && Number.isFinite(Number(trimmed))) return { kind: 'number', value: Number(trimmed) };
  if (/^(?:true|false)$/iu.test(trimmed)) return { kind: 'boolean', value: trimmed.toLowerCase() === 'true' };
  return { kind: 'text', value: text };
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
function checkedSheet(basis: WorkbookSnapshot, range: EditRange): SheetDocument {
  const sheet = basis.document.sheets.find(sheet => sheet.id === range.sheetId);
  const { top, left, bottom, right } = box(range);
  requireDocument(sheet && [range.row, range.column, range.rows, range.columns].every(Number.isSafeInteger) &&
    top >= 1 && left >= 1 && range.rows >= 1 && range.columns >= 1 && bottom <= limits.rows && right <= limits.columns,
  'Selection is outside the worksheet.');
  return sheet;
}
function point(sheetId: string, row: number, column: number): CellReference { return { sheetId, row, column }; }
function command(label: string, edits: readonly CellEdit[]): SetCellsCommand { return { kind: 'set-cells', edits, label }; }

function expandMerges(range: EditRange, sheet: SheetDocument): EditRange {
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
      requireDocument(cell?.content.kind !== 'python', 'Python clipboard operations are unavailable.');
      cells.push(cell ?? null);
    }
  }
  return ownDto({ sessionId: basis.sessionId, revision: basis.revision, range, cells, cut,
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

export function pasteText(basis: WorkbookSnapshot, target: EditRange, text: string, mode: PasteMode): SetCellsCommand {
  const fields = parseTsv(text);
  const { sheet, range } = destination(basis, target, fields.length, fields[0]!.length);
  const edits: CellEdit[] = [];
  fields.forEach((row, y) => row.forEach((field, x) => {
    const formula = !field.quoted && field.text.startsWith('=');
    if (mode === 'formulas' && !formula) return;
    const at = point(range.sheetId, range.row + y, range.column + x);
    const content: CellContent = field.quoted || mode === 'values' && formula ? { kind: 'text', value: field.text } :
      parseEntry(field.text, mode === 'formulas' ? undefined : findCell(sheet, at)?.style);
    edits.push({ target: at, content, resetCache: true });
  }));
  return command('Paste text', edits);
}

function valueOf(cell: CellDocument | null): CellContent {
  if (cell?.content.kind !== 'formula') return cell?.content ?? blank;
  requireDocument(cell.cache && cell.cache.kind !== 'error', 'Paste Values requires a valid scalar formula result.');
  return cell.cache;
}

function transferredCell(original: CellDocument | null, target: CellReference, mode: PasteMode,
  rewrite: (source: string) => string, cut: boolean): CellEdit | undefined {
  if (mode === 'formulas' && original?.content.kind !== 'formula') return;
  let content = mode === 'values' ? valueOf(original) : original?.content ?? blank;
  if (content.kind === 'formula') content = { kind: 'formula', source: rewrite(content.source) };
  return {
    target, content, preserveStyle: mode !== 'all', resetCache: true,
    ...(mode === 'all' && original?.style && { style: original.style }),
    ...(cut && original && { cellId: original.id }),
    ...(content.kind === 'formula' && original?.requiresExternalCalculation && { requiresExternalCalculation: true }),
  };
}

export function pasteBlock(basis: WorkbookSnapshot, target: EditRange, block: ClipboardBlock, mode: PasteMode): SetCellsCommand {
  requireDocument(basis.sessionId === block.sessionId, 'The internal clipboard belongs to another workbook session.');
  requireDocument(!block.cut || mode === 'all' && block.revision === basis.revision,
    'Cut source changed or paste mode is not All. Cut the selection again.');
  const { sheet, range, scalarInMerge } = destination(basis, target, block.range.rows, block.range.columns);
  const sourceSheet = checkedSheet(basis, block.range);
  const from = box(block.range), into = box(range);
  const dr = range.row - block.range.row, dc = range.column - block.range.column;
  if (block.cut && sourceSheet === sheet && dr === 0 && dc === 0) return command('Move cells', []);
  const move: ReferenceMove = { sourceSheet: sourceSheet.name, destinationSheet: sheet.name, from, dr, dc };
  const edits = new Map<string, CellEdit>();
  const add = (edit: CellEdit) => edits.set(edit.target.sheetId + ':' + cellKey(edit.target), edit);

  if (block.cut) {
    for (const original of block.cells) if (original) {
      add({ target: point(sourceSheet.id, original.row, original.column), content: blank, preserveStyle: false });
    }
  }
  block.cells.forEach((original, index) => {
    const at = point(sheet.id, range.row + Math.floor(index / range.columns), range.column + index % range.columns);
    const edit = transferredCell(original, at, mode, source =>
      block.cut ? moveReferences(source, sourceSheet.name, sheet.name, move) : copyReferences(source, dr, dc), block.cut);
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

export function fill(basis: WorkbookSnapshot, range: EditRange, direction: FillDirection): SetCellsCommand {
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
      requireDocument(original?.content.kind !== 'python', 'Python fill is unavailable.');
      const at = point(sheet.id, row + dr, column + dc);
      edits.push(transferredCell(original ?? null, at, 'all', source => copyReferences(source, dr, dc), false)!);
    }
  }
  return command('Fill ' + direction, edits);
}

export function clear(basis: WorkbookSnapshot, selections: readonly EditRange[], mode: ClearMode): SetCellsCommand {
  const bySheet = new Map<SheetDocument, ReferenceBox[]>();
  for (const selection of selections) {
    const sheet = checkedSheet(basis, selection);
    const areas = bySheet.get(sheet) ?? [];
    areas.push(box(expandMerges(selection, sheet)));
    bySheet.set(sheet, areas);
  }
  const edits: CellEdit[] = [];
  for (const [sheet, areas] of bySheet) {
    for (const cell of sheet.cells) {
      const target = point(sheet.id, cell.row, cell.column);
      if (areas.some(area => includes(area, cell))) {
        edits.push({ target, content: mode === 'formats' ? cell.content : blank, preserveStyle: mode === 'contents' });
      } else if (mode !== 'contents' && cell.style) {
        const merge = sheet.merges.find(merge => merge.row === cell.row && merge.column === cell.column);
        const style = clearAdjacentBorders(cell.style, box(merge ?? { ...cell, rows: 1, columns: 1 }), areas);
        if (style !== cell.style) edits.push({ target, content: cell.content, preserveStyle: false, ...(style && { style }) });
      }
    }
  }
  return command(mode === 'all' ? 'Clear All' : mode === 'contents' ? 'Clear Contents' : 'Clear Formats', edits);
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
