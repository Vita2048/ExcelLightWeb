import type { WorkbookSnapshot } from '../application/workbook-protocol';
import { parseFormula } from '../core/formula-references';
import { excelDateParts } from '../core/date-serial-probe';
import { temporalParts } from '../core/formats/format-code';
import { DocumentError, documentLimits, findCell, resolvedCellStyle, type CellCache, type CellDocument, type CellExtent, type PythonBlock, type SheetDocument, type WorkbookDocument, type WorkbookName } from '../core/workbook';
import { tableDesign, tableInputExtent } from '../core/workbook-names';
import { ordinalNameKey, requireDocument } from '../core/workbook-validation';
import { discoverPythonReferences, pythonNameKey } from './references';
import type { PythonInputTable, PythonRuntimeInput, PythonInputValue } from './runtime-protocol';

const frameLimit = 32 * 1024 * 1024;
const inputCellLimit = 100_000;

export interface PythonInputAnchor {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
  readonly blockId: string;
  readonly source: string;
}

export interface BuiltPythonInputs {
  readonly input: PythonRuntimeInput;
  readonly fingerprint: string;
}

interface ResolvedRange {
  readonly area: CellExtent & { readonly sheetId: string };
  readonly named?: WorkbookName;
  readonly dataOnlyTable: boolean;
}

function resolveRange(document: WorkbookDocument, localSheetId: string, reference: string): ResolvedRange {
  const namedKey = pythonNameKey(reference);
  const named = document.names.find(name => name.name.toLocaleLowerCase('en-US') === namedKey.toLocaleLowerCase('en-US'));
  if (named) {
    requireDocument(!named.area.invalid && (!reference.includes('[') || named.isTable),
      named.area.invalid ? `Workbook name '${named.name}' has a deleted or split extent; repair it first.` : 'Structured selectors require a table name.');
    let area = named.area;
    const dataOnlyTable = Boolean(named.isTable && reference.toLowerCase().endsWith('[#data]'));
    if (dataOnlyTable) {
      const input = tableInputExtent(named);
      area = { ...area, row: input.row + Number(input.headerRow), rows: input.rows - Number(input.headerRow) };
    }
    requireDocument(area.rows >= 1 && area.columns >= 1, 'Python input range is empty.');
    return { area, named, dataOnlyTable };
  }

  const parsed = parseFormula('=' + reference).expression;
  requireDocument(parsed.kind === 'reference' && parsed.reference.first.row !== undefined && parsed.reference.first.column !== undefined,
    `Python reference '${reference}' is not a supported cell range, workbook name or table.`);
  const referencePart = parsed.reference;
  const sheet = referencePart.sheet === undefined ? document.sheets.find(item => item.id === localSheetId) :
    document.sheets.find(item => ordinalNameKey(item.name) === ordinalNameKey(referencePart.sheet!));
  requireDocument(sheet, `Python input sheet '${referencePart.sheet ?? ''}' does not exist.`);
  const firstRow = referencePart.first.row!;
  const firstColumn = referencePart.first.column!;
  const lastRow = referencePart.last.row ?? firstRow;
  const lastColumn = referencePart.last.column ?? firstColumn;
  const area = { sheetId: sheet.id, row: Math.min(firstRow, lastRow), column: Math.min(firstColumn, lastColumn),
    rows: Math.abs(lastRow - firstRow) + 1, columns: Math.abs(lastColumn - firstColumn) + 1 };
  requireDocument(area.row + area.rows - 1 <= documentLimits.rows && area.column + area.columns - 1 <= documentLimits.columns,
    'Python input range exceeds worksheet limits.');
  return { area, dataOnlyTable: false };
}

/** Resolve the ranges read by one Python anchor for mixed dependency analysis. */
export function pythonInputRanges(document: WorkbookDocument, localSheetId: string, source: string,
  bindings: PythonBlock['bindings'] = []): readonly (CellExtent & { readonly sheetId: string })[] {
  const references = discoverPythonReferences(source).map(item => item.reference);
  const raw = [...references, ...bindings.map(binding => binding.reference.replace(/^=/u, ''))];
  return raw.map(reference => resolveRange(document, localSheetId, reference).area);
}

function tableHeaders(document: WorkbookDocument, name: WorkbookName): string[] {
  const design = tableDesign(name);
  if (!design.headerRow) {
    const labels = name.headerLabels ?? [];
    requireDocument(labels.length === name.area.columns, 'Python table headers are unavailable; repair the table first.');
    validateHeaders(labels, name.name);
    return [...labels];
  }
  const sheet = document.sheets.find(item => item.id === name.area.sheetId)!;
  const labels = Array.from({ length: name.area.columns }, (_, index) => {
    const content = findCell(sheet, { row: name.area.row, column: name.area.column + index })?.content;
    return content?.kind === 'text' ? content.value : '';
  });
  validateHeaders(labels, name.name);
  return labels;
}

function validateHeaders(labels: readonly string[], reference: string) {
  requireDocument(labels.length > 0 && labels.every(label => typeof label === 'string' && !!label.trim()) &&
    new Set(labels).size === labels.length, `Python input ${reference} needs nonblank unique text headers.`);
}

function dateTime(value: number, cell: CellDocument, sheet: SheetDocument): PythonInputValue {
  const parts = temporalParts(resolvedCellStyle(sheet, cell, cell)?.numberFormat);
  if (!parts.date && !parts.time) return value;
  requireDocument(Number.isFinite(value), 'Python date/time input is invalid.');
  if (parts.date) requireDocument(value >= 1 && value < 2_958_466 && Math.floor(value) !== 60,
    "Invalid Python date/time input (including Excel's fictitious 1900-02-29).");
  else requireDocument(value >= 0 && value < 1, 'Invalid Python time input.');
  const day = Math.floor(value);
  const fraction = value - day;
  let ticks = Math.round(fraction * 864_000_000_000);
  const nextDay = ticks >= 864_000_000_000;
  if (nextDay) ticks = 0;
  const hours = Math.floor(ticks / 36_000_000_000);
  const minutes = Math.floor(ticks / 600_000_000) % 60;
  const seconds = Math.floor(ticks / 10_000_000) % 60;
  const fractionText = String(ticks % 10_000_000).padStart(7, '0').replace(/0+$/u, '');
  const time = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}` +
    (fractionText ? `.${fractionText}` : '');
  if (!parts.date) return time;
  const dateSerial = day + Number(nextDay);
  requireDocument(dateSerial < 2_958_466 && dateSerial !== 60,
    "Invalid Python date/time input (including Excel's fictitious 1900-02-29).");
  const [year, month, date] = excelDateParts(dateSerial, 1900);
  const dateText = `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(date).padStart(2, '0')}`;
  return dateText + (parts.time ? `T${time}` : '');
}

function cacheValue(cache: CellCache | null | undefined, cell: CellDocument | undefined, sheet: SheetDocument,
  staleFormulaIds: ReadonlySet<string>, pythonBlocks: ReadonlyMap<string, PythonBlock>): PythonInputValue {
  if (!cell) return null;
  if (cell.requiresExternalCalculation) throw new Error('Python input contains an unverified formula; correct it before running.');
  if (cell.spillOwner) {
    const owner = pythonBlocks.get(cell.spillOwner);
    if (!owner?.hasResult || owner.status !== 'succeeded') throw new Error('Upstream Python spill is stale; run its anchor first.');
  }
  if (cell.content.kind === 'formula') {
    if (staleFormulaIds.has(cell.id)) throw new Error('Python input formula is stale; calculate native formulas first.');
    if (!cache || cache.kind === 'error') throw new Error('Python input formula has no current scalar value.');
  } else if (cell.content.kind === 'number') {
    cache = { kind: 'number', value: cell.content.value };
  } else if (cell.content.kind === 'text') {
    cache = { kind: 'text', value: cell.content.value };
  } else if (cell.content.kind === 'boolean') {
    cache = { kind: 'boolean', value: cell.content.value };
  } else if (cell.content.kind === 'python') {
    const block = cell.content.block;
    if (!block?.hasResult || block.status !== 'succeeded') throw new Error('Upstream Python result is stale; run its anchor first.');
    if (block.lastFigure) throw new Error('Figure output is not a scalar/table input; reference the chart data instead.');
    cache = block.lastValue;
  }
  if (!cache) return null;
  if (cache.kind === 'error') throw new Error('Python input contains a spreadsheet error; correct it before running.');
  if (cell.content.kind === 'number' || cache.kind === 'number') {
    const number = cache.kind === 'number' ? cache.value : undefined;
    if (number === undefined || !Number.isFinite(number)) throw new DocumentError('Python input number is invalid.');
    return dateTime(number, cell, sheet);
  }
  if (cache.kind === 'boolean') return cache.value;
  return cache.value;
}

function readRange(snapshot: WorkbookSnapshot, range: ResolvedRange, call: { reference: string; headers: boolean },
  staleFormulaIds: ReadonlySet<string>, pythonBlocks: ReadonlyMap<string, PythonBlock>, values: Record<string, PythonInputTable>, key: string): number {
  const sheet = snapshot.document.sheets.find(item => item.id === range.area.sheetId)!;
  const cells = new Map(sheet.cells.map(cell => [`${cell.row}:${cell.column}`, cell]));
  const rows: PythonInputValue[][] = [];
  let count = 0;
  for (let row = range.area.row; row < range.area.row + range.area.rows; row++) {
    const valuesRow: PythonInputValue[] = [];
    for (let column = range.area.column; column < range.area.column + range.area.columns; column++) {
      const cell = cells.get(`${row}:${column}`);
      const value = cacheValue(cell?.cache, cell, sheet, staleFormulaIds, pythonBlocks);
      valuesRow.push(value);
      count++;
    }
    rows.push(valuesRow);
  }
  requireDocument(rows.length > 0 && rows[0]!.length > 0, `Python input ${call.reference} is empty.`);
  const tableName = range.named?.isTable ? range.named : undefined;
  let headers: string[] | undefined;
  let dataRows = rows;
  if (call.headers) {
    if (tableName && range.dataOnlyTable) {
      headers = tableHeaders(snapshot.document, tableName);
    } else if (tableName && !tableDesign(tableName).headerRow) {
      headers = tableHeaders(snapshot.document, tableName);
    } else {
      const labelValues = rows[0]!;
      if (labelValues.some(value => typeof value !== 'string')) {
        throw new Error(`Python input ${call.reference} needs nonblank unique text headers.`);
      }
      headers = labelValues as string[];
      dataRows = rows.slice(1);
    }
    validateHeaders(headers, call.reference);
  }
  if (!call.headers && rows.length === 1 && rows[0]!.length === 1) {
    values[key] = { rows: [], scalarCell: true, scalar: rows[0]![0] ?? null };
  } else {
    values[key] = { rows: dataRows, ...(headers && { headers }) };
  }
  return count;
}

function hash(text: string): string {
  let value = 0xcbf29ce484222325n;
  for (const byte of new TextEncoder().encode(text)) {
    value ^= BigInt(byte);
    value = BigInt.asUintN(64, value * 0x100000001b3n);
  }
  return 'fnv1a64-' + value.toString(16).padStart(16, '0');
}

export function buildPythonInputs(snapshot: WorkbookSnapshot, anchor: PythonInputAnchor,
  staleFormulaIds: ReadonlySet<string> = new Set()): BuiltPythonInputs {
  const sheet = snapshot.document.sheets.find(item => item.id === anchor.sheetId);
  const cell = sheet && findCell(sheet, { row: anchor.row, column: anchor.column });
  requireDocument(snapshot.sessionId.length > 0 && cell?.content.kind === 'python' &&
    cell.content.blockId === anchor.blockId && cell.content.source === anchor.source, 'Python anchor changed; select it again.');
  requireDocument(new TextEncoder().encode(anchor.source).length <= 1024 * 1024, 'Source exceeds 1 MiB.');
  const references = discoverPythonReferences(anchor.source);
  requireDocument(cell.content.kind === 'python' && !(references.length > 0 &&
    cell.content.block?.bindings.some(binding => binding.name.startsWith('__xl_'))),
  'Legacy binding names beginning __xl_ are reserved for xl access.');
  const calls: PythonRuntimeInput['calls'][number][] = [];
  const bindings: PythonRuntimeInput['bindings'][number][] = [];
  const values: Record<string, PythonInputTable> = {};
  const pythonBlocks = new Map<string, PythonBlock>();
  for (const currentSheet of snapshot.document.sheets) for (const currentCell of currentSheet.cells) {
    if (currentCell.content.kind === 'python' && currentCell.content.block) {
      pythonBlocks.set(currentCell.content.blockId, currentCell.content.block);
    }
  }
  let total = 0;
  references.forEach((reference, index) => {
    const key = `__xl_${index}`;
    const resolved = resolveRange(snapshot.document, anchor.sheetId, reference.reference);
    requireDocument(total + resolved.area.rows * resolved.area.columns <= inputCellLimit,
      'Python inputs exceed 100,000 cells.');
    total += readRange(snapshot, resolved, reference, staleFormulaIds, pythonBlocks, values, key);
    calls.push({ reference: reference.reference, headers: reference.headers, key });
  });
  const block = cell.content.kind === 'python' ? cell.content.block : undefined;
  for (const [index, binding] of (block?.bindings ?? []).entries()) {
    requireDocument(binding.name !== 'xl' && !binding.name.startsWith('__xl_'), 'Legacy binding name is reserved for xl access.');
    const key = `__binding_${index}`;
    const resolved = resolveRange(snapshot.document, anchor.sheetId, binding.reference.replace(/^=/u, ''));
    requireDocument(total + resolved.area.rows * resolved.area.columns <= inputCellLimit,
      'Python inputs exceed 100,000 cells.');
    total += readRange(snapshot, resolved, { reference: binding.reference, headers: false }, staleFormulaIds, pythonBlocks, values, key);
    bindings.push({ name: binding.name, key });
  }
  requireDocument(total <= inputCellLimit, 'Python inputs exceed 100,000 cells.');
  const input: PythonRuntimeInput = { calls, bindings, values };
  const serialized = JSON.stringify({ source: anchor.source, input });
  requireDocument(new TextEncoder().encode(serialized).length <= frameLimit, 'Python inputs exceed 32 MiB.');
  return { input, fingerprint: hash(serialized) };
}
