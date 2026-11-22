import type { CellCache, CellContent, CellDocument, PythonBlock, PythonFigure, PythonTable } from '../core/workbook';
import { documentLimits, type CellReference, type WorkbookDocument } from '../core/workbook';
import { requireDocument } from '../core/workbook-validation';
import { pythonCache, type PythonInputValue, type PythonRuntimeOutput, type PythonRuntimeResult } from './runtime-protocol';

const maxResultBytes = 16 * 1024 * 1024;

function cacheFor(value: PythonInputValue): CellCache | null {
  requireDocument(value === null || typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number',
    'Python result contains an unsupported scalar value.');
  const cache = pythonCache(value);
  if (cache?.kind === 'number') requireDocument(Number.isFinite(cache.value), 'Python result contains a non-finite number.');
  if (cache?.kind === 'text') requireDocument(cache.value.length <= 32767, 'Python result text exceeds 32,767 characters.');
  return cache ?? null;
}

function resultTable(result: Extract<PythonRuntimeResult, { kind: 'table' }>): PythonTable {
  const headers = result.headers;
  const rows = result.rows;
  requireDocument(Array.isArray(rows) && (headers === null || Array.isArray(headers)), 'Invalid Python table result.');
  const columns = headers?.length ?? rows[0]?.length ?? 0;
  requireDocument(Number.isSafeInteger(columns) && columns >= 1 && columns <= documentLimits.columns && rows.length <= documentLimits.rows,
    'Python result table has an invalid shape.');
  if (headers) requireDocument(headers.length === columns && headers.every(value => typeof value === 'string' && !!value.trim()) &&
    new Set(headers).size === headers.length, 'Python result table needs nonblank unique text headers.');
  requireDocument(rows.every(row => Array.isArray(row) && row.length === columns), 'Python result table must be rectangular.');
  const rowCount = rows.length + Number(!!headers);
  requireDocument(rowCount >= 1 && rowCount <= documentLimits.rows && rowCount * columns <= 100_000,
    'Python result table exceeds worksheet/result limits.');
  const cells = [
    ...(headers ?? []).map(value => cacheFor(value)),
    ...rows.flatMap((row: readonly PythonInputValue[]) => row.map((value: PythonInputValue) => cacheFor(value))),
  ];
  const notices = result.notices;
  requireDocument(Array.isArray(notices) && notices.length <= 100 && notices.every(note => typeof note === 'string' && note.length <= 32767),
    'Python result notices exceed the workbook limits.');
  requireDocument(new TextEncoder().encode(JSON.stringify({ headers, rows, notices })).length <= maxResultBytes,
    'Python result table exceeds 16 MiB.');
  return { rows: rowCount, columns, cells, notices: [...notices] };
}

function figureResult(figure: PythonFigure): PythonFigure {
  requireDocument(Number.isSafeInteger(figure.width) && Number.isSafeInteger(figure.height) && figure.width > 0 &&
    figure.height > 0 && figure.width <= 8192 && figure.height <= 8192 && figure.width * figure.height <= 16_000_000,
  'Python figure dimensions are invalid.');
  requireDocument(typeof figure.pngBase64 === 'string' && figure.pngBase64.length <= Math.ceil(maxResultBytes * 4 / 3),
    'Python figure exceeds 16 MiB.');
  let bytes: Uint8Array;
  try { bytes = Uint8Array.from(atob(figure.pngBase64), character => character.charCodeAt(0)); }
  catch { throw new Error('Python figure is not valid base64.'); }
  requireDocument(bytes.length >= 24 && bytes.length <= maxResultBytes &&
    [137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value), 'Python figure is not a valid PNG.');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  requireDocument(view.getUint32(16) === figure.width && view.getUint32(20) === figure.height,
    'Python figure dimensions do not match its PNG.');
  return figure;
}

export function pythonBlockFromOutput(prior: PythonBlock, output: PythonRuntimeOutput,
  inputFingerprint: string, runtimeFingerprint: string): PythonBlock {
  requireDocument(inputFingerprint.length <= 128 && runtimeFingerprint.length > 0 && runtimeFingerprint.length <= 128,
    'Python runtime fingerprint is invalid.');
  requireDocument(output && (output.status === 'failed' || output.status === 'succeeded') &&
    typeof output.stdout === 'string' && new TextEncoder().encode(output.stdout).length <= 1024 * 1024 &&
    typeof output.stderr === 'string' && new TextEncoder().encode(output.stderr).length <= 1024 * 1024 &&
    typeof output.outputTruncated === 'boolean', 'Python output exceeded its log or protocol limits.');
  if (output.status === 'failed') {
    const error = typeof output.error === 'string' && output.error ? output.error.slice(-32767) : 'Python execution failed.';
    return { ...prior, status: 'failed', error, inputFingerprint, runtimeFingerprint };
  }
  requireDocument(output.status === 'succeeded' && output.result &&
    new TextEncoder().encode(JSON.stringify(output.result)).length <= 16 * 1024 * 1024,
  'Python worker returned an invalid or oversized result.');
  const result = output.result;
  if (result.kind === 'scalar') {
    return { ...prior, status: 'succeeded', error: null, hasResult: true, lastValue: cacheFor(result.value),
      inputFingerprint, runtimeFingerprint, lastTable: null, spillActive: false, lastFigure: null, placements: [] };
  }
  if (result.kind === 'table') {
    const lastTable = resultTable(result);
    return { ...prior, status: 'succeeded', error: null, hasResult: true, lastValue: lastTable.cells[0] ?? null,
      inputFingerprint, runtimeFingerprint, lastTable, spillActive: true, lastFigure: null, placements: [] };
  }
  if (result.kind === 'figure') {
    const lastFigure = figureResult(result.figure);
    return { ...prior, status: 'succeeded', error: null, hasResult: true, lastValue: null,
      inputFingerprint, runtimeFingerprint, lastTable: null, spillActive: false, lastFigure, placements: [] };
  }
  throw new Error('Python worker returned an unsupported result.');
}

export function pythonResultContent(value: CellCache | null | undefined): CellContent {
  if (!value) return { kind: 'blank' };
  switch (value.kind) {
    case 'number': return { kind: 'number', value: value.value };
    case 'text': return { kind: 'text', value: value.value };
    case 'boolean': return { kind: 'boolean', value: value.value };
    case 'error': return { kind: 'blank' };
  }
}

export function pythonInsertCommand(document: WorkbookDocument, target: CellReference, block: PythonBlock,
  makeId: () => string) {
  const values = block.spillActive && block.lastTable ? block.lastTable.cells : [block.lastValue ?? null];
  const rows = block.spillActive && block.lastTable ? block.lastTable.rows : 1;
  const columns = block.spillActive && block.lastTable ? block.lastTable.columns : 1;
  requireDocument(target.row + rows - 1 <= documentLimits.rows && target.column + columns - 1 <= documentLimits.columns,
    'Insert as Values exceeds worksheet limits.');
  const sheet = document.sheets.find(candidate => candidate.id === target.sheetId);
  requireDocument(sheet, 'Insert destination worksheet no longer exists.');
  const anchor = sheet.cells.find(cell => cell.row === target.row && cell.column === target.column);
  requireDocument(anchor?.content.kind === 'python' && anchor.content.blockId === block.id,
    'Python anchor changed before values were inserted.');
  const byPosition = new Map(sheet.cells.map(cell => [`${cell.row}:${cell.column}`, cell]));
  const replacements = new Map<string, CellDocument>();
  for (let index = 0; index < values.length; index++) {
    const row = target.row + Math.floor(index / columns);
    const column = target.column + index % columns;
    const existing = byPosition.get(`${row}:${column}`);
    const { spillOwner: _spillOwner, cache: _cache, requiresExternalCalculation: _external,
      content: _content, ...rest } = existing ?? { id: makeId(), row, column, content: { kind: 'blank' as const } };
    const id = existing?.id ?? (index === 0 ? anchor.id : rest.id);
    replacements.set(`${row}:${column}`, { ...rest, id, row, column, content: pythonResultContent(values[index]) });
  }
  const cells = sheet.cells.flatMap(cell => {
    if (cell.id === anchor.id || cell.spillOwner === block.id) return [];
    return [cell];
  });
  cells.push(...replacements.values());
  const updatedSheet = { ...sheet, cells: cells.sort((left, right) => left.row - right.row || left.column - right.column) };
  return { kind: 'structure' as const, label: 'Insert Python result as values',
    document: { ...document, sheets: document.sheets.map(current => current.id === sheet.id ? updatedSheet : current) } };
}

export function pythonResultText(block: PythonBlock): string | undefined {
  if (block.spillActive && block.lastTable) {
    return Array.from({ length: block.lastTable.rows }, (_, row) =>
      Array.from({ length: block.lastTable!.columns }, (_, column) => {
        const value = block.lastTable!.cells[row * block.lastTable!.columns + column];
        if (!value) return '';
        if (value.kind === 'boolean') return value.value ? 'TRUE' : 'FALSE';
        if (value.kind === 'error') return value.value;
        return String(value.value);
      }).join('\t')).join('\n');
  }
  if (block.hasResult && block.lastValue !== undefined) {
    if (block.lastValue === null) return '';
    if (block.lastValue.kind === 'boolean') return block.lastValue.value ? 'TRUE' : 'FALSE';
    if (block.lastValue.kind === 'error') return block.lastValue.value;
    return String(block.lastValue.value);
  }
  return undefined;
}
