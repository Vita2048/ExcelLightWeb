import { BlobReader, BlobWriter, TextReader, TextWriter, ZipReader, ZipWriter } from '@zip.js/zip.js';
import type {
  AxisDimension, CellBorder, CellCache, CellDocument, CellExtent, CellFill, CellStyle,
  Coordinate, DocumentId, FigurePlacement, FilterCriterion, FrozenPanes,
  PythonBinding, PythonBlock, PythonFigure, PythonTable, SheetDocument, SheetFilter,
  TableDesign, WorkbookDocument, WorkbookName,
} from '../../core/workbook';
import type {
  PivotAggregation, PivotAxisField, PivotFilterField, PivotReportLayout, PivotShowValuesAs,
  PivotSortRule, PivotStyleId, PivotSubtotalPosition, PivotTableDocument, PivotValueField,
} from '../../core/pivot';
import { DocumentError, documentLimits, formulaErrors, ownDto } from '../../core/workbook';
import { equalDto } from '../../core/workbook-delta';
import { migrateDetachedHeaders } from '../../core/workbook-names';
import { validateWorkbook } from '../../core/workbook-validation';
import { parseStrictJson } from './json-syntax';
import { PackageException } from './package-exception';
import { defaultPackageLimits, SaveFault, type PackageLimits, type SaveRequest } from './package-limits';
import { validateNativeFigurePayloads } from './png-validation';

const guidRegex = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
const sheetEntryRegex = /^sheets\/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\.json$/;

function requireGuid(val: unknown, name: string): DocumentId {
  if (typeof val !== 'string' || !guidRegex.test(val) || val === '00000000-0000-0000-0000-000000000000') {
    throw new PackageException(`${name} must be a GUID.`);
  }
  return val.toLowerCase();
}

function requireObject(val: unknown, name: string): Record<string, unknown> {
  if (val === null || typeof val !== 'object' || Array.isArray(val)) {
    throw new PackageException(`${name} must be a JSON object.`);
  }
  return val as Record<string, unknown>;
}

function rejectUnknown(obj: Record<string, unknown>, ...allowed: string[]): Record<string, unknown> {
  const set = new Set(allowed);
  for (const key of Object.keys(obj)) {
    if (!set.has(key)) {
      throw new PackageException(`Unexpected property ${key}.`);
    }
  }
  return obj;
}

function requireProperty(obj: Record<string, unknown>, name: string): unknown {
  if (!Object.hasOwn(obj, name)) {
    throw new PackageException(`Missing property ${name}.`);
  }
  return obj[name];
}

function requireString(val: unknown, name: string): string {
  if (typeof val !== 'string') {
    throw new PackageException(`${name} must be a string.`);
  }
  return val;
}

function requireInt(val: unknown, name: string): number {
  if (typeof val !== 'number' || !Number.isInteger(val)) {
    throw new PackageException(`${name} must be an integer.`);
  }
  return val;
}

function requireBool(val: unknown, name: string): boolean {
  if (typeof val !== 'boolean') {
    throw new PackageException(`${name} must be a boolean.`);
  }
  return val;
}

function requireFinite(val: unknown, name: string): number {
  if (typeof val !== 'number' || !Number.isFinite(val)) {
    throw new PackageException(`${name} must be a finite number.`);
  }
  return val;
}

function readEntryText(bytes: Uint8Array): string {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    throw new PackageException('Malformed package JSON.');
  }
}

function firstDtoDifference(left: unknown, right: unknown, path = 'document'): string | undefined {
  if (Object.is(left, right)) return undefined;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object' ||
    Array.isArray(left) !== Array.isArray(right)) return path;
  const leftKeys = Object.keys(left);
  const rightKeys = Object.keys(right);
  if (leftKeys.length !== rightKeys.length) {
    const missingRight = leftKeys.find(key => !Object.hasOwn(right, key));
    if (missingRight) return `${path}.${missingRight} (missing after load)`;
    const missingLeft = rightKeys.find(key => !Object.hasOwn(left, key));
    if (missingLeft) return `${path}.${missingLeft} (added after load)`;
    return path;
  }
  for (const key of leftKeys) {
    if (!Object.hasOwn(right, key)) return `${path}.${key}`;
    const child = firstDtoDifference(
      (left as Record<string, unknown>)[key],
      (right as Record<string, unknown>)[key],
      Array.isArray(left) ? `${path}[${key}]` : `${path}.${key}`,
    );
    if (child) return child;
  }
  return undefined;
}

export async function loadNativePackage(blob: Blob, limits?: PackageLimits): Promise<WorkbookDocument> {
  const cfg = { ...defaultPackageLimits, ...limits };
  if (blob.size > cfg.maxPackageBytes) {
    throw new PackageException(`Package exceeds the ${cfg.maxPackageBytes} byte limit.`);
  }

  let manifestBytes: Uint8Array | undefined;
  let workbookBytes: Uint8Array | undefined;
  const sheetBytesMap = new Map<string, Uint8Array>();

  let reader: ZipReader<Blob>;
  try {
    reader = new ZipReader(new BlobReader(blob), {
      strictness: 'strict',
      checkCrc32: true,
      checkOverlappingEntry: true,
      useWebWorkers: false,
    });
  } catch {
    throw new PackageException('Package is not a valid .exlt archive.');
  }

  const seen = new Set<string>();
  let entryCount = 0;
  let expandedBytes = 0;

  try {
    for await (const entry of reader.getEntriesGenerator()) {
      entryCount++;
      if (entryCount > cfg.maxEntries) {
        throw new PackageException(`Package has more than ${cfg.maxEntries} entries.`);
      }

      const rawName = entry.filename;
      const name = rawName.replaceAll('\\', '/');

      if (name.includes('..') || name.startsWith('/') || name.includes(':') || name.endsWith('/') || entry.directory) {
        throw new PackageException(`Package entry path is not allowed: ${entry.filename}`);
      }

      const lowerName = name.toLowerCase();
      if (seen.has(name)) {
        throw new PackageException(`Duplicate package entry '${name}'.`);
      }
      seen.add(name);

      if (lowerName.startsWith('python/') || lowerName.startsWith('results/') || lowerName.startsWith('assets/')) {
        throw new PackageException('Executable metadata is not accepted in package version 1.');
      }

      if (name !== 'manifest.json' && name !== 'workbook.json' && !sheetEntryRegex.test(name)) {
        throw new PackageException(`Unexpected package entry '${name}'.`);
      }

      if (entry.uncompressedSize > cfg.maxEntryBytes) {
        throw new PackageException(`Package entry ${entry.filename} exceeds the entry limit.`);
      }

      let entryRead = 0;
      const chunks: Uint8Array[] = [];

      await entry.getData(
        new WritableStream<Uint8Array>({
          write(chunk) {
            entryRead += chunk.byteLength;
            expandedBytes += chunk.byteLength;
            if (entryRead > cfg.maxEntryBytes || expandedBytes > cfg.maxExpandedBytes) {
              throw new PackageException(`Package entry ${entry.filename} exceeds the size limit.`);
            }
            chunks.push(chunk);
          },
        }),
      );

      const merged = new Uint8Array(entryRead);
      let offset = 0;
      for (const ch of chunks) {
        merged.set(ch, offset);
        offset += ch.byteLength;
      }

      if (name === 'manifest.json') {
        manifestBytes = merged;
      } else if (name === 'workbook.json') {
        workbookBytes = merged;
      } else {
        sheetBytesMap.set(lowerName, merged);
      }
    }
  } catch (err: unknown) {
    if (err instanceof PackageException) throw err;
    if (err && typeof err === 'object' && 'filename' in err && typeof (err as { filename: unknown }).filename === 'string') {
      throw new PackageException(`Package entry path is not allowed: ${(err as { filename: string }).filename}`);
    }
    if (err instanceof Error && err.message === 'Unsafe filename') {
      throw new PackageException('Package entry path is not allowed.');
    }
    throw new PackageException('Package is not a valid .exlt archive.');
  } finally {
    await reader.close().catch(() => {});
  }

  if (!manifestBytes || !workbookBytes) {
    throw new PackageException('Package manifest or workbook entry is missing.');
  }

  const document = readDocument(manifestBytes, workbookBytes, sheetBytesMap);
  try {
    validateWorkbook(document);
    await validateNativeFigurePayloads(document);
  } catch (ex: unknown) {
    if (ex instanceof DocumentError) {
      throw new PackageException(ex.message);
    }
    throw ex;
  }
  return document;
}

function readDocument(manifestBytes: Uint8Array, workbookBytes: Uint8Array, sheetBytes: Map<string, Uint8Array>): WorkbookDocument {
  const manifest = parseStrictJson(readEntryText(manifestBytes));
  const workbook = parseStrictJson(readEntryText(workbookBytes));

  const manifestRoot = requireObject(manifest, 'manifest.json');
  rejectUnknown(manifestRoot, 'formatVersion', 'workbookId', 'dateSystem');

  const version = requireProperty(manifestRoot, 'formatVersion');
  if (typeof version !== 'number' || !Number.isInteger(version)) {
    throw new PackageException('formatVersion must be an integer.');
  }
  if (version !== 1) {
    throw new PackageException(`Unsupported package version ${version}. This build reads version 1.`);
  }

  const workbookId = requireGuid(requireProperty(manifestRoot, 'workbookId'), 'workbookId');
  const dateSystem = requireString(requireProperty(manifestRoot, 'dateSystem'), 'dateSystem');
  if (dateSystem !== '1900') {
    throw new PackageException(`Date system '${dateSystem}' is not supported. This prototype does not convert other epochs.`);
  }

  const workbookRoot = requireObject(workbook, 'workbook.json');
  rejectUnknown(workbookRoot, 'workbookId', 'sheets', 'names', 'pivots');

  if (requireGuid(requireProperty(workbookRoot, 'workbookId'), 'workbookId') !== workbookId) {
    throw new PackageException('Workbook id does not match the manifest.');
  }

  const sheetList = requireProperty(workbookRoot, 'sheets');
  if (!Array.isArray(sheetList)) {
    throw new PackageException('sheets must be an array.');
  }

  const order: { id: DocumentId; name: string }[] = [];
  for (const item of sheetList) {
    const sheetObj = requireObject(item, 'sheet');
    rejectUnknown(sheetObj, 'id', 'name');
    order.push({
      id: requireGuid(requireProperty(sheetObj, 'id'), 'id'),
      name: requireString(requireProperty(sheetObj, 'name'), 'name'),
    });
  }

  if (order.length !== sheetBytes.size) {
    throw new PackageException('Sheet files do not match the workbook sheet list.');
  }

  const sheets: SheetDocument[] = [];
  for (const { id, name } of order) {
    const entryPath = `sheets/${id}.json`.toLowerCase();
    const bytes = sheetBytes.get(entryPath);
    if (!bytes) {
      throw new PackageException(`Sheet file sheets/${id}.json is missing.`);
    }
    sheets.push(readSheet(bytes, id, name));
  }

  let names: WorkbookName[] = [];
  if (Object.hasOwn(workbookRoot, 'names')) {
    const namesVal = workbookRoot.names;
    if (!Array.isArray(namesVal)) {
      throw new PackageException('Invalid workbook names: names must be an array.');
    }
    try {
      names = namesVal.map(readWorkbookName);
    } catch (ex) {
      if (ex instanceof PackageException) throw ex;
      throw new PackageException('Invalid workbook names: ' + (ex instanceof Error ? ex.message : String(ex)));
    }
  }

  let pivots: PivotTableDocument[] | undefined;
  if (Object.hasOwn(workbookRoot, 'pivots')) {
    const pivotsVal = workbookRoot.pivots;
    if (!Array.isArray(pivotsVal)) {
      throw new PackageException('Invalid workbook pivots: pivots must be an array.');
    }
    try {
      pivots = pivotsVal.map(readPivotTable);
    } catch (ex) {
      if (ex instanceof PackageException) throw ex;
      throw new PackageException('Invalid workbook pivots: ' + (ex instanceof Error ? ex.message : String(ex)));
    }
  }

  const doc: WorkbookDocument = {
    id: workbookId,
    formatVersion: 1,
    dateSystem: '1900',
    sheets,
    names,
    ...(pivots && pivots.length > 0 ? { pivots } : {}),
  };

  return migrateDetachedHeaders(doc);
}

function readWorkbookName(raw: unknown): WorkbookName {
  const obj = requireObject(raw, 'name');
  rejectUnknown(obj, 'Id', 'Name', 'Area', 'Comment', 'IsTable', 'Design', 'HeaderLabels');

  const id = requireGuid(requireProperty(obj, 'Id'), 'Id');
  const name = requireString(requireProperty(obj, 'Name'), 'Name');
  const areaObj = requireObject(requireProperty(obj, 'Area'), 'Area');
  rejectUnknown(areaObj, 'SheetId', 'Row', 'Column', 'Rows', 'Columns', 'IsInvalid');

  const sheetId = requireGuid(requireProperty(areaObj, 'SheetId'), 'SheetId');
  const row = requireInt(requireProperty(areaObj, 'Row'), 'Row');
  const column = requireInt(requireProperty(areaObj, 'Column'), 'Column');
  const rows = requireInt(requireProperty(areaObj, 'Rows'), 'Rows');
  const columns = requireInt(requireProperty(areaObj, 'Columns'), 'Columns');
  const isInvalid = Object.hasOwn(areaObj, 'IsInvalid') ? requireBool(areaObj.IsInvalid, 'IsInvalid') : false;

  let design: TableDesign | undefined;
  if (Object.hasOwn(obj, 'Design') && obj.Design !== null) {
    const dObj = requireObject(obj.Design, 'Design');
    rejectUnknown(dObj, 'Style', 'HeaderRow', 'TotalRow', 'FirstColumn', 'LastColumn', 'BandedRows', 'BandedColumns', 'FilterButton');
    const styleStr = requireString(requireProperty(dObj, 'Style'), 'Style');
    if (!['Neutral', 'Blue', 'Green', 'Orange', 'Purple'].includes(styleStr)) {
      throw new PackageException('Invalid table design style.');
    }
    design = {
      style: styleStr as TableDesign['style'],
      headerRow: requireBool(requireProperty(dObj, 'HeaderRow'), 'HeaderRow'),
      totalRow: requireBool(requireProperty(dObj, 'TotalRow'), 'TotalRow'),
      firstColumn: Object.hasOwn(dObj, 'FirstColumn') ? requireBool(dObj.FirstColumn, 'FirstColumn') : false,
      lastColumn: Object.hasOwn(dObj, 'LastColumn') ? requireBool(dObj.LastColumn, 'LastColumn') : false,
      bandedRows: Object.hasOwn(dObj, 'BandedRows') ? requireBool(dObj.BandedRows, 'BandedRows') : true,
      bandedColumns: Object.hasOwn(dObj, 'BandedColumns') ? requireBool(dObj.BandedColumns, 'BandedColumns') : false,
      filterButton: Object.hasOwn(dObj, 'FilterButton') ? requireBool(dObj.FilterButton, 'FilterButton') : true,
    };
  }

  let headerLabels: string[] | undefined;
  if (Object.hasOwn(obj, 'HeaderLabels') && obj.HeaderLabels !== null) {
    if (!Array.isArray(obj.HeaderLabels)) {
      throw new PackageException('HeaderLabels must be an array.');
    }
    headerLabels = obj.HeaderLabels.map((lbl, idx) => requireString(lbl, `HeaderLabels[${idx}]`));
  }

  const comment = Object.hasOwn(obj, 'Comment') && obj.Comment !== null ? requireString(obj.Comment, 'Comment') : undefined;

  return {
    id,
    name,
    area: { sheetId, row, column, rows, columns, ...(isInvalid ? { invalid: true } : {}) },
    ...(comment !== undefined ? { comment } : {}),
    isTable: Object.hasOwn(obj, 'IsTable') ? requireBool(obj.IsTable, 'IsTable') : false,
    ...(design ? { design } : {}),
    ...(headerLabels ? { headerLabels } : {}),
  };
}

function readPivotStringList(raw: unknown, prop: string): readonly string[] {
  if (!Array.isArray(raw)) throw new PackageException(`${prop} must be an array.`);
  return raw.map((item, idx) => requireString(item, `${prop}[${idx}]`));
}

function readPivotAxisField(raw: unknown): PivotAxisField {
  const obj = requireObject(raw, 'axisField');
  rejectUnknown(obj, 'fieldName', 'subtotal', 'sort', 'hiddenItems', 'numberFormat');
  const fieldName = requireString(requireProperty(obj, 'fieldName'), 'fieldName');
  let subtotal: PivotAxisField['subtotal'];
  if (Object.hasOwn(obj, 'subtotal')) {
    const val = requireString(obj.subtotal, 'subtotal');
    if (val !== 'automatic' && val !== 'none') throw new PackageException('Invalid axis subtotal.');
    subtotal = val;
  }
  let sort: PivotSortRule | undefined;
  if (Object.hasOwn(obj, 'sort')) {
    const sObj = requireObject(obj.sort, 'sort');
    rejectUnknown(sObj, 'direction', 'byValueFieldId');
    const dir = requireString(requireProperty(sObj, 'direction'), 'direction');
    if (dir !== 'asc' && dir !== 'desc') throw new PackageException('Invalid pivot sort direction.');
    sort = {
      direction: dir,
      ...(Object.hasOwn(sObj, 'byValueFieldId') ? { byValueFieldId: requireGuid(sObj.byValueFieldId, 'byValueFieldId') } : {}),
    };
  }
  return {
    fieldName,
    ...(subtotal !== undefined ? { subtotal } : {}),
    ...(sort !== undefined ? { sort } : {}),
    ...(Object.hasOwn(obj, 'hiddenItems') ? { hiddenItems: readPivotStringList(obj.hiddenItems, 'hiddenItems') } : {}),
    ...(Object.hasOwn(obj, 'numberFormat') ? { numberFormat: requireString(obj.numberFormat, 'numberFormat') } : {}),
  };
}

function readPivotValueField(raw: unknown): PivotValueField {
  const obj = requireObject(raw, 'valueField');
  rejectUnknown(obj, 'id', 'fieldName', 'customName', 'aggregation', 'showValuesAs', 'baseField', 'baseItem', 'numberFormat');
  return {
    id: requireGuid(requireProperty(obj, 'id'), 'id'),
    fieldName: requireString(requireProperty(obj, 'fieldName'), 'fieldName'),
    ...(Object.hasOwn(obj, 'customName') ? { customName: requireString(obj.customName, 'customName') } : {}),
    aggregation: requireString(requireProperty(obj, 'aggregation'), 'aggregation') as PivotAggregation,
    ...(Object.hasOwn(obj, 'showValuesAs') ? { showValuesAs: requireString(obj.showValuesAs, 'showValuesAs') as PivotShowValuesAs } : {}),
    ...(Object.hasOwn(obj, 'baseField') ? { baseField: requireString(obj.baseField, 'baseField') } : {}),
    ...(Object.hasOwn(obj, 'baseItem') ? { baseItem: requireString(obj.baseItem, 'baseItem') } : {}),
    ...(Object.hasOwn(obj, 'numberFormat') ? { numberFormat: requireString(obj.numberFormat, 'numberFormat') } : {}),
  };
}

function readPivotFilterField(raw: unknown): PivotFilterField {
  const obj = requireObject(raw, 'filterField');
  rejectUnknown(obj, 'fieldName', 'hiddenItems', 'multiSelect');
  return {
    fieldName: requireString(requireProperty(obj, 'fieldName'), 'fieldName'),
    ...(Object.hasOwn(obj, 'hiddenItems') ? { hiddenItems: readPivotStringList(obj.hiddenItems, 'hiddenItems') } : {}),
    ...(Object.hasOwn(obj, 'multiSelect') ? { multiSelect: requireBool(obj.multiSelect, 'multiSelect') } : {}),
  };
}

function readPivotExtent(raw: unknown, prop: string): CellExtent {
  const obj = requireObject(raw, prop);
  rejectUnknown(obj, 'row', 'column', 'rows', 'columns');
  return {
    row: requireInt(requireProperty(obj, 'row'), `${prop}.row`),
    column: requireInt(requireProperty(obj, 'column'), `${prop}.column`),
    rows: requireInt(requireProperty(obj, 'rows'), `${prop}.rows`),
    columns: requireInt(requireProperty(obj, 'columns'), `${prop}.columns`),
  };
}

function readPivotTable(raw: unknown): PivotTableDocument {
  const obj = requireObject(raw, 'pivot');
  rejectUnknown(
    obj,
    'id', 'name', 'source', 'target', 'extent', 'filters', 'columns', 'rows', 'values',
    'valuesInRows', 'reportLayout', 'subtotals', 'grandTotals', 'repeatLabels', 'blankRows',
    'showButtons', 'showFieldHeaders', 'emptyCellText', 'autoFitColumns', 'style', 'styleOptions', 'collapsedKeys',
  );
  const srcObj = requireObject(requireProperty(obj, 'source'), 'source');
  const srcKind = requireString(requireProperty(srcObj, 'kind'), 'source.kind');
  let source: PivotTableDocument['source'];
  if (srcKind === 'range') {
    rejectUnknown(srcObj, 'kind', 'area');
    const aObj = requireObject(requireProperty(srcObj, 'area'), 'source.area');
    rejectUnknown(aObj, 'sheetId', 'row', 'column', 'rows', 'columns', 'invalid');
    source = {
      kind: 'range',
      area: {
        sheetId: requireGuid(requireProperty(aObj, 'sheetId'), 'source.area.sheetId'),
        row: requireInt(requireProperty(aObj, 'row'), 'source.area.row'),
        column: requireInt(requireProperty(aObj, 'column'), 'source.area.column'),
        rows: requireInt(requireProperty(aObj, 'rows'), 'source.area.rows'),
        columns: requireInt(requireProperty(aObj, 'columns'), 'source.area.columns'),
        ...(Object.hasOwn(aObj, 'invalid') ? { invalid: requireBool(aObj.invalid, 'source.area.invalid') } : {}),
      },
    };
  } else if (srcKind === 'table') {
    rejectUnknown(srcObj, 'kind', 'tableId');
    source = { kind: 'table', tableId: requireGuid(requireProperty(srcObj, 'tableId'), 'source.tableId') };
  } else {
    throw new PackageException(`Invalid pivot source kind '${srcKind}'.`);
  }

  const targetObj = requireObject(requireProperty(obj, 'target'), 'target');
  rejectUnknown(targetObj, 'sheetId', 'row', 'column');
  const gtObj = requireObject(requireProperty(obj, 'grandTotals'), 'grandTotals');
  rejectUnknown(gtObj, 'rows', 'columns');
  const soObj = requireObject(requireProperty(obj, 'styleOptions'), 'styleOptions');
  rejectUnknown(soObj, 'rowHeaders', 'columnHeaders', 'bandedRows', 'bandedColumns');

  const filtersRaw = requireProperty(obj, 'filters');
  const colsRaw = requireProperty(obj, 'columns');
  const rowsRaw = requireProperty(obj, 'rows');
  const valuesRaw = requireProperty(obj, 'values');
  if (!Array.isArray(filtersRaw) || !Array.isArray(colsRaw) || !Array.isArray(rowsRaw) || !Array.isArray(valuesRaw)) {
    throw new PackageException('PivotTable filters, columns, rows, and values must be arrays.');
  }

  return {
    id: requireGuid(requireProperty(obj, 'id'), 'id'),
    name: requireString(requireProperty(obj, 'name'), 'name'),
    source,
    target: {
      sheetId: requireGuid(requireProperty(targetObj, 'sheetId'), 'target.sheetId'),
      row: requireInt(requireProperty(targetObj, 'row'), 'target.row'),
      column: requireInt(requireProperty(targetObj, 'column'), 'target.column'),
    },
    ...(Object.hasOwn(obj, 'extent') ? { extent: readPivotExtent(obj.extent, 'extent') } : {}),
    filters: filtersRaw.map(readPivotFilterField),
    columns: colsRaw.map(readPivotAxisField),
    rows: rowsRaw.map(readPivotAxisField),
    values: valuesRaw.map(readPivotValueField),
    ...(Object.hasOwn(obj, 'valuesInRows') ? { valuesInRows: requireBool(obj.valuesInRows, 'valuesInRows') } : {}),
    reportLayout: requireString(requireProperty(obj, 'reportLayout'), 'reportLayout') as PivotReportLayout,
    subtotals: requireString(requireProperty(obj, 'subtotals'), 'subtotals') as PivotSubtotalPosition,
    grandTotals: {
      rows: requireBool(requireProperty(gtObj, 'rows'), 'grandTotals.rows'),
      columns: requireBool(requireProperty(gtObj, 'columns'), 'grandTotals.columns'),
    },
    ...(Object.hasOwn(obj, 'repeatLabels') ? { repeatLabels: requireBool(obj.repeatLabels, 'repeatLabels') } : {}),
    ...(Object.hasOwn(obj, 'blankRows') ? { blankRows: requireBool(obj.blankRows, 'blankRows') } : {}),
    ...(Object.hasOwn(obj, 'showButtons') ? { showButtons: requireBool(obj.showButtons, 'showButtons') } : {}),
    ...(Object.hasOwn(obj, 'showFieldHeaders') ? { showFieldHeaders: requireBool(obj.showFieldHeaders, 'showFieldHeaders') } : {}),
    ...(Object.hasOwn(obj, 'emptyCellText') ? { emptyCellText: requireString(obj.emptyCellText, 'emptyCellText') } : {}),
    ...(Object.hasOwn(obj, 'autoFitColumns') ? { autoFitColumns: requireBool(obj.autoFitColumns, 'autoFitColumns') } : {}),
    style: requireString(requireProperty(obj, 'style'), 'style') as PivotStyleId,
    styleOptions: {
      rowHeaders: requireBool(requireProperty(soObj, 'rowHeaders'), 'styleOptions.rowHeaders'),
      columnHeaders: requireBool(requireProperty(soObj, 'columnHeaders'), 'styleOptions.columnHeaders'),
      bandedRows: requireBool(requireProperty(soObj, 'bandedRows'), 'styleOptions.bandedRows'),
      bandedColumns: requireBool(requireProperty(soObj, 'bandedColumns'), 'styleOptions.bandedColumns'),
    },
    ...(Object.hasOwn(obj, 'collapsedKeys') ? { collapsedKeys: readPivotStringList(obj.collapsedKeys, 'collapsedKeys') } : {}),
  };
}

function readSheet(bytes: Uint8Array, expectedId: DocumentId, expectedName: string): SheetDocument {
  const rootObj = requireObject(parseStrictJson(readEntryText(bytes)), expectedId);
  rejectUnknown(rootObj, 'sheetId', 'name', 'cells', 'rows', 'columns', 'filter', 'freeze', 'merges');

  const id = requireGuid(requireProperty(rootObj, 'sheetId'), 'sheetId');
  if (id !== expectedId) {
    throw new PackageException(`Sheet file does not match id ${expectedId}.`);
  }
  const name = requireString(requireProperty(rootObj, 'name'), 'name');
  if (name !== expectedName) {
    throw new PackageException(`Sheet name '${name}' does not match the workbook list.`);
  }

  const cellsElement = requireProperty(rootObj, 'cells');
  if (!Array.isArray(cellsElement)) {
    throw new PackageException('cells must be an array.');
  }

  const rows = readDimensions(rootObj, 'rows');
  const columns = readDimensions(rootObj, 'columns');
  const styledRows = new Set(rows.filter(r => r.style !== undefined).map(r => r.index));
  const styledCols = new Set(columns.filter(c => c.style !== undefined).map(c => c.index));
  const cells = cellsElement.map(raw => readCell(raw, styledRows, styledCols));
  const merges = readMerges(rootObj);
  const frozenPanes = readFreeze(rootObj);
  const filter = readFilter(rootObj);

  return {
    id,
    name,
    cells,
    rows,
    columns,
    merges,
    frozenPanes,
    ...(filter !== undefined ? { filter } : {}),
  };
}

function readMerges(root: Record<string, unknown>): readonly CellExtent[] {
  if (!Object.hasOwn(root, 'merges')) return [];
  const val = root.merges;
  if (!Array.isArray(val)) {
    throw new PackageException('merges must be an array.');
  }
  return val.map(item => {
    const obj = requireObject(item, 'merge');
    rejectUnknown(obj, 'Row', 'Column', 'Rows', 'Columns');
    return {
      row: requireInt(requireProperty(obj, 'Row'), 'Row'),
      column: requireInt(requireProperty(obj, 'Column'), 'Column'),
      rows: requireInt(requireProperty(obj, 'Rows'), 'Rows'),
      columns: requireInt(requireProperty(obj, 'Columns'), 'Columns'),
    };
  });
}

function readDimensions(root: Record<string, unknown>, key: string): readonly AxisDimension[] {
  if (!Object.hasOwn(root, key)) return [];
  const val = root[key];
  if (!Array.isArray(val)) {
    throw new PackageException('Dimensions must be an array.');
  }
  return val.map(item => {
    const obj = requireObject(item, key);
    rejectUnknown(obj, 'index', 'size', 'hidden', 'style');
    const size = Object.hasOwn(obj, 'size') ? requireFinite(obj.size, 'size') : undefined;
    const style = Object.hasOwn(obj, 'style') && obj.style !== null ? readStyle(requireObject(obj.style, 'style')) : undefined;
    return {
      index: requireInt(requireProperty(obj, 'index'), 'index'),
      ...(size !== undefined ? { size } : {}),
      hidden: Object.hasOwn(obj, 'hidden') ? requireBool(obj.hidden, 'hidden') : false,
      ...(style !== undefined ? { style } : {}),
    };
  });
}

function readFreeze(root: Record<string, unknown>): FrozenPanes {
  if (!Object.hasOwn(root, 'freeze')) return { rows: 0, columns: 0 };
  const val = root.freeze;
  if (typeof val === 'string') {
    if (val === 'topRow') return { rows: 1, columns: 0 };
    if (val === 'firstColumn') return { rows: 0, columns: 1 };
    throw new PackageException(`Freeze setting '${val}' is not supported.`);
  }
  const obj = requireObject(val, 'freeze');
  rejectUnknown(obj, 'rows', 'columns');
  const rows = requireInt(requireProperty(obj, 'rows'), 'rows');
  const columns = requireInt(requireProperty(obj, 'columns'), 'columns');
  if (rows < 0 || rows >= documentLimits.rows || columns < 0 || columns >= documentLimits.columns || (rows === 0 && columns === 0)) {
    throw new PackageException('Frozen pane boundaries are invalid.');
  }
  return { rows, columns };
}

function readFilter(root: Record<string, unknown>): SheetFilter | undefined {
  if (!Object.hasOwn(root, 'filter')) return undefined;
  const val = root.filter;
  if (val === null || val === undefined) return undefined;
  const filter = requireObject(val, 'filter');
  rejectUnknown(filter, 'row', 'column', 'rows', 'columns', 'criteria', 'sort');

  const row = requireInt(requireProperty(filter, 'row'), 'row');
  const column = requireInt(requireProperty(filter, 'column'), 'column');
  const rows = requireInt(requireProperty(filter, 'rows'), 'rows');
  const columns = requireInt(requireProperty(filter, 'columns'), 'columns');

  let sort: { column: number; ascending: boolean } | undefined;
  if (Object.hasOwn(filter, 'sort') && filter.sort !== null) {
    const order = requireObject(filter.sort, 'sort');
    rejectUnknown(order, 'column', 'ascending');
    const ascProp = requireProperty(order, 'ascending');
    if (typeof ascProp !== 'boolean') {
      throw new PackageException('Sort direction must be a boolean.');
    }
    sort = {
      column: requireInt(requireProperty(order, 'column'), 'column'),
      ascending: ascProp,
    };
  }

  const criteriaElement = requireProperty(filter, 'criteria');
  if (!Array.isArray(criteriaElement)) {
    throw new PackageException('Filter criteria must be an array.');
  }

  const criteria: FilterCriterion[] = [];
  for (const item of criteriaElement) {
    const criterion = requireObject(item, 'criterion');
    rejectUnknown(criterion, 'column', 'allowed');
    const col = requireInt(requireProperty(criterion, 'column'), 'column');
    const allowedElement = requireProperty(criterion, 'allowed');
    if (!Array.isArray(allowedElement)) {
      throw new PackageException('Filter values must be an array.');
    }
    const allowed: string[] = [];
    for (const entry of allowedElement) {
      allowed.push(requireString(entry, 'allowed'));
    }
    criteria.push({ column: col, allowed });
  }

  return { row, column, rows, columns, criteria, ...(sort ? { sort } : {}) };
}

function readCell(item: unknown, styledRows?: ReadonlySet<number>, styledCols?: ReadonlySet<number>): CellDocument {
  const cell = requireObject(item, 'cell');
  rejectUnknown(cell, 'id', 'row', 'column', 'content', 'style', 'cache', 'requiresExternalCalculation', 'spillOwner', 'spillKind', 'spillExtent');

  const id = requireGuid(requireProperty(cell, 'id'), 'id');
  const row = requireInt(requireProperty(cell, 'row'), 'row');
  const column = requireInt(requireProperty(cell, 'column'), 'column');

  const contentElement = requireObject(requireProperty(cell, 'content'), 'content');
  const kind = requireString(requireProperty(contentElement, 'kind'), 'kind');

  let content: CellDocument['content'];
  switch (kind) {
    case 'python': {
      rejectUnknown(contentElement, 'kind', 'block');
      const block = readPython(requireProperty(contentElement, 'block'));
      content = { kind: 'python', blockId: block.id, source: block.source, block };
      break;
    }
    case 'blank': {
      rejectUnknown(contentElement, 'kind');
      content = { kind: 'blank' };
      break;
    }
    case 'number': {
      rejectUnknown(contentElement, 'kind', 'value');
      content = { kind: 'number', value: requireFinite(requireProperty(contentElement, 'value'), 'value') };
      break;
    }
    case 'text': {
      rejectUnknown(contentElement, 'kind', 'value');
      content = { kind: 'text', value: requireString(requireProperty(contentElement, 'value'), 'value') };
      break;
    }
    case 'boolean': {
      rejectUnknown(contentElement, 'kind', 'value');
      content = { kind: 'boolean', value: requireBool(requireProperty(contentElement, 'value'), 'value') };
      break;
    }
    case 'formula': {
      rejectUnknown(contentElement, 'kind', 'source');
      content = { kind: 'formula', source: requireString(requireProperty(contentElement, 'source'), 'source') };
      break;
    }
    default:
      throw new PackageException(`Unexpected content kind '${kind}'.`);
  }

  let style: CellStyle | undefined;
  if (Object.hasOwn(cell, 'style') && cell.style !== null) {
    const preserveEmpty = (styledRows?.has(row) ?? false) || (styledCols?.has(column) ?? false);
    style = readStyle(requireObject(cell.style, 'style'), preserveEmpty);
  }

  let cache: CellCache | undefined;
  if (Object.hasOwn(cell, 'cache') && cell.cache !== null) {
    cache = readCache(requireObject(cell.cache, 'cache'));
  }

  const requiresExternalCalculation = Object.hasOwn(cell, 'requiresExternalCalculation')
    ? requireBool(cell.requiresExternalCalculation, 'requiresExternalCalculation')
    : undefined;

  const spillOwner = Object.hasOwn(cell, 'spillOwner') && cell.spillOwner !== null
    ? requireGuid(cell.spillOwner, 'spillOwner')
    : undefined;

  let spillKind: CellDocument['spillKind'];
  if (Object.hasOwn(cell, 'spillKind') && cell.spillKind !== null) {
    const rawKind = requireString(cell.spillKind, 'spillKind');
    if (rawKind !== 'python' && rawKind !== 'formula' && rawKind !== 'pivot') {
      throw new PackageException(`Invalid spillKind '${rawKind}'.`);
    }
    spillKind = rawKind;
  }

  let spillExtent: CellDocument['spillExtent'];
  if (Object.hasOwn(cell, 'spillExtent') && cell.spillExtent !== null) {
    const ext = requireObject(cell.spillExtent, 'spillExtent');
    rejectUnknown(ext, 'rows', 'columns');
    const rows = requireInt(requireProperty(ext, 'rows'), 'rows');
    const columns = requireInt(requireProperty(ext, 'columns'), 'columns');
    if (rows < 1 || columns < 1 || rows > documentLimits.rows || columns > documentLimits.columns || rows * columns > 100_000) {
      throw new PackageException('Invalid spillExtent.');
    }
    spillExtent = { rows, columns };
  }

  return {
    id,
    row,
    column,
    content,
    ...(style !== undefined ? { style } : {}),
    ...(cache ? { cache } : {}),
    ...(requiresExternalCalculation ? { requiresExternalCalculation: true } : {}),
    ...(spillOwner ? { spillOwner } : {}),
    ...(spillKind ? { spillKind } : {}),
    ...(spillExtent ? { spillExtent } : {}),
  };
}

function readStyle(obj: Record<string, unknown>, preserveEmpty = false): CellStyle | undefined {
  rejectUnknown(
    obj,
    'bold', 'numberFormat', 'fill', 'fontFamily', 'fontSize', 'italic', 'underline',
    'textRgb', 'borders', 'alignment', 'wrap', 'doubleUnderline', 'verticalAlignment',
    'strikethrough', 'script', 'underlineKind', 'fillDetails', 'borderDetails',
    'textRotation', 'indent', 'shrinkToFit', 'readingOrder', 'justifyLastLine',
  );

  const bold = Object.hasOwn(obj, 'bold') ? requireBool(obj.bold, 'bold') : false;
  const numberFormat = Object.hasOwn(obj, 'numberFormat') ? requireString(obj.numberFormat, 'numberFormat') : undefined;
  const fill = Object.hasOwn(obj, 'fill') ? requireString(obj.fill, 'fill').toUpperCase() : undefined;
  const fontFamily = Object.hasOwn(obj, 'fontFamily') ? requireString(obj.fontFamily, 'fontFamily') : undefined;
  const fontSize = Object.hasOwn(obj, 'fontSize') ? requireFinite(obj.fontSize, 'fontSize') : undefined;
  const italic = Object.hasOwn(obj, 'italic') ? requireBool(obj.italic, 'italic') : false;
  const underline = Object.hasOwn(obj, 'underline') ? requireBool(obj.underline, 'underline') : false;
  const textRgb = Object.hasOwn(obj, 'textRgb') ? requireString(obj.textRgb, 'textRgb') : undefined;
  const borders = Object.hasOwn(obj, 'borders') ? requireString(obj.borders, 'borders') : undefined;
  const alignment = Object.hasOwn(obj, 'alignment') ? requireString(obj.alignment, 'alignment') : undefined;
  const wrap = Object.hasOwn(obj, 'wrap') ? requireBool(obj.wrap, 'wrap') : false;
  const doubleUnderline = Object.hasOwn(obj, 'doubleUnderline') ? requireBool(obj.doubleUnderline, 'doubleUnderline') : false;
  const verticalAlignment = Object.hasOwn(obj, 'verticalAlignment') ? requireString(obj.verticalAlignment, 'verticalAlignment') : undefined;
  const strikethrough = Object.hasOwn(obj, 'strikethrough') ? requireBool(obj.strikethrough, 'strikethrough') : false;
  const script = Object.hasOwn(obj, 'script') ? requireString(obj.script, 'script') : undefined;
  const underlineKind = Object.hasOwn(obj, 'underlineKind') ? requireString(obj.underlineKind, 'underlineKind') : undefined;

  let fillDetails: CellFill | undefined;
  if (Object.hasOwn(obj, 'fillDetails') && obj.fillDetails !== null) {
    fillDetails = readFillDetails(requireObject(obj.fillDetails, 'fillDetails'));
  }

  let borderDetails: CellStyle['borderDetails'] | undefined;
  if (Object.hasOwn(obj, 'borderDetails') && obj.borderDetails !== null) {
    borderDetails = readBorderDetails(requireObject(obj.borderDetails, 'borderDetails'));
  }

  function parseIntegerProp(key: string): number {
    if (!Object.hasOwn(obj, key)) return 0;
    const val = obj[key];
    if (typeof val === 'number' && Number.isInteger(val)) return val;
    throw new PackageException('Invalid alignment integer.');
  }

  const textRotation = parseIntegerProp('textRotation');
  const indent = parseIntegerProp('indent');
  const shrinkToFit = Object.hasOwn(obj, 'shrinkToFit') ? requireBool(obj.shrinkToFit, 'shrinkToFit') : false;
  const readingOrder = Object.hasOwn(obj, 'readingOrder') ? requireString(obj.readingOrder, 'readingOrder') : undefined;
  const justifyLastLine = Object.hasOwn(obj, 'justifyLastLine') ? requireBool(obj.justifyLastLine, 'justifyLastLine') : false;

  const result: CellStyle = {
    ...(bold ? { bold: true } : {}),
    ...(numberFormat ? { numberFormat } : {}),
    ...(fill ? { fillRgb: fill } : {}),
    ...(fontFamily ? { fontFamily } : {}),
    ...(fontSize !== undefined ? { fontSize } : {}),
    ...(italic ? { italic: true } : {}),
    ...(underline ? { underline: true } : {}),
    ...(textRgb ? { textRgb } : {}),
    ...(borders ? { borders } : {}),
    ...(alignment ? { alignment } : {}),
    ...(wrap ? { wrap: true } : {}),
    ...(doubleUnderline ? { doubleUnderline: true } : {}),
    ...(verticalAlignment ? { verticalAlignment } : {}),
    ...(strikethrough ? { strikethrough: true } : {}),
    ...(script ? { script } : {}),
    ...(underlineKind ? { underlineKind } : {}),
    ...(fillDetails ? { fillDetails } : {}),
    ...(borderDetails ? { borderDetails } : {}),
    ...(textRotation !== 0 ? { textRotation } : {}),
    ...(indent !== 0 ? { indent } : {}),
    ...(shrinkToFit ? { shrinkToFit: true } : {}),
    ...(readingOrder ? { readingOrder } : {}),
    ...(justifyLastLine ? { justifyLastLine: true } : {}),
  };

  return Object.keys(result).length > 0 || preserveEmpty ? result : undefined;
}

function readBorderDetails(obj: Record<string, unknown>): CellStyle['borderDetails'] {
  rejectUnknown(obj, 'Top', 'Right', 'Bottom', 'Left', 'DiagonalUp', 'DiagonalDown');
  const res: Record<string, CellBorder> = {};
  const mapping: Record<string, 'top' | 'right' | 'bottom' | 'left' | 'diagonalUp' | 'diagonalDown'> = {
    Top: 'top', Right: 'right', Bottom: 'bottom', Left: 'left', DiagonalUp: 'diagonalUp', DiagonalDown: 'diagonalDown',
  };

  for (const [pascalKey, edgeVal] of Object.entries(obj)) {
    if (edgeVal === null || edgeVal === undefined) continue;
    const edgeObj = requireObject(edgeVal, 'border edge');
    rejectUnknown(edgeObj, 'Style', 'Rgb');
    const style = requireString(requireProperty(edgeObj, 'Style'), 'Style');
    const rgb = Object.hasOwn(edgeObj, 'Rgb') && edgeObj.Rgb !== null ? requireString(edgeObj.Rgb, 'Rgb') : undefined;
    res[mapping[pascalKey]!] = { style, ...(rgb ? { rgb } : {}) };
  }
  return res;
}

function readFillDetails(obj: Record<string, unknown>): CellFill {
  rejectUnknown(obj, 'Pattern', 'PatternRgb', 'GradientRgb', 'Shading', 'Variant');
  const patternRgb = Object.hasOwn(obj, 'PatternRgb') && obj.PatternRgb !== null ? requireString(obj.PatternRgb, 'PatternRgb') : undefined;
  const gradientRgb = Object.hasOwn(obj, 'GradientRgb') && obj.GradientRgb !== null ? requireString(obj.GradientRgb, 'GradientRgb') : undefined;
  const shading = Object.hasOwn(obj, 'Shading') && obj.Shading !== null ? requireString(obj.Shading, 'Shading') : undefined;
  return {
    pattern: Object.hasOwn(obj, 'Pattern') ? requireString(obj.Pattern, 'Pattern') : 'none',
    ...(patternRgb !== undefined ? { patternRgb } : {}),
    ...(gradientRgb !== undefined ? { gradientRgb } : {}),
    ...(shading !== undefined ? { shading } : {}),
    variant: Object.hasOwn(obj, 'Variant') ? requireInt(obj.Variant, 'Variant') : 0,
  };
}

function readCache(obj: Record<string, unknown>): CellCache {
  rejectUnknown(obj, 'kind', 'value');
  const kind = requireString(requireProperty(obj, 'kind'), 'kind');
  const val = requireProperty(obj, 'value');

  switch (kind) {
    case 'number':
      return { kind: 'number', value: requireFinite(val, 'value') };
    case 'text':
      return { kind: 'text', value: requireString(val, 'value') };
    case 'boolean':
      return { kind: 'boolean', value: requireBool(val, 'value') };
    case 'error': {
      const err = requireString(val, 'value');
      return { kind: 'error', value: err as typeof formulaErrors[number] };
    }
    default:
      throw new PackageException(`Unexpected cache kind '${kind}'.`);
  }
}

function readPython(raw: unknown): PythonBlock {
  const obj = requireObject(raw, 'Python block');
  rejectUnknown(
    obj,
    'Id', 'Source', 'Bindings', 'LastValue', 'HasResult', 'InputFingerprint',
    'RuntimeFingerprint', 'Status', 'Error', 'LastTable', 'SpillActive',
    'LastFigure', 'Placements',
  );

  const id = requireGuid(requireProperty(obj, 'Id'), 'Id');
  const source = requireString(requireProperty(obj, 'Source'), 'Source');

  const bindingsElement = requireProperty(obj, 'Bindings');
  if (!Array.isArray(bindingsElement)) {
    throw new PackageException('Invalid Python block: Bindings must be an array.');
  }

  const bindings: PythonBinding[] = [];
  for (const b of bindingsElement) {
    const bObj = requireObject(b, 'PythonBinding');
    rejectUnknown(bObj, 'Name', 'Reference');
    bindings.push({
      name: requireString(requireProperty(bObj, 'Name'), 'Name'),
      reference: requireString(requireProperty(bObj, 'Reference'), 'Reference'),
    });
  }

  let lastValue: CellCache | null = null;
  if (Object.hasOwn(obj, 'LastValue') && obj.LastValue !== null) {
    const lvObj = requireObject(obj.LastValue, 'LastValue');
    rejectUnknown(lvObj, 'Kind', 'Text');
    const lvKind = requireString(requireProperty(lvObj, 'Kind'), 'Kind');
    const lvText = requireString(requireProperty(lvObj, 'Text'), 'Text');
    lastValue = parseSerializedCache(lvKind, lvText, true);
  }

  let lastTable: PythonTable | null = null;
  if (Object.hasOwn(obj, 'LastTable') && obj.LastTable !== null) {
    const ltObj = requireObject(obj.LastTable, 'LastTable');
    rejectUnknown(ltObj, 'Rows', 'Columns', 'Cells', 'Notices');
    const tRows = requireInt(requireProperty(ltObj, 'Rows'), 'Rows');
    const tCols = requireInt(requireProperty(ltObj, 'Columns'), 'Columns');
    const cellsVal = requireProperty(ltObj, 'Cells');
    if (!Array.isArray(cellsVal)) throw new PackageException('Invalid Python block: Table Cells must be an array.');
    const noticesVal = requireProperty(ltObj, 'Notices');
    if (!Array.isArray(noticesVal)) throw new PackageException('Invalid Python block: Table Notices must be an array.');

    const cells: (CellCache | null)[] = cellsVal.map(item => {
      if (item === null) return null;
      const cObj = requireObject(item, 'Table cell cache');
      rejectUnknown(cObj, 'Kind', 'Text');
      return parseSerializedCache(requireString(requireProperty(cObj, 'Kind'), 'Kind'), requireString(requireProperty(cObj, 'Text'), 'Text'), true);
    });

    const notices = noticesVal.map(n => requireString(n, 'Notice'));
    lastTable = { rows: tRows, columns: tCols, cells, notices };
  }

  let lastFigure: PythonFigure | null = null;
  if (Object.hasOwn(obj, 'LastFigure') && obj.LastFigure !== null) {
    const lfObj = requireObject(obj.LastFigure, 'LastFigure');
    rejectUnknown(lfObj, 'PngBase64', 'Width', 'Height');
    lastFigure = {
      pngBase64: requireString(requireProperty(lfObj, 'PngBase64'), 'PngBase64'),
      width: requireInt(requireProperty(lfObj, 'Width'), 'Width'),
      height: requireInt(requireProperty(lfObj, 'Height'), 'Height'),
    };
  }

  let placements: FigurePlacement[] = [];
  if (Object.hasOwn(obj, 'Placements') && obj.Placements !== null) {
    if (!Array.isArray(obj.Placements)) throw new PackageException('Invalid Python block: Placements must be an array.');
    placements = obj.Placements.map(item => {
      const pObj = requireObject(item, 'FigurePlacement');
      rejectUnknown(pObj, 'Id', 'SheetId', 'Row', 'Column', 'Rows', 'Columns', 'OffsetX', 'OffsetY', 'Width', 'Height', 'RenderedFigure', 'RenderedSourceId');
      let renderedFig: PythonFigure | null = null;
      if (Object.hasOwn(pObj, 'RenderedFigure') && pObj.RenderedFigure !== null) {
        const rfObj = requireObject(pObj.RenderedFigure, 'RenderedFigure');
        rejectUnknown(rfObj, 'PngBase64', 'Width', 'Height');
        renderedFig = {
          pngBase64: requireString(requireProperty(rfObj, 'PngBase64'), 'PngBase64'),
          width: requireInt(requireProperty(rfObj, 'Width'), 'Width'),
          height: requireInt(requireProperty(rfObj, 'Height'), 'Height'),
        };
      }
      return {
        id: requireGuid(requireProperty(pObj, 'Id'), 'Id'),
        sheetId: requireGuid(requireProperty(pObj, 'SheetId'), 'SheetId'),
        row: requireInt(requireProperty(pObj, 'Row'), 'Row'),
        column: requireInt(requireProperty(pObj, 'Column'), 'Column'),
        rows: requireInt(requireProperty(pObj, 'Rows'), 'Rows'),
        columns: requireInt(requireProperty(pObj, 'Columns'), 'Columns'),
        offsetX: Object.hasOwn(pObj, 'OffsetX') ? requireFinite(pObj.OffsetX, 'OffsetX') : 0,
        offsetY: Object.hasOwn(pObj, 'OffsetY') ? requireFinite(pObj.OffsetY, 'OffsetY') : 0,
        width: Object.hasOwn(pObj, 'Width') ? requireFinite(pObj.Width, 'Width') : 0,
        height: Object.hasOwn(pObj, 'Height') ? requireFinite(pObj.Height, 'Height') : 0,
        renderedFigure: renderedFig,
        renderedSourceId: Object.hasOwn(pObj, 'RenderedSourceId') && pObj.RenderedSourceId !== null ? requireString(pObj.RenderedSourceId, 'RenderedSourceId') : null,
      };
    });
  }

  return {
    id,
    source,
    bindings,
    lastValue,
    hasResult: Object.hasOwn(obj, 'HasResult') ? requireBool(obj.HasResult, 'HasResult') : false,
    inputFingerprint: Object.hasOwn(obj, 'InputFingerprint') && obj.InputFingerprint !== null ? requireString(obj.InputFingerprint, 'InputFingerprint') : null,
    runtimeFingerprint: Object.hasOwn(obj, 'RuntimeFingerprint') && obj.RuntimeFingerprint !== null ? requireString(obj.RuntimeFingerprint, 'RuntimeFingerprint') : null,
    status: Object.hasOwn(obj, 'Status') ? requireString(obj.Status, 'Status') : 'never',
    error: Object.hasOwn(obj, 'Error') && obj.Error !== null ? requireString(obj.Error, 'Error') : null,
    lastTable,
    spillActive: Object.hasOwn(obj, 'SpillActive') ? requireBool(obj.SpillActive, 'SpillActive') : false,
    lastFigure,
    placements,
  };
}

function parseSerializedCache(kind: string, text: string, preserveNumberText = false): CellCache {
  switch (kind) {
    case 'number': {
      if (!/^[\t\n\v\f\r ]*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?[\t\n\v\f\r ]*$/u.test(text)) {
        throw new PackageException('Invalid table number.');
      }
      const num = Number(text);
      if (!Number.isFinite(num)) throw new PackageException('Invalid table number.');
      return { kind: 'number', value: num, ...(preserveNumberText ? { serializedText: text } : {}) };
    }
    case 'boolean': {
      if (text !== 'true' && text !== 'false') throw new PackageException('Invalid table boolean.');
      return { kind: 'boolean', value: text === 'true' };
    }
    case 'text':
      return { kind: 'text', value: text };
    case 'error':
      return { kind: 'error', value: text as typeof formulaErrors[number] };
    default:
      throw new PackageException(`Unsupported table value.`);
  }
}

function serializeCache(cache: CellCache): { Kind: string; Text: string } {
  switch (cache.kind) {
    case 'number':
      return {
        Kind: 'number',
        Text: cache.serializedText !== undefined && Number(cache.serializedText) === cache.value
          ? cache.serializedText
          : String(cache.value),
      };
    case 'boolean':
      return { Kind: 'boolean', Text: cache.value ? 'true' : 'false' };
    case 'text':
      return { Kind: 'text', Text: cache.value };
    case 'error':
      return { Kind: 'error', Text: cache.value };
  }
}

export async function saveNativePackage(document: WorkbookDocument, request?: SaveRequest): Promise<Blob> {
  const owned = ownDto(document);
  validateWorkbook(owned);
  await validateNativeFigurePayloads(owned);

  const limits = { ...defaultPackageLimits, ...request?.limits };
  if (request?.fault === SaveFault.DuringWrite) {
    throw new Error('Injected write failure.');
  }

  const writer = new ZipWriter(new BlobWriter('application/zip'), { useWebWorkers: false, zip64: false });

  try {
    const manifestJson = JSON.stringify({
      formatVersion: owned.formatVersion,
      workbookId: owned.id,
      dateSystem: owned.dateSystem,
    }, null, 2);
    await writer.add('manifest.json', new TextReader(manifestJson));

    const workbookObj: Record<string, unknown> = {
      workbookId: owned.id,
    };
    if (owned.names && owned.names.length > 0) {
      workbookObj.names = owned.names.map(name => ({
        Id: name.id,
        Name: name.name,
        Area: {
          SheetId: name.area.sheetId,
          Row: name.area.row,
          Column: name.area.column,
          Rows: name.area.rows,
          Columns: name.area.columns,
          ...(name.area.invalid ? { IsInvalid: true } : {}),
        },
        ...(name.comment !== undefined ? { Comment: name.comment } : {}),
        ...(name.isTable ? { IsTable: true } : {}),
        ...(name.design ? {
          Design: {
            Style: name.design.style,
            HeaderRow: name.design.headerRow,
            TotalRow: name.design.totalRow,
            ...(name.design.firstColumn ? { FirstColumn: true } : {}),
            ...(name.design.lastColumn ? { LastColumn: true } : {}),
            ...(name.design.bandedRows !== undefined ? { BandedRows: name.design.bandedRows } : {}),
            ...(name.design.bandedColumns ? { BandedColumns: true } : {}),
            ...(name.design.filterButton !== undefined ? { FilterButton: name.design.filterButton } : {}),
          },
        } : {}),
        ...(name.headerLabels !== undefined ? { HeaderLabels: name.headerLabels } : {}),
      }));
    }
    if (owned.pivots && owned.pivots.length > 0) {
      workbookObj.pivots = owned.pivots;
    }
    workbookObj.sheets = owned.sheets.map(sheet => ({
      id: sheet.id,
      name: sheet.name,
    }));
    await writer.add('workbook.json', new TextReader(JSON.stringify(workbookObj, null, 2)));

    for (const sheet of owned.sheets) {
      const sheetObj: Record<string, unknown> = {
        sheetId: sheet.id,
        name: sheet.name,
        cells: sheet.cells.map(serializeCell),
      };

      if (sheet.merges && sheet.merges.length > 0) {
        sheetObj.merges = sheet.merges.map(m => ({
          Row: m.row,
          Column: m.column,
          Rows: m.rows,
          Columns: m.columns,
        }));
      }

      if (sheet.rows && sheet.rows.length > 0) {
        sheetObj.rows = sheet.rows.map(r => ({
          index: r.index,
          ...(r.size !== undefined ? { size: r.size } : {}),
          hidden: r.hidden ?? false,
          ...(r.style !== undefined ? { style: serializeStyle(r.style) } : {}),
        }));
      }

      if (sheet.columns && sheet.columns.length > 0) {
        sheetObj.columns = sheet.columns.map(c => ({
          index: c.index,
          ...(c.size !== undefined ? { size: c.size } : {}),
          hidden: c.hidden ?? false,
          ...(c.style !== undefined ? { style: serializeStyle(c.style) } : {}),
        }));
      }

      if (sheet.frozenPanes.rows === 1 && sheet.frozenPanes.columns === 0) {
        sheetObj.freeze = 'topRow';
      } else if (sheet.frozenPanes.rows === 0 && sheet.frozenPanes.columns === 1) {
        sheetObj.freeze = 'firstColumn';
      } else if (sheet.frozenPanes.rows > 0 || sheet.frozenPanes.columns > 0) {
        sheetObj.freeze = {
          rows: sheet.frozenPanes.rows,
          columns: sheet.frozenPanes.columns,
        };
      }

      if (sheet.filter) {
        sheetObj.filter = {
          row: sheet.filter.row,
          column: sheet.filter.column,
          rows: sheet.filter.rows,
          columns: sheet.filter.columns,
          ...(sheet.filter.sort ? {
            sort: {
              column: sheet.filter.sort.column,
              ascending: sheet.filter.sort.ascending,
            },
          } : {}),
          criteria: sheet.filter.criteria.map(c => ({
            column: c.column,
            allowed: c.allowed,
          })),
        };
      }

      await writer.add(`sheets/${sheet.id}.json`, new TextReader(JSON.stringify(sheetObj, null, 2)));
    }
  } catch (err) {
    await writer.close().catch(() => {});
    throw err;
  }

  const blob = await writer.close();

  // Verification step
  const loaded = await loadNativePackage(blob, limits);
  if (!equalDto(owned, loaded)) {
    throw new PackageException(`Saved package did not match the document at ${firstDtoDifference(owned, loaded) ?? 'an unknown field'}.`);
  }

  request?.beforeReplace?.();
  if (request?.fault === SaveFault.BeforeReplace) {
    throw new Error('Injected replace failure.');
  }

  return blob;
}

function serializeStyle(s: CellStyle): Record<string, unknown> {
  return {
    bold: s.bold ?? false,
    ...(s.numberFormat !== undefined ? { numberFormat: s.numberFormat } : {}),
    ...(s.fillRgb !== undefined ? { fill: s.fillRgb } : {}),
    ...(s.fontFamily !== undefined ? { fontFamily: s.fontFamily } : {}),
    ...(s.fontSize !== undefined ? { fontSize: s.fontSize } : {}),
    italic: s.italic ?? false,
    underline: s.underline ?? false,
    ...(s.textRgb !== undefined ? { textRgb: s.textRgb } : {}),
    ...(s.borders !== undefined ? { borders: s.borders } : {}),
    ...(s.verticalAlignment !== undefined ? { verticalAlignment: s.verticalAlignment } : {}),
    ...(s.alignment !== undefined ? { alignment: s.alignment } : {}),
    wrap: s.wrap ?? false,
    ...(s.textRotation !== undefined && s.textRotation !== 0 ? { textRotation: s.textRotation } : {}),
    ...(s.indent !== undefined && s.indent !== 0 ? { indent: s.indent } : {}),
    ...(s.shrinkToFit ? { shrinkToFit: true } : {}),
    ...(s.readingOrder !== undefined ? { readingOrder: s.readingOrder } : {}),
    ...(s.justifyLastLine ? { justifyLastLine: true } : {}),
    ...(s.strikethrough ? { strikethrough: true } : {}),
    ...(s.script !== undefined ? { script: s.script } : {}),
    ...(s.underlineKind !== undefined ? { underlineKind: s.underlineKind } : {}),
    ...(s.borderDetails ? {
      borderDetails: {
        ...(s.borderDetails.top ? { Top: { Style: s.borderDetails.top.style, ...(s.borderDetails.top.rgb ? { Rgb: s.borderDetails.top.rgb } : {}) } } : {}),
        ...(s.borderDetails.right ? { Right: { Style: s.borderDetails.right.style, ...(s.borderDetails.right.rgb ? { Rgb: s.borderDetails.right.rgb } : {}) } } : {}),
        ...(s.borderDetails.bottom ? { Bottom: { Style: s.borderDetails.bottom.style, ...(s.borderDetails.bottom.rgb ? { Rgb: s.borderDetails.bottom.rgb } : {}) } } : {}),
        ...(s.borderDetails.left ? { Left: { Style: s.borderDetails.left.style, ...(s.borderDetails.left.rgb ? { Rgb: s.borderDetails.left.rgb } : {}) } } : {}),
        ...(s.borderDetails.diagonalUp ? { DiagonalUp: { Style: s.borderDetails.diagonalUp.style, ...(s.borderDetails.diagonalUp.rgb ? { Rgb: s.borderDetails.diagonalUp.rgb } : {}) } } : {}),
        ...(s.borderDetails.diagonalDown ? { DiagonalDown: { Style: s.borderDetails.diagonalDown.style, ...(s.borderDetails.diagonalDown.rgb ? { Rgb: s.borderDetails.diagonalDown.rgb } : {}) } } : {}),
      },
    } : {}),
    ...(s.fillDetails ? {
      fillDetails: {
        Pattern: s.fillDetails.pattern,
        ...(s.fillDetails.patternRgb ? { PatternRgb: s.fillDetails.patternRgb } : {}),
        ...(s.fillDetails.gradientRgb ? { GradientRgb: s.fillDetails.gradientRgb } : {}),
        ...(s.fillDetails.shading ? { Shading: s.fillDetails.shading } : {}),
        Variant: s.fillDetails.variant ?? 0,
      },
    } : {}),
    ...(s.doubleUnderline ? { doubleUnderline: true } : {}),
  };
}

function serializeCell(cell: CellDocument): Record<string, unknown> {
  const obj: Record<string, unknown> = {
    id: cell.id,
    row: cell.row,
    column: cell.column,
  };

  if (cell.spillOwner !== undefined) {
    obj.spillOwner = cell.spillOwner;
  }
  if (cell.spillKind !== undefined) {
    obj.spillKind = cell.spillKind;
  }
  if (cell.spillExtent !== undefined) {
    obj.spillExtent = { rows: cell.spillExtent.rows, columns: cell.spillExtent.columns };
  }

  const content: Record<string, unknown> = { kind: cell.content.kind };
  switch (cell.content.kind) {
    case 'python': {
      const blk = cell.content.block ?? {
        id: cell.content.blockId,
        source: cell.content.source,
        bindings: [],
        status: 'never',
        hasResult: false,
        spillActive: false,
        placements: [],
      };
      content.block = {
        Id: blk.id,
        Source: blk.source,
        Bindings: blk.bindings.map(b => ({ Name: b.name, Reference: b.reference })),
        LastValue: blk.lastValue ? serializeCache(blk.lastValue) : null,
        HasResult: blk.hasResult ?? false,
        InputFingerprint: blk.inputFingerprint ?? null,
        RuntimeFingerprint: blk.runtimeFingerprint ?? null,
        Status: blk.status ?? 'never',
        Error: blk.error ?? null,
        LastTable: blk.lastTable ? {
          Rows: blk.lastTable.rows,
          Columns: blk.lastTable.columns,
          Cells: blk.lastTable.cells.map(c => c ? serializeCache(c) : null),
          Notices: blk.lastTable.notices,
        } : null,
        SpillActive: blk.spillActive ?? false,
        LastFigure: blk.lastFigure ? {
          PngBase64: blk.lastFigure.pngBase64,
          Width: blk.lastFigure.width,
          Height: blk.lastFigure.height,
        } : null,
        Placements: (blk.placements ?? []).map(p => ({
          Id: p.id,
          SheetId: p.sheetId,
          Row: p.row,
          Column: p.column,
          Rows: p.rows,
          Columns: p.columns,
          OffsetX: p.offsetX ?? 0,
          OffsetY: p.offsetY ?? 0,
          Width: p.width ?? 0,
          Height: p.height ?? 0,
          RenderedFigure: p.renderedFigure ? {
            PngBase64: p.renderedFigure.pngBase64,
            Width: p.renderedFigure.width,
            Height: p.renderedFigure.height,
          } : null,
          RenderedSourceId: p.renderedSourceId ?? null,
        })),
      };
      break;
    }
    case 'number':
    case 'text':
    case 'boolean':
      content.value = cell.content.value;
      break;
    case 'formula':
      content.source = cell.content.source;
      break;
    case 'blank':
      break;
  }
  obj.content = content;

  if (cell.style) {
    obj.style = serializeStyle(cell.style);
  }

  if (cell.requiresExternalCalculation) {
    obj.requiresExternalCalculation = true;
  }

  if (cell.cache) {
    obj.cache = {
      kind: cell.cache.kind,
      value: cell.cache.value,
    };
  }

  return obj;
}

export const NativePackage = {
  load: loadNativePackage,
  save: saveNativePackage,
};
