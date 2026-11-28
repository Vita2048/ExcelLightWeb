import { formatCellValue } from '../core/formats/cell-format';
import type { CacheUpdate, FormulaSpillResult } from '../application/workbook-protocol';
import {
  parseFormula, referenceBox, UnsupportedFormulaReference,
  type FormulaExpression, type FormulaReference,
} from '../core/formula-references';
import { maskRows, RowMask } from '../core/row-mask';
import {
  cellKey, documentLimits, resolvedCellStyle,
  type CellCache, type CellDocument, type CellExtent, type CellReference, type SheetDocument, type WorkbookDocument,
} from '../core/workbook';
import { lastColumn, lastRow, resolveStructuredReference, resolveWorkbookName } from '../core/workbook-names';
import { overlaps, ordinalNameKey } from '../core/workbook-validation';
import { nowSerial, todaySerial } from './formula-dates';
import { evaluateFunction, isVolatileFunction, sumIfRange, supportedFunctions } from './formula-functions';
import { findPivotContainingCell, queryPivotData, type PivotEvaluationCache } from './pivot-engine';
import {
  arrayFromMatrix, compare, error, finite, isArray, isError, isLambda, isOmitted, isRange,
  matchesComparison, numberValue, omitted, scalarCache, textValue,
  type FormulaArray, type FormulaError, type FormulaLambda, type FormulaValue, type Range, type Scalar,
} from './formula-values';

interface FormulaNode {
  readonly id: string;
  readonly target: CellReference;
  readonly expression?: FormulaExpression;
  readonly ranges: readonly Range[];
  readonly inputs: FormulaNode[];
  readonly dependents: FormulaNode[];
  readonly containsSubtotal: boolean;
  readonly usesGetPivotData: boolean;
  readonly volatile: boolean;
  readonly hasSpillRef: boolean;
  readonly diagnostic?: 'malformed' | 'unsupported';
}
export interface NativeCalculation {
  readonly updates: readonly CacheUpdate[];
  /** Cell IDs key the analysis. Inputs include formulas referenced from inactive branches. */
  readonly dependencies: ReadonlyMap<string, readonly CellReference[]>;
  readonly diagnostics: ReadonlyMap<string, string>;
}

function address(sheet: SheetDocument, cell: CellDocument): CellReference {
  return { sheetId: sheet.id, row: cell.row, column: cell.column };
}

function baseValue(cell: CellDocument | undefined): Scalar {
  if (!cell) return null;
  if (cell.content.kind === 'blank') {
    return (cell.spillKind === 'formula' || cell.spillKind === 'pivot') && cell.cache?.kind === 'error' ? cell.cache : null;
  }
  if (cell.content.kind === 'python') {
    const block = cell.content.block;
    const value = block?.hasResult ? block.lastValue : undefined;
    return value?.kind === 'error' ? value : value?.value ?? null;
  }
  if (cell.content.kind === 'formula') {
    return cell.cache?.kind === 'error' ? cell.cache : cell.cache?.value ?? null;
  }
  return cell.content.value;
}

export function* cellsIn<T extends { row: number; column: number }>(cells: readonly T[], range: Range): IterableIterator<T> {
  let low = 0, high = cells.length;
  while (low < high) {
    const mid = (low + high) >>> 1;
    if (cells[mid]!.row < range.top) low = mid + 1;
    else high = mid;
  }
  for (let i = low; i < cells.length; i++) {
    const cell = cells[i]!;
    if (cell.row > range.bottom) break;
    if (cell.column >= range.left && cell.column <= range.right) yield cell;
  }
}

function evaluationOrder(nodes: Iterable<FormulaNode>) {
  const indexes = new Map<FormulaNode, number>();
  const lowLinks = new Map<FormulaNode, number>();
  const componentStack: FormulaNode[] = [];
  const active = new Set<FormulaNode>();
  const cycles = new Set<FormulaNode>();
  const order: FormulaNode[] = [];
  const stack: { node: FormulaNode; next: number }[] = [];
  function push(node: FormulaNode) {
    indexes.set(node, indexes.size);
    lowLinks.set(node, indexes.get(node)!);
    componentStack.push(node);
    active.add(node);
    stack.push({ node, next: 0 });
  }
  function lower(node: FormulaNode, index: number) {
    lowLinks.set(node, Math.min(lowLinks.get(node)!, index));
  }
  for (const node of nodes) {
    if (indexes.has(node)) continue;
    push(node);
    while (stack.length) {
      const frame = stack.at(-1)!;
      const dependency = frame.node.inputs[frame.next++];
      if (dependency) {
        const index = indexes.get(dependency);
        if (index === undefined) push(dependency);
        else if (active.has(dependency)) lower(frame.node, index);
        continue;
      }
      stack.pop();
      const low = lowLinks.get(frame.node)!;
      const parent = stack.at(-1)?.node;
      if (parent) lower(parent, low);
      if (low !== indexes.get(frame.node)) continue;
      const component: FormulaNode[] = [];
      let member: FormulaNode;
      do {
        member = componentStack.pop()!;
        active.delete(member);
        component.push(member);
      } while (member !== frame.node);
      if (component.length > 1 || frame.node.inputs.includes(frame.node)) {
        for (const member of component) cycles.add(member);
      }
      order.push(...component);
    }
  }
  return { order, cycles };
}

function filterVisibility(sheet: SheetDocument, keyOf: (cell: CellDocument | undefined, sheet?: SheetDocument) => string): RowMask {
  const filter = sheet.filter;
  if (!filter?.criteria.length) return RowMask.none();
  const from = filter.row + 1;
  const to = filter.row + filter.rows - 1;
  const rules = filter.criteria.map(criterion => ({ column: criterion.column, allowed: new Set(criterion.allowed) }));
  const columns = new Set(rules.map(rule => rule.column));
  const cells = [];
  for (const cell of sheet.cells) {
    if (cell.row < from || cell.row > to || !columns.has(cell.column)) continue;
    cells.push({ row: cell.row, column: cell.column, key: keyOf(cell, sheet) });
  }
  return maskRows(from, to, rules, keyOf(undefined, sheet), cells);
}

function defaultFilterText(cell: CellDocument | undefined, sheet?: SheetDocument): string {
  return formatCellValue(cell, cell && sheet ? resolvedCellStyle(sheet, cell, cell)?.numberFormat : undefined).text;
}

export interface CalculationOptions {
  readonly clock?: () => Date;
  readonly locale?: string;
  /** Override the shared display key for a caller-supplied set of filter criteria. */
  readonly filterText?: (cell: CellDocument | undefined) => string;
  /** Deterministic PRNG hook for volatile random functions such as RANDARRAY. */
  readonly random?: () => number;
}

export function prepareNativeGraph(document: WorkbookDocument) {
  const sheets = new Map(document.sheets.map(sheet => [ordinalNameKey(sheet.name), sheet]));
  const formulas = new Map(document.sheets.map(sheet => [sheet.id, sheet.cells.filter(cell => cell.content.kind === 'formula')
    .sort((a, b) => cellKey(a) - cellKey(b))]));
  const cellLookup = new Map(document.sheets.map(sheet => [sheet.id, new Map(sheet.cells.map(cell => [cellKey(cell), cell]))]));
  const nodes = new Map<string, FormulaNode>();
  const resolvedReferences = new Map<FormulaReference, Range>();
  function resolve(sheet: SheetDocument, ref: FormulaReference): Range | undefined {
    const target = ref.sheet === undefined ? sheet : sheets.get(ordinalNameKey(ref.sheet));
    const bounds = referenceBox(ref);
    if (!target || bounds.bottom > documentLimits.rows || bounds.right > documentLimits.columns) return;
    return { sheetId: target.id, ...bounds };
  }

  for (const sheet of document.sheets) for (const cell of formulas.get(sheet.id)!) {
    if (cell.content.kind !== 'formula') continue;
    const target = address(sheet, cell);
    const base = { id: cell.id, target, inputs: [], dependents: [] };
    let parsed: ReturnType<typeof parseFormula>;
    try { parsed = parseFormula(cell.content.source); }
    catch (exception) {
      nodes.set(cell.id, { ...base, ranges: [], containsSubtotal: false, usesGetPivotData: false, volatile: false, hasSpillRef: false,
        diagnostic: cell.requiresExternalCalculation || exception instanceof UnsupportedFormulaReference ? 'unsupported' : 'malformed' });
      continue;
    }

    let containsSubtotal = false;
    let usesGetPivotData = false;
    let volatile = false;
    let hasSpillRef = false;
    let unsupported = !!cell.requiresExternalCalculation;
    const ranges: Range[] = [];

    for (const ref of parsed.references) {
      if (ref.endSheet !== undefined) {
        unsupported = true;
        continue;
      }
      const range = resolve(sheet, ref);
      if (!range) continue;
      resolvedReferences.set(ref, range);
      ranges.push(range);
    }

    function analyze(expression: FormulaExpression, scope: ReadonlySet<string>): void {
      switch (expression.kind) {
        case 'name': {
          if (!scope.has(expression.name.toUpperCase())) {
            const area = resolveWorkbookName(document, expression.name);
            if (!area) unsupported = true;
            else if (!area.invalid) {
              ranges.push({ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) });
            }
          }
          break;
        }
        case 'structured': {
          const area = resolveStructuredReference(document, expression.reference, target);
          if (!area) {
            if (expression.reference.table &&
                !document.names.some(n => n.isTable && ordinalNameKey(n.name) === ordinalNameKey(expression.reference.table!))) {
              unsupported = true;
            }
          } else if (!area.invalid) {
            ranges.push({ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) });
          }
          break;
        }
        case 'spill': {
          hasSpillRef = true;
          analyze(expression.target, scope);
          if (expression.target.kind === 'reference') {
            const anchorRange = resolvedReferences.get(expression.target.reference);
            if (anchorRange && anchorRange.top === anchorRange.bottom && anchorRange.left === anchorRange.right) {
              const anchorCell = cellLookup.get(anchorRange.sheetId)?.get(cellKey({ row: anchorRange.top, column: anchorRange.left }));
              if (anchorCell?.spillExtent) {
                ranges.push({
                  sheetId: anchorRange.sheetId,
                  top: anchorRange.top,
                  left: anchorRange.left,
                  bottom: anchorRange.top + anchorCell.spillExtent.rows - 1,
                  right: anchorRange.left + anchorCell.spillExtent.columns - 1,
                });
              }
            }
          }
          break;
        }
        case 'invoke':
          analyze(expression.callee, scope);
          for (const arg of expression.args) analyze(arg, scope);
          break;
        case 'union':
          for (const part of expression.parts) analyze(part, scope);
          break;
        case 'unary':
          analyze(expression.value, scope);
          break;
        case 'binary':
          analyze(expression.left, scope);
          analyze(expression.right, scope);
          break;
        case 'call': {
          if (expression.name === 'LET') {
            const inner = new Set(scope);
            for (let i = 0; i < expression.args.length - 1; i += 2) {
              if (i + 1 < expression.args.length) analyze(expression.args[i + 1]!, inner);
              const nameNode = expression.args[i];
              if (nameNode?.kind === 'name') inner.add(nameNode.name.toUpperCase());
            }
            if (expression.args.length > 0 && expression.args.length % 2 === 1) {
              analyze(expression.args.at(-1)!, inner);
            }
            break;
          }
          if (expression.name === 'LAMBDA') {
            const inner = new Set(scope);
            for (let i = 0; i < expression.args.length - 1; i++) {
              const param = expression.args[i];
              if (param?.kind === 'name') inner.add(param.name.toUpperCase());
            }
            if (expression.args.length > 0) analyze(expression.args.at(-1)!, inner);
            break;
          }
          if (!scope.has(expression.name)) {
            if (!supportedFunctions.has(expression.name)) unsupported = true;
            if (expression.name === 'SUBTOTAL') containsSubtotal = true;
            if (expression.name === 'GETPIVOTDATA') usesGetPivotData = true;
            if (isVolatileFunction(expression.name)) volatile = true;
          }
          for (const arg of expression.args) analyze(arg, scope);
          break;
        }
      }
    }

    analyze(parsed.expression, new Set());

    function candidateRanges(expression: FormulaExpression | undefined): Range[] {
      if (expression?.kind === 'reference') {
        const range = resolvedReferences.get(expression.reference);
        return range ? [range] : [];
      }
      if (expression?.kind === 'name') {
        const area = resolveWorkbookName(document, expression.name);
        return area && !area.invalid ? [{ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) }] : [];
      }
      if (expression?.kind === 'structured') {
        const area = resolveStructuredReference(document, expression.reference, target);
        return area && !area.invalid ? [{ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) }] : [];
      }
      if (expression?.kind === 'call' && expression.name === 'IF') return expression.args.slice(1).flatMap(candidateRanges);
      return [];
    }
    for (const expr of walkCalls(parsed.expression)) {
      if (expr.kind !== 'call' || expr.name !== 'SUMIF' || expr.args.length !== 3) continue;
      const [criteria, , sum] = expr.args;
      for (const source of candidateRanges(criteria)) for (const sumTarget of candidateRanges(sum)) {
        const expanded = sumIfRange(source, sumTarget);
        if (expanded) ranges.push(expanded);
      }
    }
    if (usesGetPivotData && document.pivots?.length) {
      for (const pivot of document.pivots) {
        const pivotBox = pivot.extent ?? { row: pivot.target.row, column: pivot.target.column, rows: 1, columns: 1 };
        const hitsPivot = ranges.some(range =>
          range.sheetId === pivot.target.sheetId &&
          overlaps({ row: range.top, column: range.left, rows: range.bottom - range.top + 1, columns: range.right - range.left + 1 }, pivotBox),
        );
        if (hitsPivot) {
          ranges.push({
            sheetId: pivot.target.sheetId,
            top: pivotBox.row,
            left: pivotBox.column,
            bottom: pivotBox.row + pivotBox.rows - 1,
            right: pivotBox.column + pivotBox.columns - 1,
          });
        }
      }
    }
    nodes.set(cell.id, { ...base, expression: parsed.expression, ranges, containsSubtotal, usesGetPivotData, volatile, hasSpillRef,
      ...(unsupported && { diagnostic: 'unsupported' as const }) });
  }
  const dependencies = new Map<string, readonly CellReference[]>();
  for (const [id, node] of nodes) {
    const inputs = new Set<FormulaNode>();
    for (const range of node.ranges) {
      for (const other of cellsIn(formulas.get(range.sheetId)!, range)) inputs.add(nodes.get(other.id)!);
      const anchor = cellLookup.get(range.sheetId)?.get(cellKey({ row: range.top, column: range.left }));
      if (anchor?.spillKind === 'formula' && anchor.spillOwner) {
        const ownerNode = nodes.get(anchor.spillOwner);
        if (ownerNode && ownerNode !== node) inputs.add(ownerNode);
      }
      // Also attach any formula anchor whose existing spill overlaps `range`.
      for (const formulaCell of formulas.get(range.sheetId)!) {
        if (!formulaCell.spillExtent) continue;
        const spillBox = {
          row: formulaCell.row,
          column: formulaCell.column,
          rows: formulaCell.spillExtent.rows,
          columns: formulaCell.spillExtent.columns,
        };
        const rangeBox = {
          row: range.top,
          column: range.left,
          rows: range.bottom - range.top + 1,
          columns: range.right - range.left + 1,
        };
        if (overlaps(spillBox, rangeBox)) {
          const ownerNode = nodes.get(formulaCell.id);
          if (ownerNode && ownerNode !== node) inputs.add(ownerNode);
        }
      }
    }
    const list = [...inputs];
    node.inputs.push(...list);
    for (const input of list) input.dependents.push(node);
    dependencies.set(id, list.map(input => input.target));
  }
  return { nodes, resolvedReferences, dependencies, names: document.names, pivots: document.pivots ?? [], ...evaluationOrder(nodes.values()) };
}

function* walkCalls(expression: FormulaExpression): Iterable<FormulaExpression> {
  const stack = [expression];
  while (stack.length) {
    const node = stack.pop()!;
    yield node;
    switch (node.kind) {
      case 'call': stack.push(...node.args); break;
      case 'invoke': stack.push(...node.args, node.callee); break;
      case 'spill': stack.push(node.target); break;
      case 'union': stack.push(...node.parts); break;
      case 'binary': stack.push(node.right, node.left); break;
      case 'unary': stack.push(node.value); break;
    }
  }
}

export function indexNativeInputs(document: WorkbookDocument) {
  return new Map(document.sheets.map(sheet => [sheet.id, {
    sheet, cells: new Map(sheet.cells.map(cell => [cellKey(cell), cell])),
  }]));
}

export function calculateNative(document: WorkbookDocument, options: CalculationOptions = {}): NativeCalculation {
  return evaluateNative(indexNativeInputs(document), prepareNativeGraph(document), undefined, options);
}

const emptyEnv: ReadonlyMap<string, FormulaValue> = new Map();

export function evaluateNative(inputs: ReturnType<typeof indexNativeInputs>, graph: ReturnType<typeof prepareNativeGraph>,
  selected: ReadonlySet<FormulaNode> | undefined, options: CalculationOptions, previewMode = false): NativeCalculation {
  const clockDate = options.clock?.() ?? new Date();
  const today = todaySerial(clockDate);
  const now = nowSerial(clockDate);
  const random = options.random ?? Math.random;
  const numberFormat = new Intl.NumberFormat(options.locale);
  const locale = numberFormat.resolvedOptions().locale;
  const parts = numberFormat.formatToParts(12345.6);
  const decimal = parts.find(part => part.type === 'decimal')?.value;
  const group = parts.find(part => part.type === 'group')?.value;
  const number = (value: Scalar) => numberValue(value, decimal, group);
  const text = (value: Exclude<Scalar, FormulaError>) => textValue(value, decimal);
  const { nodes, resolvedReferences, order, cycles, dependencies, names, pivots } = graph;
  const sorted = new Map<string, CellDocument[]>();
  const evaluated = new Map<CellDocument, Scalar>();
  const diagnostics = new Map<string, string>();
  const filtered = new Map<SheetDocument, RowMask>();
  const manual = new Map<SheetDocument, ReadonlySet<number>>();
  const pivotCache: PivotEvaluationCache = new Map();
  const docView = {
    id: '',
    formatVersion: 1 as const,
    dateSystem: '1900' as const,
    sheets: [...inputs.values()].map(entry => entry.sheet),
    names,
    ...(pivots.length > 0 && { pivots }),
  } satisfies WorkbookDocument;
  const sheetsByName = new Map(docView.sheets.map(entry => [ordinalNameKey(entry.name), entry]));

  function cellAt(sheetId: string, row: number, column: number): CellDocument | undefined {
    return inputs.get(sheetId)?.cells.get(cellKey({ row, column }));
  }

  const liveSpills = new Map<string, { readonly rows: number; readonly columns: number }>();
  const liveSheetSpills = new Map<string, Map<number, { readonly ownerId: string; readonly value: Scalar }>>();
  const clearedSheetSpills = new Map<string, Set<number>>();

  function computedValue(cell: CellDocument | undefined): Scalar {
    if (cell?.requiresExternalCalculation) return error('#NAME?');
    return (cell && evaluated.get(cell)) ?? baseValue(cell);
  }
  function read(range: Range, row: number, column: number): Scalar {
    const key = cellKey({ row, column });
    const live = liveSheetSpills.get(range.sheetId)?.get(key);
    if (live !== undefined) return live.value;
    if (clearedSheetSpills.get(range.sheetId)?.has(key)) return null;
    return computedValue(inputs.get(range.sheetId)!.cells.get(key));
  }

  function rangeCells(range: Range) {
    let cells = sorted.get(range.sheetId);
    if (!cells) {
      cells = [...inputs.get(range.sheetId)!.cells.values()].sort((a, b) => cellKey(a) - cellKey(b));
      sorted.set(range.sheetId, cells);
    }
    return cellsIn(cells, range);
  }

  function* entries(range: Range) {
    const liveMap = liveSheetSpills.get(range.sheetId);
    const clearedSet = clearedSheetSpills.get(range.sheetId);
    if (!liveMap?.size && !clearedSet?.size) {
      for (const cell of rangeCells(range)) {
        yield { row: cell.row, column: cell.column, value: computedValue(cell) };
      }
      return;
    }
    const seen = new Set<number>();
    for (const cell of rangeCells(range)) {
      const key = cellKey(cell);
      seen.add(key);
      const live = liveMap?.get(key);
      if (live !== undefined) yield { row: cell.row, column: cell.column, value: live.value };
      else if (!clearedSet?.has(key)) yield { row: cell.row, column: cell.column, value: computedValue(cell) };
    }
    if (liveMap) {
      for (const [key, entry] of liveMap) {
        if (seen.has(key)) continue;
        const row = Math.floor(key / documentLimits.columns) + 1;
        const column = (key % documentLimits.columns) + 1;
        if (row >= range.top && row <= range.bottom && column >= range.left && column <= range.right) {
          yield { row, column, value: entry.value };
        }
      }
    }
  }

  function asArray(value: FormulaValue, includeRanges: boolean): FormulaArray | undefined {
    if (isArray(value)) return value;
    if (!includeRanges || !isRange(value) || value.length !== 1) return;
    const area = value[0]!;
    const rows = area.bottom - area.top + 1, columns = area.right - area.left + 1;
    if (rows === 1 && columns === 1) return;
    return { kind: 'array', rows, columns, at: (row, column) => read(area, area.top + row, area.left + column) };
  }

  function arrayOf(value: FormulaValue): FormulaArray | undefined {
    if (isArray(value)) return value;
    if (isRange(value)) {
      if (value.length !== 1) return undefined;
      const area = value[0]!;
      const rows = area.bottom - area.top + 1, columns = area.right - area.left + 1;
      return { kind: 'array', rows, columns, at: (row, column) => read(area, area.top + row, area.left + column) };
    }
    if (isOmitted(value) || isLambda(value)) return undefined;
    return arrayFromMatrix(1, 1, [value]);
  }

  function* values(range: Range, subtotal?: { skipHidden: boolean }): Iterable<Scalar> {
    let excluded: RowMask | undefined, hidden: ReadonlySet<number> | undefined;
    if (subtotal) {
      const { sheet } = inputs.get(range.sheetId)!;
      if (!filtered.has(sheet)) filtered.set(sheet, filterVisibility(sheet, options.filterText ?? defaultFilterText));
      excluded = filtered.get(sheet);
      if (subtotal.skipHidden) {
        if (!manual.has(sheet)) manual.set(sheet, new Set(sheet.rows.filter(row => row.hidden).map(row => row.index)));
        hidden = manual.get(sheet);
      }
    }
    const liveMap = liveSheetSpills.get(range.sheetId);
    const clearedSet = clearedSheetSpills.get(range.sheetId);
    if (!liveMap?.size && !clearedSet?.size) {
      for (const cell of rangeCells(range)) {
        if (subtotal && (excluded?.has(cell.row) || hidden?.has(cell.row) || nodes.get(cell.id)?.containsSubtotal)) continue;
        yield computedValue(cell);
      }
      return;
    }
    for (const entry of entries(range)) {
      const existing = inputs.get(range.sheetId)!.cells.get(cellKey({ row: entry.row, column: entry.column }));
      if (subtotal && (excluded?.has(entry.row) || hidden?.has(entry.row) || (existing && nodes.get(existing.id)?.containsSubtotal))) continue;
      yield entry.value;
    }
  }

  function clearAnchorSpill(sheetId: string, anchor: CellDocument) {
    if (!anchor.spillExtent) return;
    let cleared = clearedSheetSpills.get(sheetId);
    if (!cleared) clearedSheetSpills.set(sheetId, cleared = new Set());
    const liveMap = liveSheetSpills.get(sheetId);
    for (let r = 0; r < anchor.spillExtent.rows; r++) for (let c = 0; c < anchor.spillExtent.columns; c++) {
      if (r === 0 && c === 0) continue;
      const key = cellKey({ row: anchor.row + r, column: anchor.column + c });
      cleared.add(key);
      if (liveMap?.get(key)?.ownerId === anchor.id) liveMap.delete(key);
    }
  }

  function trySpillResult(sheet: SheetDocument, anchor: CellDocument, rows: number, columns: number,
    at: (row: number, column: number) => Scalar): { cache: CellCache; spill?: FormulaSpillResult } {
    if (rows * columns <= 1) {
      return { cache: scalarCache(at(0, 0)) };
    }
    if (rows * columns > 100_000 || anchor.row + rows - 1 > documentLimits.rows || anchor.column + columns - 1 > documentLimits.columns) {
      return { cache: error('#SPILL!') };
    }
    const extent: CellExtent = { row: anchor.row, column: anchor.column, rows, columns };
    if (sheet.merges.some(merge => overlaps(merge, extent))) {
      return { cache: error('#SPILL!') };
    }
    if (names.some(name => name.isTable && !name.area.invalid && name.area.sheetId === sheet.id && overlaps(name.area, extent))) {
      return { cache: error('#SPILL!') };
    }

    const sheetCells = inputs.get(sheet.id)!.cells;
    const liveMap = liveSheetSpills.get(sheet.id);
    const clearedSet = clearedSheetSpills.get(sheet.id);
    for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) {
      if (r === 0 && c === 0) continue;
      const key = cellKey({ row: anchor.row + r, column: anchor.column + c });
      const live = liveMap?.get(key);
      if (live !== undefined && live.ownerId !== anchor.id) {
        return { cache: error('#SPILL!') };
      }
      if (clearedSet?.has(key)) continue;
      const existing = sheetCells.get(key);
      if (existing && (existing.content.kind !== 'blank' || (existing.spillOwner !== undefined && existing.spillOwner !== anchor.id) ||
          (existing.spillKind === 'formula' && existing.spillOwner !== anchor.id))) {
        return { cache: error('#SPILL!') };
      }
    }

    const cells: (CellCache | null)[] = new Array(rows * columns);
    const scalars: Scalar[] = new Array(rows * columns);
    for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) {
      const val = at(r, c);
      const index = r * columns + c;
      scalars[index] = val;
      cells[index] = val === null ? null : scalarCache(val);
    }

    let targetLive = liveMap;
    if (!targetLive) liveSheetSpills.set(sheet.id, targetLive = new Map());
    for (let r = 0; r < rows; r++) for (let c = 0; c < columns; c++) {
      if (r === 0 && c === 0) continue;
      const key = cellKey({ row: anchor.row + r, column: anchor.column + c });
      clearedSet?.delete(key);
      targetLive.set(key, { ownerId: anchor.id, value: scalars[r * columns + c] ?? null });
    }
    liveSpills.set(anchor.id, { rows, columns });

    const anchorCache = cells[0] ?? { kind: 'number', value: 0 };
    return { cache: anchorCache, spill: { rows, columns, cells } };
  }

  const updates: CacheUpdate[] = [];
  for (const node of order) {
    if (selected && !selected.has(node)) continue;
    const { target, expression, diagnostic } = node;
    const sheet = inputs.get(target.sheetId)!.sheet;
    const cell = inputs.get(target.sheetId)!.cells.get(cellKey(target))!;
    clearAnchorSpill(target.sheetId, cell);

    let cache: CellCache | undefined;
    let spill: FormulaSpillResult | undefined;

    if (diagnostic === 'unsupported') {
      diagnostics.set(cell.id, 'Unsupported function or reference');
      cache = error('#NAME?');
    } else if (cycles.has(node)) {
      diagnostics.set(cell.id, 'Circular reference');
      cache = error('#REF!');
    } else if (!expression) {
      diagnostics.set(cell.id, 'Malformed formula');
      cache = error('#VALUE!');
    } else {
      const scalar = (value: FormulaValue): Scalar => {
        if (isOmitted(value)) return null;
        if (isLambda(value)) return error('#CALC!');
        if (isArray(value)) return value.at(0, 0);
        if (!isRange(value)) return value;
        if (value.length !== 1) return error('#VALUE!');
        const range = value[0]!;
        if (range.left === range.right && range.top === range.bottom) return read(range, range.top, range.left);
        if (range.left === range.right && cell.row >= range.top && cell.row <= range.bottom) return read(range, cell.row, range.left);
        if (range.top === range.bottom && cell.column >= range.left && cell.column <= range.right) return read(range, range.top, cell.column);
        return error('#VALUE!');
      };

      const implicitIntersect = (value: FormulaValue): FormulaValue => {
        if (isRange(value)) {
          if (value.length !== 1) return error('#VALUE!');
          const range = value[0]!;
          if (range.left === range.right && range.top === range.bottom) return read(range, range.top, range.left);
          if (range.left === range.right && cell.row >= range.top && cell.row <= range.bottom) return read(range, cell.row, range.left);
          if (range.top === range.bottom && cell.column >= range.left && cell.column <= range.right) return read(range, range.top, cell.column);
          return error('#VALUE!');
        }
        if (isArray(value)) return value.at(0, 0);
        if (isOmitted(value)) return null;
        if (isLambda(value)) return error('#VALUE!');
        return value;
      };

      function resolveR1C1(refText: string): readonly Range[] | undefined {
        const prefixMatch = /^(?:(?:'((?:[^']|'')+)'|([\p{L}_][\p{L}\p{N}_.]*))!)?(.+)$/u.exec(refText);
        if (!prefixMatch) return undefined;
        const sheetName = prefixMatch[1]?.replaceAll("''", "'") ?? prefixMatch[2];
        const targetSheet = sheetName === undefined ? sheet : sheetsByName.get(ordinalNameKey(sheetName));
        if (!targetSheet) return undefined;
        const body = prefixMatch[3]!.trim();
        const coord = (token: string, origin: number, limit: number): number | undefined => {
          if (token === '') return origin;
          if (token.startsWith('[') && token.endsWith(']')) {
            const delta = Number.parseInt(token.slice(1, -1), 10);
            if (!Number.isFinite(delta)) return undefined;
            const resolved = origin + delta;
            return resolved >= 1 && resolved <= limit ? resolved : undefined;
          }
          const absolute = Number.parseInt(token, 10);
          return Number.isFinite(absolute) && absolute >= 1 && absolute <= limit ? absolute : undefined;
        };
        const parseCell = (part: string): { row: number; column: number } | undefined => {
          const m = /^R(\[[+-]?\d+\]|[1-9]\d*)?C(\[[+-]?\d+\]|[1-9]\d*)?$/i.exec(part.trim());
          if (!m) return undefined;
          const r = coord(m[1] ?? '', target.row, documentLimits.rows);
          const c = coord(m[2] ?? '', target.column, documentLimits.columns);
          return r === undefined || c === undefined ? undefined : { row: r, column: c };
        };
        const parts = body.split(':').map(part => part.trim());
        if (parts.length === 1) {
          const single = parseCell(parts[0]!);
          if (single) return [{ sheetId: targetSheet.id, top: single.row, left: single.column, bottom: single.row, right: single.column }];
          const rowOnly = /^R(\[[+-]?\d+\]|[1-9]\d*)?$/i.exec(parts[0]!);
          if (rowOnly) {
            const r = coord(rowOnly[1] ?? '', target.row, documentLimits.rows);
            return r === undefined ? undefined : [{ sheetId: targetSheet.id, top: r, left: 1, bottom: r, right: documentLimits.columns }];
          }
          const colOnly = /^C(\[[+-]?\d+\]|[1-9]\d*)?$/i.exec(parts[0]!);
          if (colOnly) {
            const c = coord(colOnly[1] ?? '', target.column, documentLimits.columns);
            return c === undefined ? undefined : [{ sheetId: targetSheet.id, top: 1, left: c, bottom: documentLimits.rows, right: c }];
          }
          return undefined;
        }
        if (parts.length !== 2) return undefined;
        const leftCell = parseCell(parts[0]!);
        const rightCell = parseCell(parts[1]!);
        if (leftCell && rightCell) {
          return [{
            sheetId: targetSheet.id,
            top: Math.min(leftCell.row, rightCell.row),
            left: Math.min(leftCell.column, rightCell.column),
            bottom: Math.max(leftCell.row, rightCell.row),
            right: Math.max(leftCell.column, rightCell.column),
          }];
        }
        const rowA = /^R(\[[+-]?\d+\]|[1-9]\d*)?$/i.exec(parts[0]!);
        const rowB = /^R(\[[+-]?\d+\]|[1-9]\d*)?$/i.exec(parts[1]!);
        if (rowA && rowB) {
          const r1 = coord(rowA[1] ?? '', target.row, documentLimits.rows);
          const r2 = coord(rowB[1] ?? '', target.row, documentLimits.rows);
          if (r1 === undefined || r2 === undefined) return undefined;
          return [{ sheetId: targetSheet.id, top: Math.min(r1, r2), left: 1, bottom: Math.max(r1, r2), right: documentLimits.columns }];
        }
        const colA = /^C(\[[+-]?\d+\]|[1-9]\d*)?$/i.exec(parts[0]!);
        const colB = /^C(\[[+-]?\d+\]|[1-9]\d*)?$/i.exec(parts[1]!);
        if (colA && colB) {
          const c1 = coord(colA[1] ?? '', target.column, documentLimits.columns);
          const c2 = coord(colB[1] ?? '', target.column, documentLimits.columns);
          if (c1 === undefined || c2 === undefined) return undefined;
          return [{ sheetId: targetSheet.id, top: 1, left: Math.min(c1, c2), bottom: documentLimits.rows, right: Math.max(c1, c2) }];
        }
        return undefined;
      }

      function resolveRefText(refText: string, a1 = true): readonly Range[] | undefined {
        const trimmed = refText.trim();
        if (!trimmed) return undefined;
        if (!a1) return resolveR1C1(trimmed);
        let parsedRef: ReturnType<typeof parseFormula>;
        try {
          parsedRef = parseFormula('=' + trimmed);
        } catch {
          return undefined;
        }
        const expr = parsedRef.expression;
        if (expr.kind === 'reference' && expr.reference.endSheet === undefined) {
          const targetSheet = expr.reference.sheet === undefined ? sheet : sheetsByName.get(ordinalNameKey(expr.reference.sheet));
          const bounds = referenceBox(expr.reference);
          if (!targetSheet || bounds.bottom > documentLimits.rows || bounds.right > documentLimits.columns) return undefined;
          return [{ sheetId: targetSheet.id, ...bounds }];
        }
        if (expr.kind === 'name') {
          const area = resolveWorkbookName(docView, expr.name);
          if (!area || area.invalid) return undefined;
          return [{ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) }];
        }
        if (expr.kind === 'structured') {
          const area = resolveStructuredReference(docView, expr.reference, target);
          if (!area || area.invalid) return undefined;
          return [{ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) }];
        }
        if (expr.kind === 'spill' && expr.target.kind === 'reference' && expr.target.reference.endSheet === undefined) {
          const targetSheet = expr.target.reference.sheet === undefined ? sheet : sheetsByName.get(ordinalNameKey(expr.target.reference.sheet));
          const bounds = referenceBox(expr.target.reference);
          if (!targetSheet || bounds.top !== bounds.bottom || bounds.left !== bounds.right) return undefined;
          const anchorCell = cellAt(targetSheet.id, bounds.top, bounds.left);
          if (!anchorCell) return undefined;
          if (anchorCell.content.kind === 'formula') {
            const ext = liveSpills.get(anchorCell.id) ?? anchorCell.spillExtent;
            return [{
              sheetId: targetSheet.id,
              top: bounds.top,
              left: bounds.left,
              bottom: bounds.top + (ext?.rows ?? 1) - 1,
              right: bounds.left + (ext?.columns ?? 1) - 1,
            }];
          }
          if (anchorCell.content.kind === 'python' && anchorCell.content.block?.spillActive && anchorCell.content.block.lastTable) {
            const t = anchorCell.content.block.lastTable;
            return [{ sheetId: targetSheet.id, top: bounds.top, left: bounds.left, bottom: bounds.top + t.rows - 1, right: bounds.left + t.columns - 1 }];
          }
        }
        return undefined;
      }

      function evaluateInEnv(
        expression: FormulaExpression,
        env: ReadonlyMap<string, FormulaValue>,
        includeRanges = false,
        depth = 0,
      ): FormulaValue {
        const operators: Extract<FormulaExpression, { kind: 'binary' | 'unary' }>[] = [];
        while (expression.kind === 'binary' || expression.kind === 'unary') {
          operators.push(expression);
          expression = expression.kind === 'binary' ? expression.left : expression.value;
        }
        let value = evaluateTerm(expression, env, depth);
        while (operators.length) {
          const opExpr = operators.pop()!;
          if (opExpr.kind === 'unary' && opExpr.operator === '@') {
            value = implicitIntersect(value);
            continue;
          }
          if (opExpr.kind === 'binary' && opExpr.operator === ':') {
            if (!isRange(value)) {
              value = isError(value) ? value : error('#VALUE!');
              continue;
            }
            const right = evaluateInEnv(opExpr.right, env, true, depth);
            if (!isRange(right)) {
              value = isError(right) ? right : error('#VALUE!');
              continue;
            }
            if (value.length !== 1 || right.length !== 1 || value[0]!.sheetId !== right[0]!.sheetId) {
              value = error('#VALUE!');
              continue;
            }
            const a = value[0]!;
            const b = right[0]!;
            value = [{
              sheetId: a.sheetId,
              top: Math.min(a.top, b.top),
              left: Math.min(a.left, b.left),
              bottom: Math.max(a.bottom, b.bottom),
              right: Math.max(a.right, b.right),
            }];
            continue;
          }
          const leftArray = asArray(value, includeRanges);
          const left = leftArray ? null : scalar(value);
          if (isError(left)) return left;
          if (opExpr.kind === 'binary') {
            const right = evaluateInEnv(opExpr.right, env, includeRanges, depth);
            const rightArray = asArray(right, includeRanges);
            const apply = (a: Scalar, b: Scalar) => isError(a) ? a : isError(b) ? b :
              evaluateBinary(opExpr.operator, a, b, number, text);
            const rightScalar = rightArray ? null : scalar(right);
            value = leftArray || rightArray
              ? new OperatorArray(leftArray ?? left, rightArray ?? rightScalar, apply)
              : apply(left, rightScalar);
          } else {
            const apply = (input: Scalar) => {
              if (isError(input)) return input;
              const numeric = number(input);
              if (Number.isNaN(numeric)) return error('#VALUE!');
              return finite(opExpr.operator === '-' ? -numeric : opExpr.operator === '%' ? numeric / 100 : numeric);
            };
            value = leftArray ? new OperatorArray(leftArray, null, apply) : apply(left);
          }
        }
        return value;
      }

      function invokeLambda(lambda: FormulaLambda, argValues: readonly FormulaValue[], depth = 0): FormulaValue {
        if (depth >= 64) return error('#NUM!');
        if (argValues.length !== lambda.params.length) return error('#VALUE!');
        const callEnv = new Map(lambda.env);
        for (let i = 0; i < lambda.params.length; i++) {
          callEnv.set(lambda.params[i]!, argValues[i]!);
        }
        return evaluateInEnv(lambda.body, callEnv, true, depth + 1);
      }

      function evaluateTerm(
        expression: Exclude<FormulaExpression, { kind: 'unary' | 'binary' }>,
        env: ReadonlyMap<string, FormulaValue>,
        depth: number,
      ): FormulaValue {
        switch (expression.kind) {
          case 'missing': return omitted;
          case 'literal': return typeof expression.value === 'number' ? finite(expression.value) : expression.value;
          case 'error': return error(expression.value as Parameters<typeof error>[0]);
          case 'array': return { kind: 'array', rows: expression.rows.length, columns: expression.rows[0]!.length,
            at: (row, column) => scalar(evaluateTerm(expression.rows[row]![column]!, env, depth)) };
          case 'name': {
            const key = expression.name.toUpperCase();
            if (env.has(key)) return env.get(key)!;
            const nameDoc = { names } as WorkbookDocument;
            const area = resolveWorkbookName(nameDoc, expression.name);
            if (!area) return error('#NAME?');
            if (area.invalid) return error('#REF!');
            return [{ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) }];
          }
          case 'reference': {
            const range = resolvedReferences.get(expression.reference);
            return range ? [range] : error('#REF!');
          }
          case 'structured': {
            const area = resolveStructuredReference(docView, expression.reference, target);
            if (!area) {
              const hasTable = !expression.reference.table ||
                names.some(n => n.isTable && ordinalNameKey(n.name) === ordinalNameKey(expression.reference.table!));
              return error(hasTable ? '#REF!' : '#NAME?');
            }
            if (area.invalid) return error('#REF!');
            return [{ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) }];
          }
          case 'spill': {
            const inner = evaluateInEnv(expression.target, env, true, depth);
            if (!isRange(inner)) return isError(inner) ? inner : error('#VALUE!');
            if (inner.length !== 1) return error('#VALUE!');
            const anchorRange = inner[0]!;
            if (anchorRange.top !== anchorRange.bottom || anchorRange.left !== anchorRange.right) return error('#VALUE!');
            const anchorCell = inputs.get(anchorRange.sheetId)?.cells.get(cellKey({ row: anchorRange.top, column: anchorRange.left }));
            if (!anchorCell) return error('#REF!');
            if (anchorCell.content.kind === 'formula') {
              const live = liveSpills.get(anchorCell.id);
              const ext = live ?? anchorCell.spillExtent;
              if (!ext) return [anchorRange];
              return [{
                sheetId: anchorRange.sheetId,
                top: anchorRange.top,
                left: anchorRange.left,
                bottom: anchorRange.top + ext.rows - 1,
                right: anchorRange.left + ext.columns - 1,
              }];
            }
            if (anchorCell.content.kind === 'python' && anchorCell.content.block?.spillActive && anchorCell.content.block.lastTable) {
              const t = anchorCell.content.block.lastTable;
              return [{
                sheetId: anchorRange.sheetId,
                top: anchorRange.top,
                left: anchorRange.left,
                bottom: anchorRange.top + t.rows - 1,
                right: anchorRange.left + t.columns - 1,
              }];
            }
            return error('#REF!');
          }
          case 'invoke': {
            const callee = evaluateInEnv(expression.callee, env, true, depth);
            if (!isLambda(callee)) return isError(callee) ? callee : error('#VALUE!');
            const argValues = expression.args.map(arg => evaluateInEnv(arg, env, true, depth));
            return invokeLambda(callee, argValues, depth);
          }
          case 'union': {
            const ranges: Range[] = [];
            for (const part of expression.parts) {
              const value = evaluateInEnv(part, env, true, depth);
              if (!isRange(value)) return !isArray(value) && isError(value) ? value : error('#VALUE!');
              ranges.push(...value);
            }
            return ranges;
          }
          case 'call': {
            const bound = env.get(expression.name);
            if (bound !== undefined) {
              if (!isLambda(bound)) return error('#VALUE!');
              const argValues = expression.args.map(arg => evaluateInEnv(arg, env, true, depth));
              return invokeLambda(bound, argValues, depth);
            }
            return evaluateFunction(expression.name, expression.args, {
              evaluate: (expr, inc) => evaluateInEnv(expr, env, inc, depth),
              scalar,
              read,
              values,
              entries,
              number,
              text,
              locale,
              today,
              now,
              target,
              cellAt,
              resolveRefText,
              random,
              env,
              evaluateInEnv: (expr, nextEnv, inc) => evaluateInEnv(expr, nextEnv, inc, depth),
              invokeLambda: (fn, argVals) => invokeLambda(fn, argVals, depth),
              arrayOf,
              queryPivotData: (pivotCell, dataField, pairs) => {
                const pivot = findPivotContainingCell(docView, pivotCell);
                return pivot ? queryPivotData(docView, pivot, dataField, pairs, pivotCache) : error('#REF!');
              },
            });
          }
        }
      }

      const raw = evaluateInEnv(expression, emptyEnv, true, 0);
      if (isLambda(raw)) {
        cache = error('#CALC!');
      } else if (isArray(raw)) {
        if (raw.rows * raw.columns > 1) {
          if (previewMode) {
            const rLim = Math.min(raw.rows, 4), cLim = Math.min(raw.columns, 4);
            const cells: (CellCache | null)[] = [];
            for (let r = 0; r < rLim; r++) for (let c = 0; c < cLim; c++) {
              const val = raw.at(r, c);
              cells.push(val === null ? null : scalarCache(val));
            }
            cache = cells[0] ?? { kind: 'number', value: 0 };
            spill = { rows: raw.rows, columns: raw.columns, cells };
          } else {
            ({ cache, spill } = trySpillResult(sheet, cell, raw.rows, raw.columns, (r, c) => raw.at(r, c)));
          }
        } else {
          cache = scalarCache(raw.at(0, 0));
        }
      } else if (isRange(raw)) {
        if (raw.length !== 1) {
          cache = error('#VALUE!');
        } else {
          const area = raw[0]!;
          const rows = area.bottom - area.top + 1;
          const columns = area.right - area.left + 1;
          if (rows * columns > 1) {
            if (previewMode) {
              const rLim = Math.min(rows, 4), cLim = Math.min(columns, 4);
              const cells: (CellCache | null)[] = [];
              for (let r = 0; r < rLim; r++) for (let c = 0; c < cLim; c++) {
                const val = read(area, area.top + r, area.left + c);
                cells.push(val === null ? null : scalarCache(val));
              }
              cache = cells[0] ?? { kind: 'number', value: 0 };
              spill = { rows, columns, cells };
            } else {
              ({ cache, spill } = trySpillResult(sheet, cell, rows, columns, (r, c) => read(area, area.top + r, area.left + c)));
            }
          } else {
            cache = scalarCache(read(area, area.top, area.left));
          }
        }
      } else {
        cache = scalarCache(scalar(raw));
      }
    }

    const finalCache = cell.requiresExternalCalculation && cell.cache ? cell.cache : cache;
    evaluated.set(cell, finalCache.kind === 'error' ? finalCache : finalCache.value);
    updates.push({
      target: node.target,
      cache: finalCache,
      ...(spill && { spill }),
    });
  }
  return { updates, diagnostics, dependencies };
}

type ArrayOperand = FormulaArray | Scalar;

function arrayElement(value: ArrayOperand, row: number, column: number): Scalar {
  if (!isArray(value)) return value;
  const r = value.rows === 1 ? 0 : row, c = value.columns === 1 ? 0 : column;
  return r < value.rows && c < value.columns ? value.at(r, c) : error('#N/A');
}

// Walk left-associated chains iteratively and compute just the requested array element.
class OperatorArray implements FormulaArray {
  readonly kind = 'array' as const;
  readonly rows: number;
  readonly columns: number;

  constructor(private readonly left: ArrayOperand, private readonly right: ArrayOperand,
    private readonly apply: (left: Scalar, right: Scalar) => Scalar) {
    this.rows = Math.max(isArray(left) ? left.rows : 1, isArray(right) ? right.rows : 1);
    this.columns = Math.max(isArray(left) ? left.columns : 1, isArray(right) ? right.columns : 1);
  }

  at(row: number, column: number): Scalar {
    const pending: { array: OperatorArray; row: number; column: number }[] = [];
    let value: ArrayOperand = this;
    while (value instanceof OperatorArray) {
      if (value.rows === 1) row = 0;
      if (value.columns === 1) column = 0;
      if (row >= value.rows || column >= value.columns) break;
      pending.push({ array: value, row, column });
      value = value.left;
    }
    let result = arrayElement(value, row, column);
    while (pending.length) {
      const { array, row, column } = pending.pop()!;
      result = array.apply(result, arrayElement(array.right, row, column));
    }
    return result;
  }
}

function evaluateBinary(operator: string, left: Exclude<Scalar, FormulaError>, right: Exclude<Scalar, FormulaError>,
  number: (value: Scalar) => number, text: (value: Exclude<Scalar, FormulaError>) => string): Scalar {
  if (operator === '&') {
    const joined = text(left) + text(right);
    return joined.length <= documentLimits.textLength ? joined : error('#VALUE!');
  }
  if (['=', '<>', '<', '>', '<=', '>='].includes(operator)) return matchesComparison(compare(left, right), operator);
  const a = number(left), b = number(right);
  if (Number.isNaN(a) || Number.isNaN(b)) return error('#VALUE!');
  switch (operator) {
    case '+': return finite(a + b);
    case '-': return finite(a - b);
    case '*': return finite(a * b);
    case '/': return b === 0 ? error('#DIV/0!') : finite(a / b);
    case '^': return finite(a ** b);
    default: return error('#VALUE!');
  }
}

function formatPreviewCache(cache: CellCache | null | undefined, quoteStrings: boolean): { readonly text: string; readonly isError: boolean } {
  if (!cache) return { text: quoteStrings ? '""' : '', isError: false };
  if (cache.kind === 'error') return { text: cache.value, isError: true };
  if (cache.kind === 'boolean') return { text: cache.value ? 'TRUE' : 'FALSE', isError: false };
  if (cache.kind === 'text') return { text: quoteStrings ? `"${cache.value}"` : cache.value, isError: false };
  return { text: String(cache.value), isError: false };
}

export function previewFormula(
  document: WorkbookDocument,
  target: CellReference,
  source: string,
  locale?: string,
  quoteStrings = true,
): { readonly text: string; readonly isError: boolean } | undefined {
  if (!source.startsWith('=')) return undefined;
  const targetSheet = document.sheets.find(s => s.id === target.sheetId);
  if (!targetSheet) return undefined;
  let parsed: ReturnType<typeof parseFormula>;
  try {
    parsed = parseFormula(source);
  } catch {
    return undefined;
  }

  const sheets = new Map(document.sheets.map(sheet => [ordinalNameKey(sheet.name), sheet]));
  const resolvedReferences = new Map<FormulaReference, Range>();
  for (const ref of parsed.references) {
    if (ref.endSheet !== undefined) return { text: '#VALUE!', isError: true };
    const refSheet = ref.sheet === undefined ? targetSheet : sheets.get(ordinalNameKey(ref.sheet));
    const bounds = referenceBox(ref);
    if (!refSheet || bounds.bottom > documentLimits.rows || bounds.right > documentLimits.columns) continue;
    resolvedReferences.set(ref, { sheetId: refSheet.id, ...bounds });
  }

  const previewCell: CellDocument = {
    id: '__preview__',
    row: target.row,
    column: target.column,
    content: { kind: 'formula', source },
  };
  const inputs = indexNativeInputs(document);
  const sheetInputs = inputs.get(targetSheet.id);
  if (!sheetInputs) return undefined;
  sheetInputs.cells.set(cellKey(previewCell), previewCell);

  const node: FormulaNode = {
    id: previewCell.id,
    target,
    expression: parsed.expression,
    ranges: [...resolvedReferences.values()],
    inputs: [],
    dependents: [],
    containsSubtotal: false,
    usesGetPivotData: false,
    volatile: false,
    hasSpillRef: false,
  };
  const graph = {
    nodes: new Map([[node.id, node]]),
    resolvedReferences,
    dependencies: new Map<string, readonly CellReference[]>(),
    names: document.names,
    pivots: document.pivots ?? [],
    order: [node],
    cycles: new Set<FormulaNode>(),
  };
  const result = evaluateNative(inputs, graph, undefined, locale !== undefined ? { locale } : {}, true);
  const update = result.updates[0];
  if (!update) return undefined;
  if (update.spill) {
    const { cells, rows, columns } = update.spill;
    const rLim = Math.min(rows, 4), cLim = Math.min(columns, 4);
    const flatLimit = 6;
    let count = 0;
    const rowStrings: string[] = [];
    for (let r = 0; r < rLim && count < flatLimit; r++) {
      const cols: string[] = [];
      for (let c = 0; c < cLim && count < flatLimit; c++) {
        cols.push(formatPreviewCache(cells[r * cLim + c], true).text);
        count++;
      }
      rowStrings.push(cols.join(', '));
    }
    const truncated = rows * columns > count ? '; …' : '';
    return { text: `{${rowStrings.join('; ')}${truncated}}`, isError: false };
  }
  return formatPreviewCache(update.cache, quoteStrings);
}
