import { formatCellValue } from '../core/formats/cell-format';
import type { CacheUpdate } from '../application/workbook-protocol';
import { parseFormula, referenceBox, UnsupportedFormulaReference, type FormulaExpression, type FormulaReference } from '../core/formula-references';
import { maskRows, RowMask } from '../core/row-mask';
import { cellKey, documentLimits, resolvedCellStyle, type CellDocument, type CellReference, type SheetDocument, type WorkbookDocument } from '../core/workbook';
import { lastColumn, lastRow, resolveWorkbookName } from '../core/workbook-names';
import { ordinalNameKey } from '../core/workbook-validation';
import { todaySerial } from './formula-dates';
import { evaluateFunction, sumIfRange, supportedFunctions } from './formula-functions';
import { compare, error, finite, formulaNodes, isArray, isError, isRange, matchesComparison, numberValue, scalarCache, textValue,
  type FormulaArray, type FormulaError, type FormulaValue, type Range, type Scalar } from './formula-values';

interface FormulaNode {
  readonly id: string;
  readonly target: CellReference;
  readonly expression?: FormulaExpression;
  readonly ranges: readonly Range[];
  readonly inputs: FormulaNode[];
  readonly dependents: FormulaNode[];
  readonly containsSubtotal: boolean;
  readonly volatile: boolean;
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
  if (!cell || cell.content.kind === 'blank') return null;
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
}

export function prepareNativeGraph(document: WorkbookDocument) {
  const sheets = new Map(document.sheets.map(sheet => [ordinalNameKey(sheet.name), sheet]));
  const formulas = new Map(document.sheets.map(sheet => [sheet.id, sheet.cells.filter(cell => cell.content.kind === 'formula')
    .sort((a, b) => cellKey(a) - cellKey(b))]));
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
      nodes.set(cell.id, { ...base, ranges: [], containsSubtotal: false, volatile: false,
        diagnostic: cell.requiresExternalCalculation || exception instanceof UnsupportedFormulaReference ? 'unsupported' : 'malformed' });
      continue;
    }
    let containsSubtotal = false;
    let volatile = false;
    let unsupported = !!cell.requiresExternalCalculation;
    for (const node of formulaNodes(parsed.expression)) {
      if (node.kind === 'name') {
        const area = resolveWorkbookName(document, node.name);
        if (!area) unsupported = true;
      }
      if (node.kind === 'call') {
        if (!supportedFunctions.has(node.name)) unsupported = true;
        if (node.name === 'SUBTOTAL') containsSubtotal = true;
        if (node.name === 'TODAY') volatile = true;
      }
    }
    const ranges = parsed.references.flatMap(ref => {
      const range = resolve(sheet, ref);
      if (!range) return [];
      resolvedReferences.set(ref, range);
      return [range];
    });
    for (const node of formulaNodes(parsed.expression)) {
      if (node.kind === 'name') {
        const area = resolveWorkbookName(document, node.name);
        if (area && !area.invalid) {
          ranges.push({ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) });
        }
      }
    }
    function candidateRanges(expression: FormulaExpression | undefined): Range[] {
      if (expression?.kind === 'reference') {
        const range = resolvedReferences.get(expression.reference);
        return range ? [range] : [];
      }
      if (expression?.kind === 'name') {
        const area = resolveWorkbookName(document, expression.name);
        return area && !area.invalid ? [{ sheetId: area.sheetId, top: area.row, left: area.column, bottom: lastRow(area), right: lastColumn(area) }] : [];
      }
      if (expression?.kind === 'call' && expression.name === 'IF') return expression.args.slice(1).flatMap(candidateRanges);
      return [];
    }
    for (const expression of formulaNodes(parsed.expression)) {
      if (expression.kind !== 'call' || expression.name !== 'SUMIF' || expression.args.length !== 3) continue;
      const [criteria, , sum] = expression.args;
      for (const source of candidateRanges(criteria)) for (const target of candidateRanges(sum)) {
        const expanded = sumIfRange(source, target);
        if (expanded) ranges.push(expanded);
      }
    }
    nodes.set(cell.id, { ...base, expression: parsed.expression, ranges, containsSubtotal, volatile,
      ...(unsupported && { diagnostic: 'unsupported' as const }) });
  }
  const dependencies = new Map<string, readonly CellReference[]>();
  for (const [id, node] of nodes) {
    const inputs = new Set<FormulaNode>();
    for (const range of node.ranges) {
      for (const other of cellsIn(formulas.get(range.sheetId)!, range)) inputs.add(nodes.get(other.id)!);
    }
    const list = [...inputs];
    node.inputs.push(...list);
    for (const input of list) input.dependents.push(node);
    dependencies.set(id, list.map(input => input.target));
  }
  return { nodes, resolvedReferences, dependencies, names: document.names, ...evaluationOrder(nodes.values()) };
}

export function indexNativeInputs(document: WorkbookDocument) {
  return new Map(document.sheets.map(sheet => [sheet.id, {
    sheet, cells: new Map(sheet.cells.map(cell => [cellKey(cell), cell])),
  }]));
}

export function calculateNative(document: WorkbookDocument, options: CalculationOptions = {}): NativeCalculation {
  return evaluateNative(indexNativeInputs(document), prepareNativeGraph(document), undefined, options);
}

export function evaluateNative(inputs: ReturnType<typeof indexNativeInputs>, graph: ReturnType<typeof prepareNativeGraph>,
  selected: ReadonlySet<FormulaNode> | undefined, options: CalculationOptions): NativeCalculation {
  const today = todaySerial(options.clock?.() ?? new Date());
  const numberFormat = new Intl.NumberFormat(options.locale);
  const locale = numberFormat.resolvedOptions().locale;
  const parts = numberFormat.formatToParts(12345.6);
  const decimal = parts.find(part => part.type === 'decimal')?.value;
  const group = parts.find(part => part.type === 'group')?.value;
  const number = (value: Scalar) => numberValue(value, decimal, group);
  const text = (value: Exclude<Scalar, FormulaError>) => textValue(value, decimal);
  const { nodes, resolvedReferences, order, cycles, dependencies, names } = graph;
  const sorted = new Map<string, CellDocument[]>();
  const evaluated = new Map<CellDocument, Scalar>();
  const diagnostics = new Map<string, string>();
  const filtered = new Map<SheetDocument, RowMask>();
  const manual = new Map<SheetDocument, ReadonlySet<number>>();

  function computedValue(cell: CellDocument | undefined): Scalar {
    if (cell?.requiresExternalCalculation) return error('#NAME?');
    return (cell && evaluated.get(cell)) ?? baseValue(cell);
  }
  function read(range: Range, row: number, column: number): Scalar {
    return computedValue(inputs.get(range.sheetId)!.cells.get(cellKey({ row, column })));
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
    for (const cell of rangeCells(range)) {
      yield { row: cell.row, column: cell.column, value: computedValue(cell) };
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
    for (const cell of rangeCells(range)) {
      if (subtotal && (excluded?.has(cell.row) || hidden?.has(cell.row) || nodes.get(cell.id)?.containsSubtotal)) continue;
      yield computedValue(cell);
    }
  }

  const updates: CacheUpdate[] = [];
  for (const node of order) {
    if (selected && !selected.has(node)) continue;
    const { target, expression, diagnostic } = node;
    const cell = inputs.get(target.sheetId)!.cells.get(cellKey(target))!;
    let value: Scalar;
    if (diagnostic === 'unsupported') {
      diagnostics.set(cell.id, 'Unsupported function or reference');
      value = error('#NAME?');
    } else if (cycles.has(node)) {
      diagnostics.set(cell.id, 'Circular reference');
      value = error('#REF!');
    } else if (!expression) {
      diagnostics.set(cell.id, 'Malformed formula');
      value = error('#VALUE!');
    } else {
      const scalar = (value: FormulaValue): Scalar => {
        if (isArray(value)) return value.at(0, 0);
        if (!isRange(value)) return value;
        if (value.length !== 1) return error('#VALUE!');
        const range = value[0]!;
        if (range.left === range.right && range.top === range.bottom) return read(range, range.top, range.left);
        if (range.left === range.right && cell.row >= range.top && cell.row <= range.bottom) return read(range, cell.row, range.left);
        if (range.top === range.bottom && cell.column >= range.left && cell.column <= range.right) return read(range, range.top, cell.column);
        return error('#VALUE!');
      };
      // Flat operator chains can fill the source-length limit without nesting parentheses.
      function evaluate(expression: FormulaExpression, includeRanges = false): FormulaValue {
        const operators: Extract<FormulaExpression, { kind: 'binary' | 'unary' }>[] = [];
        while (expression.kind === 'binary' || expression.kind === 'unary') {
          operators.push(expression);
          expression = expression.kind === 'binary' ? expression.left : expression.value;
        }
        let value = evaluateTerm(expression);
        while (operators.length) {
          const expression = operators.pop()!;
          const leftArray = asArray(value, includeRanges);
          const left = leftArray ? null : scalar(value);
          if (isError(left)) return left;
          if (expression.kind === 'binary') {
            const right = evaluate(expression.right, includeRanges);
            const rightArray = asArray(right, includeRanges);
            const apply = (a: Scalar, b: Scalar) => isError(a) ? a : isError(b) ? b :
              evaluateBinary(expression.operator, a, b, number, text);
            const rightScalar = rightArray ? null : scalar(right);
            value = leftArray || rightArray
              ? new OperatorArray(leftArray ?? left, rightArray ?? rightScalar, apply)
              : apply(left, rightScalar);
          } else {
            const apply = (input: Scalar) => {
              if (isError(input)) return input;
              const numeric = number(input);
              if (Number.isNaN(numeric)) return error('#VALUE!');
              return finite(expression.operator === '-' ? -numeric : expression.operator === '%' ? numeric / 100 : numeric);
            };
            value = leftArray ? new OperatorArray(leftArray, null, apply) : apply(left);
          }
        }
        return value;
      }
      function evaluateTerm(expression: Exclude<FormulaExpression, { kind: 'unary' | 'binary' }>): FormulaValue {
        switch (expression.kind) {
          case 'missing': return null;
          case 'literal': return typeof expression.value === 'number' ? finite(expression.value) : expression.value;
          case 'error': return error(expression.value as Parameters<typeof error>[0]);
          case 'array': return { kind: 'array', rows: expression.rows.length, columns: expression.rows[0]!.length,
            at: (row, column) => scalar(evaluateTerm(expression.rows[row]![column]!)) };
          case 'name': {
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
          case 'union': {
            const ranges: Range[] = [];
            for (const part of expression.parts) {
              const value = evaluate(part);
              if (!isRange(value)) return !isArray(value) && isError(value) ? value : error('#VALUE!');
              ranges.push(...value);
            }
            return ranges;
          }
          case 'call': return evaluateFunction(expression.name, expression.args, { evaluate, scalar, read, values, entries, number, text, locale, today });
        }
      }
      value = scalar(evaluate(expression, true));
    }
    const cache = scalarCache(value);
    evaluated.set(cell, cache.kind === 'error' ? cache : cache.value);
    updates.push({ target: node.target, cache: cell.requiresExternalCalculation && cell.cache ? cell.cache : cache });
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
