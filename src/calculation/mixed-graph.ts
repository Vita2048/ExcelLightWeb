import { cellKey, documentLimits, type CellDocument, type CellExtent, type CellReference, type WorkbookDocument } from '../core/workbook';
import { equalDto } from '../core/workbook-delta';
import { pythonInputRanges } from '../python/python-inputs';
import { prepareNativeGraph } from './native-calculator';

export interface MixedRange {
  readonly sheetId: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export interface MixedNode {
  readonly id: string;
  readonly target: CellReference;
  readonly kind: 'formula' | 'python';
  readonly cell: CellDocument;
  readonly ranges: readonly MixedRange[];
  readonly dependencies: Set<string>;
  readonly dependents: Set<string>;
  error?: string;
}

function extentRange(extent: CellExtent & { readonly sheetId: string }): MixedRange {
  return { sheetId: extent.sheetId, top: extent.row, left: extent.column,
    bottom: extent.row + extent.rows - 1, right: extent.column + extent.columns - 1 };
}

function contains(range: MixedRange, target: CellReference) {
  return range.sheetId === target.sheetId && target.row >= range.top && target.row <= range.bottom &&
    target.column >= range.left && target.column <= range.right;
}

function sortedIds(ids: Iterable<string>, nodes: ReadonlyMap<string, MixedNode>) {
  return [...ids].sort((left, right) => {
    const a = nodes.get(left)!.target, b = nodes.get(right)!.target;
    return a.sheetId.localeCompare(b.sheetId) || cellKey(a) - cellKey(b);
  });
}

/** Formula and Python inputs share one graph; spill cells point back to their owner. */
export class MixedGraph {
  readonly nodes = new Map<string, MixedNode>();
  readonly order: readonly MixedNode[];
  readonly blocked: ReadonlySet<string>;

  constructor(readonly source: WorkbookDocument) {
    const native = prepareNativeGraph(source);
    const sheets = new Map(source.sheets.map(sheet => [sheet.id, sheet]));
    const owners = new Map<string, string>();
    const cellsBySheet = new Map<string, CellDocument[]>();
    for (const sheet of source.sheets) {
      cellsBySheet.set(sheet.id, [...sheet.cells].sort((a, b) => cellKey(a) - cellKey(b)));
      for (const cell of sheet.cells) if (cell.content.kind === 'python') owners.set(cell.content.blockId, cell.id);
    }

    for (const sheet of source.sheets) for (const cell of sheet.cells) {
      if (cell.content.kind !== 'formula' && cell.content.kind !== 'python') continue;
      const target = { sheetId: sheet.id, row: cell.row, column: cell.column };
      const ranges: MixedRange[] = [];
      let error: string | undefined;
      if (cell.content.kind === 'formula') {
        const formula = native.nodes.get(cell.id);
        if (formula) ranges.push(...formula.ranges);
      } else {
        try {
          ranges.push(...pythonInputRanges(source, sheet.id, cell.content.source, cell.content.block?.bindings)
            .map(extentRange));
        } catch (exception) {
          error = exception instanceof Error ? exception.message : String(exception);
        }
      }
      this.nodes.set(cell.id, { id: cell.id, target, kind: cell.content.kind, cell, ranges,
        dependencies: new Set(), dependents: new Set(), ...(error && { error }) });
    }

    for (const node of this.nodes.values()) for (const range of node.ranges) {
      const sheet = sheets.get(range.sheetId);
      if (!sheet) continue;
      for (const cell of cellsBySheet.get(range.sheetId) ?? []) {
        if (cell.row < range.top || cell.row > range.bottom || cell.column < range.left || cell.column > range.right) continue;
        const inputId = cell.spillOwner ? owners.get(cell.spillOwner) : cell.id;
        if (inputId && inputId !== node.id && this.nodes.has(inputId)) node.dependencies.add(inputId);
      }
    }
    for (const node of this.nodes.values()) for (const dependency of node.dependencies) {
      this.nodes.get(dependency)!.dependents.add(node.id);
    }

    const remaining = new Map([...this.nodes].map(([id, node]) => [id, node.dependencies.size]));
    const ready = sortedIds([...remaining].filter(([, count]) => count === 0).map(([id]) => id), this.nodes);
    const order: MixedNode[] = [];
    while (ready.length) {
      const id = ready.shift()!;
      order.push(this.nodes.get(id)!);
      for (const dependent of this.nodes.get(id)!.dependents) {
        const next = remaining.get(dependent)! - 1;
        remaining.set(dependent, next);
        if (next === 0) {
          ready.push(dependent);
          ready.sort((left, right) => {
            const a = this.nodes.get(left)!.target, b = this.nodes.get(right)!.target;
            return a.sheetId.localeCompare(b.sheetId) || cellKey(a) - cellKey(b);
          });
        }
      }
    }
    this.order = order;
    this.blocked = new Set([...remaining].filter(([, count]) => count > 0).map(([id]) => id));
  }

  closure(roots: Iterable<string>): Set<string> {
    const found = new Set([...roots].filter(id => this.nodes.has(id)));
    const pending = [...found];
    while (pending.length) for (const dependency of this.nodes.get(pending.pop()!)!.dependencies) {
      if (!found.has(dependency)) { found.add(dependency); pending.push(dependency); }
    }
    return found;
  }

  dependents(roots: Iterable<string>): Set<string> {
    const found = new Set([...roots].filter(id => this.nodes.has(id)));
    const pending = [...found];
    while (pending.length) for (const dependent of this.nodes.get(pending.pop()!)!.dependents) {
      if (!found.has(dependent)) { found.add(dependent); pending.push(dependent); }
    }
    return found;
  }

  nativePrerequisites(root: string): Set<string> {
    const formulas = new Set<string>();
    const visited = new Set<string>([root]);
    const pending = [root];
    while (pending.length) {
      const current = this.nodes.get(pending.pop()!);
      if (!current) continue;
      for (const id of current.dependencies) {
        if (visited.has(id)) continue;
        visited.add(id);
        const dependency = this.nodes.get(id)!;
        if (dependency.kind === 'python') continue;
        formulas.add(id);
        pending.push(id);
      }
    }
    return formulas;
  }

  topologyMatches(other: MixedGraph): boolean {
    if (this.nodes.size !== other.nodes.size) return false;
    for (const [id, node] of this.nodes) {
      const next = other.nodes.get(id);
      if (!next || node.kind !== next.kind || node.dependencies.size !== next.dependencies.size ||
          [...node.dependencies].some(dependency => !next.dependencies.has(dependency))) return false;
    }
    return true;
  }

  affectedPython(before: WorkbookDocument, after: WorkbookDocument): Set<string> {
    const allPython = () => new Set([...this.nodes.values()].filter(node => node.kind === 'python').map(node => node.id));
    if (!equalDto(before.names, after.names) || before.sheets.length !== after.sheets.length ||
        before.sheets.some((sheet, index) => sheet.id !== after.sheets[index]?.id || sheet.name !== after.sheets[index]?.name)) {
      return allPython();
    }
    const changed = new Map<string, Map<number, CellReference>>();
    for (const [index, sheet] of before.sheets.entries()) {
      const next = after.sheets[index]!;
      const left = new Map(sheet.cells.map(cell => [cellKey(cell), cell]));
      const right = new Map(next.cells.map(cell => [cellKey(cell), cell]));
      for (const key of new Set([...left.keys(), ...right.keys()])) {
        const oldCell = left.get(key), newCell = right.get(key);
        const oldPython = oldCell?.content.kind === 'python' ? oldCell.content : undefined;
        const newPython = newCell?.content.kind === 'python' ? newCell.content : undefined;
        const samePythonSource = oldPython && newPython && oldPython.blockId === newPython.blockId &&
          oldPython.source === newPython.source && equalDto(oldPython.block?.bindings, newPython.block?.bindings);
        const changedInput = !samePythonSource && !equalDto(oldCell?.content, newCell?.content) ||
          oldCell?.style?.numberFormat !== newCell?.style?.numberFormat || oldCell?.spillOwner !== newCell?.spillOwner;
        if (changedInput) {
          let points = changed.get(sheet.id);
          if (!points) changed.set(sheet.id, points = new Map());
          points.set(key, { sheetId: sheet.id, row: (key / documentLimits.columns | 0) + 1,
            column: key % documentLimits.columns + 1 });
        }
      }
      if (!equalDto(sheet.rows, next.rows) || !equalDto(sheet.filter, next.filter)) {
        for (const node of this.nodes.values()) if (node.kind === 'formula' &&
            (node.cell.content.kind === 'formula' && /\bSUBTOTAL\s*\(/iu.test(node.cell.content.source))) {
          let points = changed.get(sheet.id);
          if (!points) changed.set(sheet.id, points = new Map());
          points.set(cellKey(node.target), node.target);
        }
      }
    }
    const roots = new Set<string>();
    for (const node of this.nodes.values()) {
      const oldCell = before.sheets.flatMap(sheet => sheet.cells).find(cell => cell.id === node.id);
      if (!oldCell || !equalDto(oldCell.content, node.cell.content) ||
          oldCell.style?.numberFormat !== node.cell.style?.numberFormat) {
        roots.add(node.id);
        continue;
      }
      if (node.ranges.some(range => [...(changed.get(range.sheetId)?.values() ?? [])].some(target => contains(range, target)))) roots.add(node.id);
    }
    return new Set([...this.dependents(roots)].filter(id => this.nodes.get(id)?.kind === 'python'));
  }
}

