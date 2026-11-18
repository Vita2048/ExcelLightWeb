import type { CacheUpdate, NativeCalculationScope } from '../application/workbook-protocol';
import { equalDto } from '../core/workbook-delta';
import { cellKey, type CellCache, type Coordinate, type WorkbookDocument } from '../core/workbook';
import { cellsIn, evaluateNative, indexNativeInputs, prepareNativeGraph, type CalculationOptions } from './native-calculator';

type Graph = ReturnType<typeof prepareNativeGraph>;
type Node = Graph['order'][number];
type Inputs = ReturnType<typeof indexNativeInputs>;
interface Generation {
  readonly inputs: Inputs;
  readonly graph: Graph;
  readonly stale: ReadonlySet<Node>;
  readonly caches: ReadonlyMap<string, CellCache>;
}

function expand(nodes: Set<Node>, adjacent: 'inputs' | 'dependents') {
  // Set iteration visits each newly added node, including through cycles.
  for (const node of nodes) for (const next of node[adjacent]) nodes.add(next);
}

function hiddenRows(rows: WorkbookDocument['sheets'][number]['rows']) {
  return rows.filter(row => row.hidden).map(row => row.index).sort((a, b) => a - b);
}

export function runNativeGeneration(document: WorkbookDocument, scope: NativeCalculationScope,
  previous?: Generation, options: CalculationOptions = {}) {
  const inputs = indexNativeInputs(document);
  let topologyChanged = !previous || inputs.size !== previous.inputs.size;
  const changed = new Map<string, Coordinate[]>();
  const visibilityChanged = new Set<string>();
  for (const [sheetId, input] of inputs) {
    const prior = previous?.inputs.get(sheetId);
    if (!prior || prior.sheet.name !== input.sheet.name) topologyChanged = true;
    if (!prior) continue;
    const changes: Coordinate[] = [];
    if (!equalDto(prior.sheet.filter, input.sheet.filter) || !equalDto(hiddenRows(prior.sheet.rows), hiddenRows(input.sheet.rows))) {
      visibilityChanged.add(sheetId);
    }
    const filter = input.sheet.filter;
    const criteria = new Set(filter?.criteria.map(rule => rule.column));
    for (const [key, cell] of input.cells) {
      const before = prior.cells.get(key);
      const formula = cell.content.kind === 'formula' || before?.content.kind === 'formula';
      const contentChanged = !equalDto(before?.content, cell.content) ||
        !!before?.requiresExternalCalculation !== !!cell.requiresExternalCalculation;
      if (formula && (contentChanged || before?.id !== cell.id)) topologyChanged = true;
      const expected = before?.content.kind === 'formula' ? previous!.caches.get(before.id) : before?.cache;
      const cacheChanged = !equalDto(expected, cell.cache);
      if (contentChanged || cacheChanged) changes.push(cell);
      // Filter display uses the cache at the start of the generation.
      if (filter && criteria.has(cell.column) && cell.row > filter.row && cell.row < filter.row + filter.rows &&
          (contentChanged || !equalDto(before?.cache, cell.cache))) visibilityChanged.add(sheetId);
    }
    for (const [key, before] of prior.cells) {
      if (input.cells.has(key)) continue;
      changes.push(before);
      if (before.content.kind === 'formula') topologyChanged = true;
      if (filter && criteria.has(before.column) && before.row > filter.row && before.row < filter.row + filter.rows) {
        visibilityChanged.add(sheetId);
      }
    }
    if (changes.length) changed.set(sheetId, changes.sort((a, b) => cellKey(a) - cellKey(b)));
  }

  const graph = topologyChanged ? prepareNativeGraph(document) : previous!.graph;
  const stale = new Set<Node>();
  for (const node of graph.order) {
    const { target, id } = node;
    const priorNode = previous?.graph.nodes.get(id);
    const ownCell = inputs.get(target.sheetId)!.cells.get(cellKey(target))!;
    if (!priorNode || !equalDto(priorNode.target, target) ||
        !equalDto(previous!.inputs.get(target.sheetId)?.cells.get(cellKey(target))?.content, ownCell.content) ||
        !equalDto(priorNode.ranges, node.ranges) || priorNode.diagnostic !== node.diagnostic ||
        previous!.graph.cycles.has(priorNode) !== graph.cycles.has(node) ||
        previous!.stale.has(priorNode) || !ownCell.cache || !equalDto(previous!.caches.get(id), ownCell.cache) ||
        node.ranges.some(range => (node.containsSubtotal && visibilityChanged.has(range.sheetId)) ||
          !cellsIn(changed.get(range.sheetId) ?? [], range).next().done)) {
      stale.add(node);
    }
  }
  expand(stale, 'dependents');
  const selected = new Set<Node>();
  if (scope === 'automatic') {
    for (const node of graph.order) if (stale.has(node) || node.volatile) selected.add(node);
    expand(selected, 'dependents');
  } else if (scope !== 'observe') {
    for (const node of graph.order) if (scope === 'workbook' || node.target.sheetId === scope.sheetId) selected.add(node);
    expand(selected, 'inputs');
  }

  const updates: readonly CacheUpdate[] = selected.size ? evaluateNative(inputs, graph, selected, options).updates : [];
  const caches = new Map<string, CellCache>();
  for (const node of graph.order) {
    const cache = inputs.get(node.target.sheetId)!.cells.get(cellKey(node.target))!.cache;
    if (cache) caches.set(node.id, cache);
  }
  // A scoped calculation can change an earlier clean precedent (for example TODAY).
  // Reflag downstream cells outside the selected closure if its visible result changes.
  for (const update of updates) {
    const id = inputs.get(update.target.sheetId)!.cells.get(cellKey(update.target))!.id;
    if (!equalDto(caches.get(id), update.cache)) stale.add(graph.nodes.get(id)!);
    if (update.cache) caches.set(id, update.cache);
  }
  expand(stale, 'dependents');
  for (const node of selected) stale.delete(node);
  const generation: Generation = { inputs, graph, stale, caches };
  return { generation, updates, staleCellIds: [...stale].map(node => node.id) };
}
