export const documentLimits = Object.freeze({
  sheets: 10, rows: 10_000, columns: 100, cells: 200_000, formulas: 10_000,
  formulaLength: 8192, textLength: 32767, sheetNameLength: 31, numberFormatLength: 128,
  names: 1000, pythonSourceBytes: 1024 * 1024,
});
export const worksheetViewport = Object.freeze({ rows: 1000, columns: 26, address: 'A1:Z1000' });
/** Source given to a new =PY anchor, matching the desktop's working starter script. */
export const defaultPythonSource = '42\n';

export type DocumentId = string;
export type IdFactory = () => DocumentId;
export class DocumentError extends Error {}

export interface Coordinate {
  readonly row: number;
  readonly column: number;
}
export interface CellReference extends Coordinate {
  readonly sheetId: DocumentId;
  readonly absoluteRow?: boolean;
  readonly absoluteColumn?: boolean;
  readonly invalid?: boolean;
}
export interface CellExtent extends Coordinate {
  readonly rows: number;
  readonly columns: number;
}
export interface AxisDimension {
  readonly index: number;
  readonly size?: number;
  readonly hidden?: boolean;
}
export interface FrozenPanes {
  readonly rows: number;
  readonly columns: number;
}
export interface FilterCriterion {
  readonly column: number;
  readonly allowed: readonly string[];
}
export interface SheetFilter extends CellExtent {
  readonly criteria: readonly FilterCriterion[];
  readonly sort?: { readonly column: number; readonly ascending: boolean };
}

export interface TableDesign {
  readonly style: 'Neutral' | 'Blue' | 'Green' | 'Orange' | 'Purple';
  readonly headerRow: boolean;
  readonly totalRow: boolean;
  readonly firstColumn?: boolean;
  readonly lastColumn?: boolean;
  readonly bandedRows?: boolean;
  readonly bandedColumns?: boolean;
  readonly filterButton?: boolean;
}
export interface WorkbookName {
  readonly id: DocumentId;
  readonly name: string;
  readonly area: CellExtent & { readonly sheetId: DocumentId; readonly invalid?: boolean };
  readonly comment?: string;
  readonly isTable?: boolean;
  readonly design?: TableDesign;
  readonly headerLabels?: readonly string[];
}

export interface CellBorder {
  readonly style: string;
  readonly rgb?: string;
}
export interface CellFill {
  readonly pattern: string;
  readonly patternRgb?: string;
  readonly gradientRgb?: string;
  readonly shading?: string;
  readonly variant?: number;
}
export interface CellStyle {
  readonly bold?: boolean;
  readonly numberFormat?: string;
  readonly fillRgb?: string;
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly textRgb?: string;
  readonly borders?: string;
  readonly alignment?: string;
  readonly wrap?: boolean;
  readonly doubleUnderline?: boolean;
  readonly verticalAlignment?: string;
  readonly strikethrough?: boolean;
  readonly script?: string;
  readonly underlineKind?: string;
  readonly fillDetails?: CellFill;
  readonly borderDetails?: Readonly<Partial<Record<
    'top' | 'right' | 'bottom' | 'left' | 'diagonalUp' | 'diagonalDown', CellBorder>>>;
  readonly textRotation?: number;
  readonly indent?: number;
  readonly shrinkToFit?: boolean;
  readonly readingOrder?: string;
  readonly justifyLastLine?: boolean;
}

export const formulaErrors = Object.freeze(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'] as const);
export type CellScalar =
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'boolean'; readonly value: boolean };
export type CellCache =
  | { readonly kind: 'number'; readonly value: number; readonly serializedText?: string }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly value: typeof formulaErrors[number] };

export interface PythonBinding {
  readonly name: string;
  readonly reference: string;
}

export interface PythonFigure {
  readonly pngBase64: string;
  readonly width: number;
  readonly height: number;
  /** SHA-256 of the PNG for the current live worker session; omitted from native files. */
  readonly figureId?: string;
}

export interface FigurePlacement {
  readonly id: DocumentId;
  readonly sheetId: DocumentId;
  readonly row: number;
  readonly column: number;
  readonly rows: number;
  readonly columns: number;
  readonly offsetX?: number;
  readonly offsetY?: number;
  readonly width?: number;
  readonly height?: number;
  readonly renderedFigure?: PythonFigure | null;
  readonly renderedSourceId?: string | null;
}

export interface PythonTable {
  readonly rows: number;
  readonly columns: number;
  readonly cells: readonly (CellCache | null)[];
  readonly notices: readonly string[];
}

export interface PythonBlock {
  readonly id: DocumentId;
  readonly source: string;
  readonly bindings: readonly PythonBinding[];
  readonly lastValue?: CellCache | null;
  readonly hasResult?: boolean;
  readonly inputFingerprint?: string | null;
  readonly runtimeFingerprint?: string | null;
  readonly status?: 'never' | 'succeeded' | 'failed' | 'cancelled' | 'timedOut' | 'workerExited' | string;
  readonly error?: string | null;
  readonly lastTable?: PythonTable | null;
  readonly spillActive?: boolean;
  readonly lastFigure?: PythonFigure | null;
  readonly placements?: readonly FigurePlacement[];
}

export type CellContent = CellScalar
  | { readonly kind: 'blank' }
  | { readonly kind: 'formula'; readonly source: string }
  | { readonly kind: 'python'; readonly blockId: DocumentId; readonly source: string; readonly block?: PythonBlock };
export interface CellDocument extends Coordinate {
  readonly id: DocumentId;
  readonly content: CellContent;
  readonly style?: CellStyle;
  readonly cache?: CellCache;
  readonly requiresExternalCalculation?: boolean;
  readonly spillOwner?: DocumentId;
}
export interface SheetDocument {
  readonly id: DocumentId;
  readonly name: string;
  readonly cells: readonly CellDocument[];
  readonly rows: readonly AxisDimension[];
  readonly columns: readonly AxisDimension[];
  readonly merges: readonly CellExtent[];
  readonly frozenPanes: FrozenPanes;
  readonly filter?: SheetFilter;
}
export interface WorkbookDocument {
  readonly id: DocumentId;
  readonly formatVersion: 1;
  readonly dateSystem: '1900';
  readonly sheets: readonly SheetDocument[];
  readonly names: readonly WorkbookName[];
}

export function emptySheet(id: DocumentId, name: string): SheetDocument {
  return { id, name, cells: [], rows: [], columns: [], merges: [], frozenPanes: { rows: 0, columns: 0 } };
}

export function newWorkbook(makeId: IdFactory): WorkbookDocument {
  return { id: makeId(), formatVersion: 1, dateSystem: '1900', sheets: [emptySheet(makeId(), 'Sheet1')], names: [] };
}

export function inRange(value: number, low: number, high: number): boolean {
  return Number.isInteger(value) && value >= low && value <= high;
}

export function isCoordinate({ row, column }: Coordinate): boolean {
  return inRange(row, 1, documentLimits.rows) && inRange(column, 1, documentLimits.columns);
}

export function cellKey({ row, column }: Coordinate): number {
  return (row - 1) * documentLimits.columns + column - 1;
}

const indexes = new WeakMap<SheetDocument, ReadonlyMap<number, CellDocument>>();

export function findCell(sheet: SheetDocument, point: Coordinate): CellDocument | undefined {
  if (!isCoordinate(point)) return undefined;
  // Cache indexes only for snapshots owned by the model. Callers may query mutable builders.
  if (!isOwned(sheet)) return sheet.cells.find(cell => cell.row === point.row && cell.column === point.column);
  let index = indexes.get(sheet);
  if (!index) {
    index = new Map(sheet.cells.map(cell => [cellKey(cell), cell]));
    indexes.set(sheet, index);
  }
  return index.get(cellKey(point));
}

export function findReferencedCell(document: WorkbookDocument, target: CellReference): CellDocument | undefined {
  const sheet = !target.invalid && document.sheets.find(sheet => sheet.id === target.sheetId);
  return sheet ? findCell(sheet, target) : undefined;
}

// Only DTOs owned here may be structurally shared. Shallow freezing caller objects
// would still let their owners mutate nested records after a command or adoption.
const owned = new WeakSet<object>();

function isOwned(value: object): boolean {
  return owned.has(value);
}

export function ownDto<T>(value: T): T {
  return clone(value, new Set<object>());

  function clone<V>(item: V, path: Set<object>): V {
    if (item === null || typeof item !== 'object') {
      if (typeof item === 'function' || typeof item === 'symbol' || typeof item === 'bigint') {
        throw new DocumentError('Expected a plain document value.');
      }
      return item;
    }
    if (isOwned(item)) return item;
    const array = Array.isArray(item);
    if (path.has(item) || (!array && Object.getPrototypeOf(item) !== Object.prototype)) {
      throw new DocumentError('Expected an acyclic plain document record.');
    }
    path.add(item);
    const copy = array ? Array.from(item, entry => clone(entry, path)) : Object.fromEntries(
      Object.entries(item).filter(([, entry]) => entry !== undefined).map(([key, entry]) => [key, clone(entry, path)]),
    );
    path.delete(item);
    Object.freeze(copy);
    owned.add(copy);
    return copy as V;
  }
}
