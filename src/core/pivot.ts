import type {
  CellExtent, CellReference, DocumentId, WorkbookDocument, WorkbookName,
} from './workbook';

export type PivotSourceArea = WorkbookName['area'];

export type PivotAggregation =
  | 'sum'
  | 'count'
  | 'average'
  | 'max'
  | 'min'
  | 'product'
  | 'countNumbers'
  | 'stdDev'
  | 'stdDevp'
  | 'var'
  | 'varp';

export type PivotShowValuesAs =
  | 'none'
  | 'percentOfGrandTotal'
  | 'percentOfColumnTotal'
  | 'percentOfRowTotal'
  | 'percentOfParentRowTotal'
  | 'percentOfParentColumnTotal'
  | 'differenceFrom'
  | 'percentDifferenceFrom'
  | 'runningTotalIn'
  | 'rankAscending'
  | 'rankDescending'
  | 'index';

export type PivotReportLayout = 'compact' | 'outline' | 'tabular';
export type PivotSubtotalPosition = 'top' | 'bottom' | 'off';
export type PivotStyleId = 'Blue' | 'Green' | 'Orange' | 'Purple';

export interface PivotSortRule {
  readonly direction: 'asc' | 'desc';
  readonly byValueFieldId?: DocumentId;
}

export type PivotSource =
  | { readonly kind: 'range'; readonly area: PivotSourceArea }
  | { readonly kind: 'table'; readonly tableId: DocumentId };

export interface PivotAxisField {
  readonly fieldName: string;
  readonly subtotal?: 'automatic' | 'none';
  readonly sort?: PivotSortRule;
  readonly hiddenItems?: readonly string[];
}

export interface PivotValueField {
  readonly id: DocumentId;
  readonly fieldName: string;
  readonly customName?: string;
  readonly aggregation: PivotAggregation;
  readonly showValuesAs?: PivotShowValuesAs;
  readonly baseField?: string;
  readonly baseItem?: string;
  readonly numberFormat?: string;
}

export interface PivotFilterField {
  readonly fieldName: string;
  readonly hiddenItems?: readonly string[];
  readonly multiSelect?: boolean;
}

export interface PivotTableDocument {
  readonly id: DocumentId;
  readonly name: string;
  readonly source: PivotSource;
  readonly target: CellReference;
  readonly extent?: CellExtent;
  readonly filters: readonly PivotFilterField[];
  readonly columns: readonly PivotAxisField[];
  readonly rows: readonly PivotAxisField[];
  readonly values: readonly PivotValueField[];
  readonly valuesInRows?: boolean;
  readonly reportLayout: PivotReportLayout;
  readonly subtotals: PivotSubtotalPosition;
  readonly grandTotals: { readonly rows: boolean; readonly columns: boolean };
  readonly repeatLabels?: boolean;
  readonly blankRows?: boolean;
  readonly showButtons?: boolean;
  readonly showFieldHeaders?: boolean;
  readonly emptyCellText?: string;
  readonly autoFitColumns?: boolean;
  readonly style: PivotStyleId;
  readonly styleOptions: {
    readonly rowHeaders: boolean;
    readonly columnHeaders: boolean;
    readonly bandedRows: boolean;
    readonly bandedColumns: boolean;
  };
  readonly collapsedKeys?: readonly string[];
}

const aggregationLabels: Readonly<Record<PivotAggregation, string>> = {
  sum: 'Sum of',
  count: 'Count of',
  average: 'Average of',
  max: 'Max of',
  min: 'Min of',
  product: 'Product of',
  countNumbers: 'Count Numbers of',
  stdDev: 'StdDev of',
  stdDevp: 'StdDevp of',
  var: 'Var of',
  varp: 'Varp of',
};

export const pivotAggregations: readonly PivotAggregation[] = [
  'sum', 'count', 'average', 'max', 'min', 'product', 'countNumbers', 'stdDev', 'stdDevp', 'var', 'varp',
];

export const pivotShowValuesAsModes: readonly PivotShowValuesAs[] = [
  'none',
  'percentOfGrandTotal',
  'percentOfColumnTotal',
  'percentOfRowTotal',
  'percentOfParentRowTotal',
  'percentOfParentColumnTotal',
  'differenceFrom',
  'percentDifferenceFrom',
  'runningTotalIn',
  'rankAscending',
  'rankDescending',
  'index',
];

export const pivotReportLayouts: readonly PivotReportLayout[] = ['compact', 'outline', 'tabular'];
export const pivotSubtotalPositions: readonly PivotSubtotalPosition[] = ['top', 'bottom', 'off'];
export const pivotStyles: readonly PivotStyleId[] = ['Blue', 'Green', 'Orange', 'Purple'];

export function defaultPivotValueLabel(fieldName: string, aggregation: PivotAggregation): string {
  return `${aggregationLabels[aggregation]} ${fieldName}`;
}

export function nextPivotName(document: WorkbookDocument): string {
  const used = new Set((document.pivots ?? []).map(pivot => pivot.name.toUpperCase()));
  for (let index = 1; ; index++) {
    const candidate = `PivotTable${index}`;
    if (!used.has(candidate.toUpperCase())) return candidate;
  }
}
