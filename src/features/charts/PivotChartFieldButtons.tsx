import type { ChartDocument } from '../../core/chart';
import type { DocumentId, WorkbookDocument } from '../../core/workbook';
import { describePivotChartFields, type PivotChartFieldButtonDescriptor } from '../../charts/pivot-chart';

export interface PivotChartFieldButtonsProps {
  readonly document: WorkbookDocument;
  readonly chart: ChartDocument;
  readonly onPivotFilter: (
    pivotId: DocumentId,
    kind: 'reportFilter' | 'rowLabel' | 'colLabel',
    fieldNames: readonly string[],
    anchor: { readonly x: number; readonly y: number },
  ) => void;
}

function FilterGlyph({ active, sorted }: { readonly active: boolean; readonly sorted?: 'asc' | 'desc' | undefined }) {
  return (
    <svg width="10" height="10" viewBox="0 0 12 12" aria-hidden="true">
      {active ? (
        <path d="M1.5 2 H10.5 L7 6 V10 L5 11 V6 Z" fill="var(--accent)" />
      ) : (
        <path d="M2.5 4 L6 8 L9.5 4" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
      )}
      {sorted && (
        <path
          d={sorted === 'asc' ? 'M10 10 V2 M8.5 3.5 L10 2 L11.5 3.5' : 'M10 2 V10 M8.5 8.5 L10 10 L11.5 8.5'}
          fill="none"
          stroke="currentColor"
          strokeWidth="1.2"
          strokeLinecap="round"
        />
      )}
    </svg>
  );
}

export function PivotChartFieldButtons({
  document,
  chart,
  onPivotFilter,
}: PivotChartFieldButtonsProps) {
  if (chart.showFieldButtons === false) return null;
  const summary = describePivotChartFields(document, chart);
  if (!summary) return null;

  const { pivot, reportFilters, valueFields, axisFields, legendFields } = summary;

  function renderFilterButton(
    btn: PivotChartFieldButtonDescriptor,
    prefix: 'Filter' | 'Filter axis field' | 'Filter legend field',
  ) {
    if (btn.kind === 'value') return null;
    const interactiveKind = btn.kind;
    return (
      <button
        key={`${interactiveKind}:${btn.fieldName}`}
        type="button"
        className={'pivot-chart-field-btn' + (btn.filtered ? ' filtered' : '')}
        aria-label={`${prefix} ${btn.label}`}
        data-filtered={btn.filtered ? 'true' : undefined}
        onPointerDown={event => event.stopPropagation()}
        onClick={event => {
          event.stopPropagation();
          const rect = event.currentTarget.getBoundingClientRect();
          onPivotFilter(pivot.id, interactiveKind, [btn.fieldName], {
            x: rect.left,
            y: rect.bottom + 4,
          });
        }}
      >
        <span>{btn.label}</span>
        <FilterGlyph active={btn.filtered} sorted={btn.sorted} />
      </button>
    );
  }

  return (
    <div className="pivot-chart-field-layer" aria-label="PivotChart Field Buttons">
      {(valueFields.length > 0 || reportFilters.length > 0) && (
        <div className="pivot-chart-buttons-tl">
          {valueFields.map(v => (
            <span key={`val:${v.fieldName}:${v.label}`} className="pivot-chart-field-pill value-pill">
              {v.label}
            </span>
          ))}
          {reportFilters.map(f => renderFilterButton(f, 'Filter'))}
        </div>
      )}

      {legendFields.length > 0 && (
        <div className="pivot-chart-buttons-tr">
          {legendFields.map(f => renderFilterButton(f, 'Filter legend field'))}
        </div>
      )}

      {axisFields.length > 0 && (
        <div className="pivot-chart-buttons-bl">
          {axisFields.map(f => renderFilterButton(f, 'Filter axis field'))}
        </div>
      )}
    </div>
  );
}
