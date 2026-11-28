import { useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  defaultChartElements,
  type ChartDocument,
  type ChartSeries,
  type ChartType,
} from '../../core/chart';
import type { DocumentId, WorkbookDocument } from '../../core/workbook';
import { inferChartSeriesFromSource, resolveChartData } from '../../charts/chart-data';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import { ChartCanvas } from './ChartCanvas';
import {
  ChartCategoryListGlyph,
  ChartTypeGlyph,
  coreDialogCategories,
  type ChartDialogCategory,
  type ChartSubtypeOption,
  type ComboPresetId,
} from './chart-catalog';

const categoryOrder: readonly ChartDialogCategory[] = [
  'Column',
  'Line',
  'Pie',
  'Bar',
  'Area',
  'X Y (Scatter)',
  'Stock',
  'Radar',
  'Treemap',
  'Sunburst',
  'Histogram',
  'Box & Whisker',
  'Waterfall',
  'Funnel',
  'Combo',
];

const recommendedCandidates: readonly { readonly type: ChartType; readonly label: string; readonly reason: string }[] = [
  { type: 'clusteredColumn', label: 'Clustered Column', reason: 'Compare values across categories' },
  { type: 'clusteredBar', label: 'Clustered Bar', reason: 'Compare category labels horizontally' },
  { type: 'lineMarkers', label: 'Line with Markers', reason: 'Show trends over ordered categories' },
  { type: 'pie', label: 'Pie', reason: 'Show proportions of a whole for one series' },
  { type: 'waterfall', label: 'Waterfall', reason: 'Show running cumulative increases and decreases' },
  { type: 'histogram', label: 'Histogram', reason: 'Summarize numeric value distribution across bins' },
  { type: 'combo', label: 'Combo', reason: 'Combine columns and lines on primary/secondary axes' },
];

export interface ComboSeriesOverride {
  readonly chartType: ChartType;
  readonly secondaryAxis: boolean;
}

const comboSeriesTypes: readonly { readonly type: ChartType; readonly label: string }[] = [
  { type: 'clusteredColumn', label: 'Clustered Column' },
  { type: 'stackedColumn', label: 'Stacked Column' },
  { type: 'line', label: 'Line' },
  { type: 'lineMarkers', label: 'Line with Markers' },
  { type: 'area', label: 'Area' },
  { type: 'stackedArea', label: 'Stacked Area' },
];

export function buildComboPresetOverrides(
  count: number,
  preset: ComboPresetId,
): readonly ComboSeriesOverride[] {
  const split = Math.max(1, Math.ceil(count / 2));
  return Array.from({ length: count }, (_, idx) => {
    if (preset === 'areaCol') {
      return idx < split
        ? { chartType: 'stackedArea', secondaryAxis: false }
        : { chartType: 'clusteredColumn', secondaryAxis: false };
    }
    if (preset === 'colLineSecondary') {
      return idx < split
        ? { chartType: 'clusteredColumn', secondaryAxis: false }
        : { chartType: 'line', secondaryAxis: true };
    }
    return idx < split
      ? { chartType: 'clusteredColumn', secondaryAxis: false }
      : { chartType: 'line', secondaryAxis: false };
  });
}

function supportsDualOrientation(category: ChartDialogCategory): boolean {
  return category === 'Column' || category === 'Line' || category === 'Bar' || category === 'Area' || category === 'Radar';
}

function stockGuidance(type: ChartType): string {
  if (type === 'stockOhlc') return 'To create this stock chart, arrange the data on your sheet in this order: opening price, high price, low price, closing price. Use dates or stock names as labels.';
  if (type === 'stockVhlc') return 'To create this stock chart, arrange the data on your sheet in this order: volume traded, high price, low price, closing price. Use dates or stock names as labels.';
  if (type === 'stockVohlc') return 'To create this stock chart, arrange the data on your sheet in this order: volume traded, opening price, high price, low price, closing price. Use dates or stock names as labels.';
  return 'To create this stock chart, arrange the data on your sheet in this order: high price, low price, closing price. Use dates or stock names as labels.';
}

export function InsertChartDialog({
  document,
  sheetId,
  sourceFormula,
  initialTab,
  initialCategory,
  onSubmit,
  onCancel,
}: {
  readonly document: WorkbookDocument;
  readonly sheetId: DocumentId;
  readonly sourceFormula: string;
  readonly initialTab: 'recommended' | 'all';
  readonly initialCategory?: ChartDialogCategory;
  readonly onSubmit: (
    chartType: ChartType,
    seriesBy?: 'columns' | 'rows',
    comboOverrides?: readonly ComboSeriesOverride[],
  ) => void;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [tab, setTab] = useState<'recommended' | 'all'>(initialTab);
  const [category, setCategory] = useState<ChartDialogCategory>(initialCategory ?? 'Column');
  const [selectedType, setSelectedType] = useState<ChartType>(
    initialTab === 'recommended'
      ? 'clusteredColumn'
      : coreDialogCategories[initialCategory ?? 'Column'][0]!.type,
  );
  const [comboPreset, setComboPreset] = useState<ComboPresetId>('colLine');
  const [comboOverrides, setComboOverrides] = useState<readonly ComboSeriesOverride[] | undefined>(undefined);

  const inferredBase = useMemo(
    () => inferChartSeriesFromSource(document, { kind: 'range', formula: sourceFormula }, sheetId, selectedType),
    [document, sheetId, sourceFormula, selectedType],
  );
  const [seriesByOverride, setSeriesByOverride] = useState<'columns' | 'rows' | undefined>(undefined);
  const activeSeriesBy = seriesByOverride ?? inferredBase.seriesBy;

  const activeComboOverrides = useMemo(
    () => comboOverrides ?? buildComboPresetOverrides(inferredBase.series.length, comboPreset),
    [comboOverrides, inferredBase.series.length, comboPreset],
  );

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (!dialog.open) dialog.showModal();
    dialog.querySelector<HTMLButtonElement>('button[type="submit"]')?.focus();
  }, []);

  function buildPreview(orientation: 'columns' | 'rows') {
    const source = { kind: 'range' as const, formula: sourceFormula };
    const inferred = inferChartSeriesFromSource(document, source, sheetId, selectedType, orientation);
    const series: ChartSeries[] = inferred.series.map((s, idx) => {
      const ov = selectedType === 'combo' ? activeComboOverrides[idx] : undefined;
      return {
        id: `preview-${orientation}-${idx}`,
        ...s,
        ...(ov && { chartType: ov.chartType, secondaryAxis: ov.secondaryAxis }),
      };
    });
    const chart: ChartDocument = {
      id: `preview-chart-${orientation}`,
      name: 'Preview Chart',
      chartType: selectedType,
      placement: { sheetId, row: 1, column: 1, offsetX: 0, offsetY: 0, width: 320, height: 215 },
      source,
      seriesBy: inferred.seriesBy,
      series,
      elements: defaultChartElements(selectedType),
      hiddenEmptyMode: 'gaps',
      showHiddenCells: false,
      styleId: 1,
      paletteId: 'colorful-1',
    };
    return { chart, data: resolveChartData(document, chart) };
  }

  const colPreview = useMemo(
    () => buildPreview('columns'),
    [document, sheetId, sourceFormula, selectedType, activeComboOverrides],
  );
  const rowPreview = useMemo(
    () => buildPreview('rows'),
    [document, sheetId, sourceFormula, selectedType, activeComboOverrides],
  );
  const singlePreview = activeSeriesBy === 'columns' ? colPreview : rowPreview;

  const categoryOptions: readonly ChartSubtypeOption[] = coreDialogCategories[category];
  const selectedLabel = selectedType === 'combo'
    ? (categoryOptions.find(c => c.comboPreset === comboPreset)?.label ?? 'Combo')
    : (recommendedCandidates.find(c => c.type === selectedType)?.label ??
      categoryOptions.find(c => c.type === selectedType)?.label ??
      selectedType);
  const dualMode = tab === 'all' && supportsDualOrientation(category);

  const requiredStockSeries = selectedType === 'stockVohlc' ? 5
    : selectedType === 'stockOhlc' || selectedType === 'stockVhlc' ? 4
    : selectedType === 'stockHlc' ? 3 : 0;
  const stockTooFewSeries = selectedType.startsWith('stock') && singlePreview.data.series.length < requiredStockSeries;

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog insert-chart-dialog"
      aria-label="Insert Chart"
      onCancel={event => {
        event.preventDefault();
        onCancel();
      }}
    >
      <form
        onSubmit={event => {
          event.preventDefault();
          onSubmit(
            selectedType,
            activeSeriesBy,
            selectedType === 'combo' ? activeComboOverrides : undefined,
          );
        }}
      >
        <DialogTitleBar>
          <h2>Insert Chart</h2>
          <button
            type="button"
            className="format-dialog-close"
            aria-label="Close Insert Chart"
            onClick={onCancel}
          >
            ×
          </button>
        </DialogTitleBar>

        <div className="insert-chart-tabs" role="tablist" aria-label="Insert Chart Mode">
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'recommended'}
            className={tab === 'recommended' ? 'active' : undefined}
            onClick={() => setTab('recommended')}
          >
            Recommended Charts
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'all'}
            className={tab === 'all' ? 'active' : undefined}
            onClick={() => setTab('all')}
          >
            All Charts
          </button>
        </div>

        <div className="insert-chart-body">
          {tab === 'recommended' ? (
            <div className="insert-chart-recommended-list" role="listbox" aria-label="Recommended chart types">
              {recommendedCandidates.map(item => (
                <button
                  key={item.type}
                  type="button"
                  role="option"
                  aria-selected={selectedType === item.type}
                  className={'insert-chart-rec-item' + (selectedType === item.type ? ' selected' : '')}
                  onClick={() => {
                    setSelectedType(item.type);
                    setSeriesByOverride(undefined);
                    setComboOverrides(undefined);
                  }}
                  onDoubleClick={() => onSubmit(
                    item.type,
                    activeSeriesBy,
                    item.type === 'combo' ? activeComboOverrides : undefined,
                  )}
                >
                  <ChartTypeGlyph type={item.type} size={32} />
                  <span>
                    <strong>{item.label}</strong>
                    <small>{item.reason}</small>
                  </span>
                </button>
              ))}
            </div>
          ) : (
            <div className="insert-chart-categories" role="listbox" aria-label="Chart categories">
              {categoryOrder.map(cat => (
                <button
                  key={cat}
                  type="button"
                  role="option"
                  aria-selected={category === cat}
                  className={'insert-chart-cat-item' + (category === cat ? ' selected' : '')}
                  onClick={() => {
                    setCategory(cat);
                    const first = coreDialogCategories[cat][0]!;
                    setSelectedType(first.type);
                    if (first.comboPreset) setComboPreset(first.comboPreset);
                    setComboOverrides(undefined);
                    setSeriesByOverride(undefined);
                  }}
                >
                  <ChartCategoryListGlyph category={cat} size={18} />
                  <span>{cat}</span>
                </button>
              ))}
            </div>
          )}

          <div className="insert-chart-preview-col">
            {tab === 'all' && (
              <div className="insert-chart-subtypes" role="listbox" aria-label="Chart subtypes">
                {categoryOptions.map(opt => {
                  const isSelected = selectedType === opt.type &&
                    (!opt.comboPreset || comboPreset === opt.comboPreset);
                  return (
                    <button
                      key={opt.label}
                      type="button"
                      role="option"
                      aria-selected={isSelected}
                      title={opt.label}
                      aria-label={opt.label}
                      className={'insert-chart-subtype-btn' + (isSelected ? ' selected' : '')}
                      onClick={() => {
                        setSelectedType(opt.type);
                        if (opt.comboPreset) {
                          setComboPreset(opt.comboPreset);
                          if (opt.comboPreset !== 'custom') setComboOverrides(undefined);
                        }
                        setSeriesByOverride(undefined);
                      }}
                      onDoubleClick={() => onSubmit(
                        opt.type,
                        activeSeriesBy,
                        opt.type === 'combo' ? activeComboOverrides : undefined,
                      )}
                    >
                      <ChartTypeGlyph type={opt.type} comboPreset={opt.comboPreset} size={38} />
                    </button>
                  );
                })}
              </div>
            )}

            <div className="insert-chart-preview-pane">
              <div className="insert-chart-preview-label">{selectedLabel}</div>
              {stockTooFewSeries ? (
                <p className="insert-chart-stock-note">{stockGuidance(selectedType)}</p>
              ) : dualMode ? (
                <div className="insert-chart-preview-grid">
                  {(['columns', 'rows'] as const).map(orientation => {
                    const preview = orientation === 'columns' ? colPreview : rowPreview;
                    const chosen = activeSeriesBy === orientation;
                    return (
                      <button
                        key={orientation}
                        type="button"
                        className={'insert-chart-preview-card' + (chosen ? ' selected' : '')}
                        aria-pressed={chosen}
                        aria-label={`${selectedLabel} (${orientation === 'columns' ? 'Series in Columns' : 'Series in Rows'})`}
                        onClick={() => setSeriesByOverride(orientation)}
                        onDoubleClick={() => onSubmit(selectedType, orientation)}
                      >
                        <ChartCanvas chart={preview.chart} data={preview.data} width={260} height={190} interactive={false} />
                      </button>
                    );
                  })}
                </div>
              ) : (
                <div className="insert-chart-preview-single">
                  <div className="insert-chart-preview-card selected">
                    <ChartCanvas chart={singlePreview.chart} data={singlePreview.data} width={360} height={210} interactive={false} />
                  </div>
                </div>
              )}

              {selectedType === 'combo' && singlePreview.data.series.length > 0 && (
                <div className="insert-chart-combo-wrap">
                  <div className="insert-chart-combo-caption">
                    Choose the chart type and axis for your data series:
                  </div>
                  <div className="insert-chart-combo-table" role="table" aria-label="Combo series configuration">
                    <div className="insert-chart-combo-head" role="row">
                      <span role="columnheader">Series Name</span>
                      <span role="columnheader">Chart Type</span>
                      <span role="columnheader">Secondary Axis</span>
                    </div>
                    {singlePreview.data.series.map((s, idx) => {
                      const ov = activeComboOverrides[idx] ?? { chartType: 'clusteredColumn', secondaryAxis: false };
                      return (
                        <div key={s.id} className="insert-chart-combo-row" role="row">
                          <span role="cell" className="combo-series-name">{s.name}</span>
                          <span role="cell">
                            <select
                              aria-label={`Chart type for ${s.name}`}
                              value={ov.chartType}
                              onChange={event => {
                                const nextType = event.currentTarget.value as ChartType;
                                setComboPreset('custom');
                                setComboOverrides(activeComboOverrides.map((item, i) =>
                                  i === idx ? { ...item, chartType: nextType } : item));
                              }}
                            >
                              {comboSeriesTypes.map(t => (
                                <option key={t.type} value={t.type}>{t.label}</option>
                              ))}
                            </select>
                          </span>
                          <span role="cell" className="combo-secondary-cell">
                            <input
                              type="checkbox"
                              aria-label={`Secondary Axis for ${s.name}`}
                              checked={ov.secondaryAxis}
                              onChange={event => {
                                const checked = event.currentTarget.checked;
                                setComboPreset('custom');
                                setComboOverrides(activeComboOverrides.map((item, i) =>
                                  i === idx ? { ...item, secondaryAxis: checked } : item));
                              }}
                            />
                          </span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>

        <DialogActions disabled={stockTooFewSeries} onCancel={onCancel} />
      </form>
    </dialog>
  );
}
