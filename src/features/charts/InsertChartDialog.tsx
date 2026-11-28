import { useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  defaultChartElements,
  type ChartDocument,
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
  type ChartSubtypeOption,
  type CoreChartCategoryTab,
} from './chart-catalog';

const categoryOrder: readonly CoreChartCategoryTab[] = [
  'Column',
  'Line',
  'Pie',
  'Bar',
  'Area',
  'X Y (Scatter)',
];

const recommendedCandidates: readonly { readonly type: ChartType; readonly label: string; readonly reason: string }[] = [
  { type: 'clusteredColumn', label: 'Clustered Column', reason: 'Compare values across categories' },
  { type: 'clusteredBar', label: 'Clustered Bar', reason: 'Compare category labels horizontally' },
  { type: 'lineMarkers', label: 'Line with Markers', reason: 'Show trends over ordered categories' },
  { type: 'pie', label: 'Pie', reason: 'Show proportions of a whole for one series' },
  { type: 'stackedColumn', label: 'Stacked Column', reason: 'Compare parts of a whole across categories' },
  { type: 'scatter', label: 'Scatter', reason: 'Compare pairs of numeric X/Y values' },
];

function supportsDualOrientation(category: CoreChartCategoryTab): boolean {
  return category === 'Column' || category === 'Line' || category === 'Bar' || category === 'Area';
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
  readonly initialCategory?: CoreChartCategoryTab;
  readonly onSubmit: (chartType: ChartType, seriesBy?: 'columns' | 'rows') => void;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [tab, setTab] = useState<'recommended' | 'all'>(initialTab);
  const [category, setCategory] = useState<CoreChartCategoryTab>(initialCategory ?? 'Column');
  const [selectedType, setSelectedType] = useState<ChartType>(
    initialTab === 'recommended'
      ? 'clusteredColumn'
      : coreDialogCategories[initialCategory ?? 'Column'][0]!.type,
  );
  const defaultSeriesBy = useMemo(
    () => inferChartSeriesFromSource(document, { kind: 'range', formula: sourceFormula }, sheetId, selectedType).seriesBy,
    [document, sheetId, sourceFormula, selectedType],
  );
  const [seriesByOverride, setSeriesByOverride] = useState<'columns' | 'rows' | undefined>(undefined);
  const activeSeriesBy = seriesByOverride ?? defaultSeriesBy;

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (!dialog.open) dialog.showModal();
    dialog.querySelector<HTMLButtonElement>('button[type="submit"]')?.focus();
  }, []);

  function buildPreview(orientation: 'columns' | 'rows') {
    const source = { kind: 'range' as const, formula: sourceFormula };
    const inferred = inferChartSeriesFromSource(document, source, sheetId, selectedType, orientation);
    const chart: ChartDocument = {
      id: `preview-chart-${orientation}`,
      name: 'Preview Chart',
      chartType: selectedType,
      placement: { sheetId, row: 1, column: 1, offsetX: 0, offsetY: 0, width: 320, height: 215 },
      source,
      seriesBy: inferred.seriesBy,
      series: inferred.series.map((s, idx) => ({ id: `preview-${orientation}-${idx}`, ...s })),
      elements: defaultChartElements(selectedType),
      hiddenEmptyMode: 'gaps',
      showHiddenCells: false,
      styleId: 1,
      paletteId: 'colorful-1',
    };
    return { chart, data: resolveChartData(document, chart) };
  }

  const colPreview = useMemo(() => buildPreview('columns'), [document, sheetId, sourceFormula, selectedType]);
  const rowPreview = useMemo(() => buildPreview('rows'), [document, sheetId, sourceFormula, selectedType]);
  const singlePreview = activeSeriesBy === 'columns' ? colPreview : rowPreview;

  const categoryOptions: readonly ChartSubtypeOption[] = coreDialogCategories[category];
  const selectedLabel =
    recommendedCandidates.find(c => c.type === selectedType)?.label ??
    categoryOptions.find(c => c.type === selectedType)?.label ??
    selectedType;
  const dualMode = tab === 'all' && supportsDualOrientation(category);

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
          onSubmit(selectedType, activeSeriesBy);
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
                  }}
                  onDoubleClick={() => onSubmit(item.type, activeSeriesBy)}
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
                    setSelectedType(coreDialogCategories[cat][0]!.type);
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
                {categoryOptions.map(opt => (
                  <button
                    key={opt.type}
                    type="button"
                    role="option"
                    aria-selected={selectedType === opt.type}
                    title={opt.label}
                    aria-label={opt.label}
                    className={'insert-chart-subtype-btn' + (selectedType === opt.type ? ' selected' : '')}
                    onClick={() => {
                      setSelectedType(opt.type);
                      setSeriesByOverride(undefined);
                    }}
                    onDoubleClick={() => onSubmit(opt.type, activeSeriesBy)}
                  >
                    <ChartTypeGlyph type={opt.type} size={38} />
                  </button>
                ))}
              </div>
            )}

            <div className="insert-chart-preview-pane">
              <div className="insert-chart-preview-label">{selectedLabel}</div>
              {dualMode ? (
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
                    <ChartCanvas chart={singlePreview.chart} data={singlePreview.data} width={360} height={220} interactive={false} />
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>

        <DialogActions onCancel={onCancel} />
      </form>
    </dialog>
  );
}
