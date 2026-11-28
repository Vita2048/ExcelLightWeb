import { useEffect, useRef, useState } from 'react';
import type {
  ChartDataLabelPosition,
  ChartDocument,
  ChartErrorBarType,
  ChartLegendPosition,
} from '../../core/chart';
import type { DocumentId } from '../../core/workbook';
import type { ChartUpdatePatch } from '../../application/chart-commands';
import type { ResolvedChartData } from '../../charts/chart-data';
import { chartPalettes, getSeriesColor } from '../../charts/palettes';
import {
  applyChartStylePatch,
  chartStylePresets,
  patchChartElements,
  setChartDataLabels,
  setChartLegend,
  supportedElementsForType,
  toggleChartAxis,
  toggleChartAxisTitle,
  toggleChartGridline,
  togglePrimaryErrorBars,
  togglePrimaryTrendline,
  type ChartSelectedElement,
} from './chart-chrome';

type QuickTab = 'elements' | 'styles' | 'filters';
type ElementKey =
  | 'axes'
  | 'axisTitles'
  | 'chartTitle'
  | 'dataLabels'
  | 'dataTable'
  | 'errorBars'
  | 'gridlines'
  | 'legend'
  | 'trendline';

function PlusIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <path d="M8 2.5 V13.5 M2.5 8 H13.5" stroke="#5cc471" strokeWidth="2.2" strokeLinecap="round" />
    </svg>
  );
}

function BrushIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <path d="M11.8 1.8 L14.2 4.2 L7.4 11 L5 8.6 Z" fill="#69b3e7" stroke="#d8dee9" strokeWidth="1" />
      <path d="M5 8.6 C3.4 9.2 3 11 2 12.8 C3.6 13.4 5.5 13 6.8 11.4 Z" fill="#d8dee9" />
    </svg>
  );
}

function FunnelIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
      <path d="M2 3 H14 L9.5 8.2 V13 L6.5 14 V8.2 Z" fill="none" stroke="#d8dee9" strokeWidth="1.4" strokeLinejoin="round" />
    </svg>
  );
}

function describeRefSource(refs: readonly (string | undefined)[], seriesBy: 'columns' | 'rows', kind: 'series' | 'categories'): string {
  const first = refs.find((r): r is string => !!r && r.startsWith('='));
  if (!first) return kind === 'series' ? 'Series labels' : 'Category labels';
  const bang = first.indexOf('!');
  const body = (bang >= 0 ? first.slice(bang + 1) : first.slice(1)).replace(/\$/g, '');
  const cell = body.split(':')[0] ?? body;
  const colMatch = /^([A-Z]+)(\d+)$/i.exec(cell);
  if (!colMatch) return body;
  const col = colMatch[1]!.toUpperCase();
  const row = colMatch[2]!;
  if (kind === 'series') {
    return seriesBy === 'columns' ? `Row ${row}` : `Column ${col}`;
  }
  return seriesBy === 'columns' ? `Column ${col}` : `Row ${row}`;
}

export function ChartQuickButtons({
  chart,
  data,
  flipLeft,
  onUpdateChart,
  onOpenSelectData,
  onOpenFormatPane,
}: {
  readonly chart: ChartDocument;
  readonly data: ResolvedChartData;
  readonly flipLeft?: boolean;
  readonly onUpdateChart: (chartId: DocumentId, patch: ChartUpdatePatch) => void;
  readonly onOpenSelectData: (chartId: DocumentId) => void;
  readonly onOpenFormatPane: (chartId: DocumentId, el?: ChartSelectedElement) => void;
}) {
  const rootRef = useRef<HTMLDivElement>(null);
  const elementsPopoverRef = useRef<HTMLDivElement>(null);
  const [openTab, setOpenTab] = useState<QuickTab | undefined>(undefined);
  const [subElement, setSubElement] = useState<ElementKey | undefined>(undefined);
  const [subTop, setSubTop] = useState<number>(28);
  const [styleTab, setStyleTab] = useState<'style' | 'color'>('style');
  const [filterTab, setFilterTab] = useState<'values' | 'names'>('values');

  const [stagedHiddenSeries, setStagedHiddenSeries] = useState<ReadonlySet<DocumentId>>(
    () => new Set(chart.series.filter(s => s.hidden).map(s => s.id)),
  );
  const [stagedHiddenCats, setStagedHiddenCats] = useState<ReadonlySet<number>>(
    () => new Set(chart.hiddenCategories ?? []),
  );

  // Staged Names state (matching Excel Chart Filters -> Names radio groups)
  const hasSeriesNameRefs = chart.series.some(s => !!s.nameRef);
  const hasCategoryRefs = chart.series.some(s => !!s.categoriesRef);
  const savedNameRefs = useRef(new Map<DocumentId, string>());
  const savedCatRefs = useRef(new Map<DocumentId, string>());
  for (const s of chart.series) {
    if (s.nameRef) savedNameRefs.current.set(s.id, s.nameRef);
    if (s.categoriesRef) savedCatRefs.current.set(s.id, s.categoriesRef);
  }
  const [stagedUseSeriesNames, setStagedUseSeriesNames] = useState(hasSeriesNameRefs);
  const [stagedUseCatNames, setStagedUseCatNames] = useState(hasCategoryRefs);

  useEffect(() => {
    setStagedHiddenSeries(new Set(chart.series.filter(s => s.hidden).map(s => s.id)));
    setStagedHiddenCats(new Set(chart.hiddenCategories ?? []));
    setStagedUseSeriesNames(chart.series.some(s => !!s.nameRef));
    setStagedUseCatNames(chart.series.some(s => !!s.categoriesRef));
  }, [chart.id, chart.series, chart.hiddenCategories]);

  useEffect(() => {
    if (!openTab) return;
    const onDown = (event: globalThis.PointerEvent) => {
      if (rootRef.current && !rootRef.current.contains(event.target as Node)) {
        setOpenTab(undefined);
        setSubElement(undefined);
      }
    };
    const onKey = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape') {
        setOpenTab(undefined);
        setSubElement(undefined);
      }
    };
    window.addEventListener('pointerdown', onDown, true);
    window.addEventListener('keydown', onKey);
    return () => {
      window.removeEventListener('pointerdown', onDown, true);
      window.removeEventListener('keydown', onKey);
    };
  }, [openTab]);

  const supported = supportedElementsForType(chart.chartType);
  const axesShown = chart.elements.axes?.category?.show !== false || chart.elements.axes?.value?.show !== false;
  const axisTitlesShown = !!chart.elements.axes?.category?.title?.show || !!chart.elements.axes?.value?.title?.show;
  const titleShown = chart.elements.title?.show !== false;
  const dataLabelsShown = !!chart.elements.dataLabels?.show || chart.series.some(s => s.dataLabels?.show);
  const dataTableShown = !!chart.elements.dataTable?.show;
  const errorBarsShown = chart.series.some(s => s.errorBars !== undefined);
  const gridlinesShown = !!chart.elements.gridlines?.primaryMajorHorizontal ||
    !!chart.elements.gridlines?.primaryMajorVertical ||
    !!chart.elements.gridlines?.primaryMinorHorizontal ||
    !!chart.elements.gridlines?.primaryMinorVertical;
  const legendShown = chart.elements.legend?.show !== false;
  const trendlineShown = chart.series.some(s => (s.trendlines?.length ?? 0) > 0);

  const elementItems: {
    readonly key: ElementKey;
    readonly label: string;
    readonly visible: boolean;
    readonly checked: boolean;
    readonly toggle: (checked: boolean) => void;
  }[] = [
    {
      key: 'axes',
      label: 'Axes',
      visible: supported.axes,
      checked: axesShown,
      toggle: checked => onUpdateChart(chart.id, patchChartElements(chart, {
        axes: {
          ...chart.elements.axes,
          category: { ...(chart.elements.axes?.category ?? { show: true }), show: checked },
          value: { ...(chart.elements.axes?.value ?? { show: true }), show: checked },
        },
      })),
    },
    {
      key: 'axisTitles',
      label: 'Axis Titles',
      visible: supported.axisTitles,
      checked: axisTitlesShown,
      toggle: checked => onUpdateChart(chart.id, patchChartElements(chart, {
        axes: {
          ...chart.elements.axes,
          category: {
            ...(chart.elements.axes?.category ?? { show: true }),
            title: { show: checked, text: chart.elements.axes?.category?.title?.text || 'Axis Title' },
          },
          value: {
            ...(chart.elements.axes?.value ?? { show: true }),
            title: { show: checked, text: chart.elements.axes?.value?.title?.text || 'Axis Title' },
          },
        },
      })),
    },
    {
      key: 'chartTitle',
      label: 'Chart Title',
      visible: supported.chartTitle,
      checked: titleShown,
      toggle: checked => onUpdateChart(chart.id, patchChartElements(chart, {
        title: {
          ...chart.elements.title,
          show: checked,
          text: chart.elements.title?.text || 'Chart Title',
        },
      })),
    },
    {
      key: 'dataLabels',
      label: 'Data Labels',
      visible: supported.dataLabels,
      checked: dataLabelsShown,
      toggle: checked => onUpdateChart(chart.id, setChartDataLabels(chart, checked)),
    },
    {
      key: 'dataTable',
      label: 'Data Table',
      visible: supported.dataTable,
      checked: dataTableShown,
      toggle: checked => onUpdateChart(chart.id, patchChartElements(chart, {
        dataTable: { show: checked, showLegendKeys: chart.elements.dataTable?.showLegendKeys ?? true },
      })),
    },
    {
      key: 'errorBars',
      label: 'Error Bars',
      visible: supported.errorBars,
      checked: errorBarsShown,
      toggle: checked => onUpdateChart(chart.id, checked
        ? togglePrimaryErrorBars(chart, 'percentage')
        : togglePrimaryErrorBars(chart, undefined)),
    },
    {
      key: 'gridlines',
      label: 'Gridlines',
      visible: supported.gridlines,
      checked: gridlinesShown,
      toggle: checked => onUpdateChart(chart.id, patchChartElements(chart, {
        gridlines: {
          ...chart.elements.gridlines,
          primaryMajorHorizontal: checked,
          primaryMajorVertical: checked ? !!chart.elements.gridlines?.primaryMajorVertical : false,
          ...(checked ? {} : { primaryMinorHorizontal: false, primaryMinorVertical: false }),
        },
      })),
    },
    {
      key: 'legend',
      label: 'Legend',
      visible: supported.legend,
      checked: legendShown,
      toggle: checked => onUpdateChart(chart.id, setChartLegend(chart, checked)),
    },
    {
      key: 'trendline',
      label: 'Trendline',
      visible: supported.trendline,
      checked: trendlineShown,
      toggle: checked => onUpdateChart(chart.id, checked
        ? togglePrimaryTrendline(chart, 'linear')
        : togglePrimaryTrendline(chart, undefined)),
    },
  ];

  const allSeriesChecked = stagedHiddenSeries.size === 0;
  const allCatsChecked = stagedHiddenCats.size === 0;

  function openSubForRow(key: ElementKey, rowEl: HTMLElement | null) {
    setSubElement(key);
    if (rowEl && elementsPopoverRef.current) {
      setSubTop(Math.max(4, rowEl.offsetTop - 2));
    }
  }

  function applyFilters() {
    const nextSeries = chart.series.map(s => {
      const hidden = stagedHiddenSeries.has(s.id);
      if (hidden) return { ...s, hidden: true };
      const { hidden: _prev, ...rest } = s;
      return rest;
    });
    const nextHiddenCats = [...stagedHiddenCats].sort((a, b) => a - b);
    onUpdateChart(chart.id, {
      series: nextSeries,
      hiddenCategories: nextHiddenCats.length > 0 ? nextHiddenCats : undefined,
    });
    setOpenTab(undefined);
  }

  function applyNames() {
    const nextSeries = chart.series.map(s => {
      const { nameRef: _nr, name: _nm, categoriesRef: _cr, ...rest } = s;
      const restoredNameRef = stagedUseSeriesNames ? (s.nameRef ?? savedNameRefs.current.get(s.id)) : undefined;
      const restoredCatRef = stagedUseCatNames ? (s.categoriesRef ?? savedCatRefs.current.get(s.id)) : undefined;
      return {
        ...rest,
        ...(restoredNameRef ? { nameRef: restoredNameRef } : stagedUseSeriesNames && s.name ? { name: s.name } : {}),
        ...(restoredCatRef && { categoriesRef: restoredCatRef }),
      };
    });
    onUpdateChart(chart.id, { series: nextSeries });
    setOpenTab(undefined);
  }

  const seriesSourceLabel = describeRefSource(
    chart.series.map(s => s.nameRef ?? savedNameRefs.current.get(s.id)),
    chart.seriesBy,
    'series',
  );
  const categorySourceLabel = describeRefSource(
    chart.series.map(s => s.categoriesRef ?? savedCatRefs.current.get(s.id)),
    chart.seriesBy,
    'categories',
  );

  return (
    <div
      ref={rootRef}
      className={'chart-quick-buttons' + (flipLeft ? ' flip-left' : '')}
      onPointerDown={event => event.stopPropagation()}
      onClick={event => event.stopPropagation()}
    >
      <div className="chart-quick-rail" role="toolbar" aria-label="Chart quick actions">
        <button
          type="button"
          className={'chart-quick-btn' + (openTab === 'elements' ? ' active' : '')}
          aria-label="Chart Elements"
          title="Chart Elements"
          aria-expanded={openTab === 'elements'}
          onClick={() => {
            setSubElement(undefined);
            setOpenTab(cur => (cur === 'elements' ? undefined : 'elements'));
          }}
        >
          <PlusIcon />
        </button>
        <button
          type="button"
          className={'chart-quick-btn' + (openTab === 'styles' ? ' active' : '')}
          aria-label="Chart Styles"
          title="Chart Styles"
          aria-expanded={openTab === 'styles'}
          onClick={() => {
            setSubElement(undefined);
            setOpenTab(cur => (cur === 'styles' ? undefined : 'styles'));
          }}
        >
          <BrushIcon />
        </button>
        <button
          type="button"
          className={'chart-quick-btn' + (openTab === 'filters' ? ' active' : '')}
          aria-label="Chart Filters"
          title="Chart Filters"
          aria-expanded={openTab === 'filters'}
          onClick={() => {
            setSubElement(undefined);
            setOpenTab(cur => (cur === 'filters' ? undefined : 'filters'));
          }}
        >
          <FunnelIcon />
        </button>
      </div>

      {openTab === 'elements' && (
        <div
          ref={elementsPopoverRef}
          className="chart-quick-popover chart-elements-popover"
          role="dialog"
          aria-label="Chart Elements"
        >
          <div className="chart-quick-popover-title">Chart Elements</div>
          <div className="chart-elements-list">
            {elementItems.filter(item => item.visible).map(item => (
              <div
                key={item.key}
                className={'chart-elements-row' + (subElement === item.key ? ' active' : '')}
                onMouseEnter={event => openSubForRow(item.key, event.currentTarget)}
              >
                <label className="chart-elements-check">
                  <input
                    type="checkbox"
                    checked={item.checked}
                    onChange={event => item.toggle(event.currentTarget.checked)}
                  />
                  <span>{item.label}</span>
                </label>
                <button
                  type="button"
                  className="chart-elements-more-btn"
                  aria-label={`${item.label} options`}
                  aria-expanded={subElement === item.key}
                  onClick={event => {
                    const rowEl = event.currentTarget.closest('.chart-elements-row') as HTMLElement | null;
                    openSubForRow(item.key, rowEl);
                  }}
                >
                  ▸
                </button>
              </div>
            ))}
          </div>

          {subElement && (
            <div
              className={'chart-elements-side-submenu' + (flipLeft ? ' flip-left' : '')}
              style={{ top: subTop }}
              role="menu"
              aria-label="Element options"
            >
              {subElement === 'axes' && (
                <>
                  <button type="button" role="menuitemcheckbox" aria-checked={chart.elements.axes?.category?.show !== false}
                    onClick={() => onUpdateChart(chart.id, toggleChartAxis(chart, 'category'))}>
                    Primary Horizontal
                  </button>
                  <button type="button" role="menuitemcheckbox" aria-checked={chart.elements.axes?.value?.show !== false}
                    onClick={() => onUpdateChart(chart.id, toggleChartAxis(chart, 'value'))}>
                    Primary Vertical
                  </button>
                  <button type="button" role="menuitem" onClick={() => { setOpenTab(undefined); onOpenFormatPane(chart.id, { kind: 'axis', axis: 'value' }); }}>
                    More Options...
                  </button>
                </>
              )}
              {subElement === 'axisTitles' && (
                <>
                  <button type="button" role="menuitemcheckbox" aria-checked={!!chart.elements.axes?.category?.title?.show}
                    onClick={() => onUpdateChart(chart.id, toggleChartAxisTitle(chart, 'category'))}>
                    Primary Horizontal
                  </button>
                  <button type="button" role="menuitemcheckbox" aria-checked={!!chart.elements.axes?.value?.title?.show}
                    onClick={() => onUpdateChart(chart.id, toggleChartAxisTitle(chart, 'value'))}>
                    Primary Vertical
                  </button>
                  <button type="button" role="menuitem" onClick={() => { setOpenTab(undefined); onOpenFormatPane(chart.id, { kind: 'axis', axis: 'value' }); }}>
                    More Options...
                  </button>
                </>
              )}
              {subElement === 'chartTitle' && (
                <>
                  <button type="button" role="menuitemradio" aria-checked={titleShown && !chart.elements.title?.overlay}
                    onClick={() => onUpdateChart(chart.id, patchChartElements(chart, { title: { ...chart.elements.title, show: true, overlay: false, text: chart.elements.title?.text || 'Chart Title' } }))}>
                    Above Chart
                  </button>
                  <button type="button" role="menuitemradio" aria-checked={titleShown && !!chart.elements.title?.overlay}
                    onClick={() => onUpdateChart(chart.id, patchChartElements(chart, { title: { ...chart.elements.title, show: true, overlay: true, text: chart.elements.title?.text || 'Chart Title' } }))}>
                    Centered Overlay
                  </button>
                  <button type="button" role="menuitem" onClick={() => { setOpenTab(undefined); onOpenFormatPane(chart.id, { kind: 'title' }); }}>
                    More Options...
                  </button>
                </>
              )}
              {subElement === 'dataLabels' && (
                <>
                  {(['center', 'insideEnd', 'insideBase', 'outsideEnd', 'callout'] as ChartDataLabelPosition[]).map(pos => (
                    <button key={pos} type="button" role="menuitemradio"
                      aria-checked={dataLabelsShown && chart.elements.dataLabels?.position === pos}
                      onClick={() => onUpdateChart(chart.id, setChartDataLabels(chart, true, pos))}>
                      {pos === 'center' ? 'Center'
                        : pos === 'insideEnd' ? 'Inside End'
                        : pos === 'insideBase' ? 'Inside Base'
                        : pos === 'outsideEnd' ? 'Outside End'
                        : 'Data Callout'}
                    </button>
                  ))}
                  <button type="button" role="menuitem" onClick={() => { setOpenTab(undefined); onOpenFormatPane(chart.id, { kind: 'dataLabels' }); }}>
                    More Options...
                  </button>
                </>
              )}
              {subElement === 'dataTable' && (
                <>
                  <button type="button" role="menuitemradio"
                    aria-checked={dataTableShown && chart.elements.dataTable?.showLegendKeys !== false}
                    onClick={() => onUpdateChart(chart.id, patchChartElements(chart, { dataTable: { show: true, showLegendKeys: true } }))}>
                    With Legend Keys
                  </button>
                  <button type="button" role="menuitemradio"
                    aria-checked={dataTableShown && chart.elements.dataTable?.showLegendKeys === false}
                    onClick={() => onUpdateChart(chart.id, patchChartElements(chart, { dataTable: { show: true, showLegendKeys: false } }))}>
                    No Legend Keys
                  </button>
                  <button type="button" role="menuitem" onClick={() => { setOpenTab(undefined); onOpenFormatPane(chart.id, { kind: 'chartArea' }); }}>
                    More Options...
                  </button>
                </>
              )}
              {subElement === 'errorBars' && (
                <>
                  {([
                    ['stdErr', 'Standard Error'],
                    ['percentage', 'Percentage'],
                    ['stdDev', 'Standard Deviation'],
                  ] as readonly [ChartErrorBarType, string][]).map(([type, label]) => (
                    <button key={type} type="button" role="menuitem"
                      onClick={() => onUpdateChart(chart.id, togglePrimaryErrorBars(chart, type))}>
                      {label}
                    </button>
                  ))}
                  <button type="button" role="menuitem" onClick={() => {
                    const owner = chart.series[0];
                    setOpenTab(undefined);
                    onOpenFormatPane(chart.id, owner ? { kind: 'series', seriesId: owner.id } : { kind: 'chartArea' });
                  }}>
                    More Options...
                  </button>
                </>
              )}
              {subElement === 'gridlines' && (
                <>
                  <button type="button" role="menuitemcheckbox" aria-checked={!!chart.elements.gridlines?.primaryMajorHorizontal}
                    onClick={() => onUpdateChart(chart.id, toggleChartGridline(chart, 'primaryMajorHorizontal'))}>
                    Primary Major Horizontal
                  </button>
                  <button type="button" role="menuitemcheckbox" aria-checked={!!chart.elements.gridlines?.primaryMajorVertical}
                    onClick={() => onUpdateChart(chart.id, toggleChartGridline(chart, 'primaryMajorVertical'))}>
                    Primary Major Vertical
                  </button>
                  <button type="button" role="menuitemcheckbox" aria-checked={!!chart.elements.gridlines?.primaryMinorHorizontal}
                    onClick={() => onUpdateChart(chart.id, toggleChartGridline(chart, 'primaryMinorHorizontal'))}>
                    Primary Minor Horizontal
                  </button>
                  <button type="button" role="menuitemcheckbox" aria-checked={!!chart.elements.gridlines?.primaryMinorVertical}
                    onClick={() => onUpdateChart(chart.id, toggleChartGridline(chart, 'primaryMinorVertical'))}>
                    Primary Minor Vertical
                  </button>
                  <button type="button" role="menuitem" onClick={() => { setOpenTab(undefined); onOpenFormatPane(chart.id, { kind: 'axis', axis: 'value' }); }}>
                    More Options...
                  </button>
                </>
              )}
              {subElement === 'legend' && (
                <>
                  {(['right', 'top', 'left', 'bottom'] as ChartLegendPosition[]).map(pos => (
                    <button key={pos} type="button" role="menuitemradio" aria-checked={legendShown && chart.elements.legend?.position === pos}
                      onClick={() => onUpdateChart(chart.id, setChartLegend(chart, true, pos))}>
                      {pos.charAt(0).toUpperCase() + pos.slice(1)}
                    </button>
                  ))}
                  <button type="button" role="menuitem" onClick={() => { setOpenTab(undefined); onOpenFormatPane(chart.id, { kind: 'legend' }); }}>
                    More Options...
                  </button>
                </>
              )}
              {subElement === 'trendline' && (
                <>
                  <button type="button" role="menuitem" onClick={() => onUpdateChart(chart.id, togglePrimaryTrendline(chart, 'linear'))}>
                    Linear
                  </button>
                  <button type="button" role="menuitem" onClick={() => onUpdateChart(chart.id, togglePrimaryTrendline(chart, 'exponential'))}>
                    Exponential
                  </button>
                  <button type="button" role="menuitem" onClick={() => onUpdateChart(chart.id, togglePrimaryTrendline(chart, 'linear', undefined, 2))}>
                    Linear Forecast
                  </button>
                  <button type="button" role="menuitem" onClick={() => onUpdateChart(chart.id, togglePrimaryTrendline(chart, 'movingAverage'))}>
                    Two Period Moving Average
                  </button>
                  <button type="button" role="menuitem" onClick={() => {
                    const owner = chart.series.find(s => (s.trendlines?.length ?? 0) > 0);
                    const tl = owner?.trendlines?.[0];
                    setOpenTab(undefined);
                    if (owner && tl) onOpenFormatPane(chart.id, { kind: 'trendline', seriesId: owner.id, trendlineId: tl.id });
                  }}>
                    More Options...
                  </button>
                </>
              )}
            </div>
          )}
        </div>
      )}

      {openTab === 'styles' && (
        <div className="chart-quick-popover chart-styles-popover" role="dialog" aria-label="Chart Styles">
          <div className="chart-quick-tabs" role="tablist">
            <button
              type="button"
              role="tab"
              aria-selected={styleTab === 'style'}
              className={styleTab === 'style' ? 'active' : undefined}
              onClick={() => setStyleTab('style')}
            >
              Style
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={styleTab === 'color'}
              className={styleTab === 'color' ? 'active' : undefined}
              onClick={() => setStyleTab('color')}
            >
              Color
            </button>
          </div>

          {styleTab === 'style' ? (
            <div className="chart-style-cards" role="listbox" aria-label="Style presets">
              {chartStylePresets.map(preset => (
                <button
                  key={preset.id}
                  type="button"
                  role="option"
                  aria-selected={chart.styleId === preset.id}
                  className={'chart-style-card' + (chart.styleId === preset.id ? ' selected' : '')}
                  onClick={() => onUpdateChart(chart.id, applyChartStylePatch(preset.id))}
                >
                  <svg width="56" height="34" viewBox="0 0 56 34" aria-hidden="true">
                    <rect x="1" y="1" width="54" height="32" rx="3" fill={preset.chartAreaBg} stroke={preset.chartAreaBorder === 'none' ? '#2c323c' : preset.chartAreaBorder} />
                    <rect x="7" y="6" width="42" height="22" fill={preset.plotAreaBg === 'transparent' ? 'none' : preset.plotAreaBg} stroke={preset.plotAreaBorder ?? 'none'} />
                    {preset.showGridlines && (
                      <>
                        <line x1="8" y1="13" x2="48" y2="13" stroke="#2b313b" strokeWidth="1" />
                        <line x1="8" y1="20" x2="48" y2="20" stroke="#2b313b" strokeWidth="1" />
                      </>
                    )}
                    <rect x="11" y="16" width="7" height="11" fill={getSeriesColor(chart.paletteId, 0)} />
                    <rect x="20" y="10" width="7" height="17" fill={getSeriesColor(chart.paletteId, 1)} />
                    <rect x="29" y="13" width="7" height="14" fill={getSeriesColor(chart.paletteId, 2)} />
                    <rect x="38" y="8" width="7" height="19" fill={getSeriesColor(chart.paletteId, 0)} />
                  </svg>
                  <span>{preset.name}</span>
                </button>
              ))}
            </div>
          ) : (
            <div className="chart-color-palettes">
              <div className="chart-palette-group-title">Colorful</div>
              {chartPalettes.filter(p => p.id.startsWith('colorful')).map(pal => (
                <button
                  key={pal.id}
                  type="button"
                  aria-label={pal.name}
                  aria-pressed={chart.paletteId === pal.id}
                  className={'chart-palette-row' + (chart.paletteId === pal.id ? ' selected' : '')}
                  onClick={() => onUpdateChart(chart.id, { paletteId: pal.id })}
                >
                  <span className="chart-palette-swatches">
                    {pal.colors.slice(0, 6).map(hex => (
                      <span key={hex} className="chart-palette-swatch" style={{ background: hex }} />
                    ))}
                  </span>
                </button>
              ))}
              <div className="chart-palette-group-title">Monochromatic</div>
              {chartPalettes.filter(p => p.id.startsWith('monochrome')).map(pal => (
                <button
                  key={pal.id}
                  type="button"
                  aria-label={pal.name}
                  aria-pressed={chart.paletteId === pal.id}
                  className={'chart-palette-row' + (chart.paletteId === pal.id ? ' selected' : '')}
                  onClick={() => onUpdateChart(chart.id, { paletteId: pal.id })}
                >
                  <span className="chart-palette-swatches">
                    {pal.colors.slice(0, 6).map(hex => (
                      <span key={hex} className="chart-palette-swatch" style={{ background: hex }} />
                    ))}
                  </span>
                </button>
              ))}
            </div>
          )}
        </div>
      )}

      {openTab === 'filters' && (
        <div className="chart-quick-popover chart-filters-popover" role="dialog" aria-label="Chart Filters">
          <div className="chart-quick-tabs" role="tablist">
            <button
              type="button"
              role="tab"
              aria-selected={filterTab === 'values'}
              className={filterTab === 'values' ? 'active' : undefined}
              onClick={() => setFilterTab('values')}
            >
              Values
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={filterTab === 'names'}
              className={filterTab === 'names' ? 'active' : undefined}
              onClick={() => setFilterTab('names')}
            >
              Names
            </button>
          </div>

          {filterTab === 'values' ? (
            <>
              <div className="chart-filter-scroll">
                <div className="chart-filter-section-title">SERIES</div>
                <label className="chart-filter-check">
                  <input
                    type="checkbox"
                    checked={allSeriesChecked}
                    onChange={event => {
                      if (event.currentTarget.checked) setStagedHiddenSeries(new Set());
                      else setStagedHiddenSeries(new Set(data.series.map(s => s.id)));
                    }}
                  />
                  <span>(Select All)</span>
                </label>
                {data.series.map((s, idx) => (
                  <label key={s.id} className="chart-filter-check">
                    <input
                      type="checkbox"
                      checked={!stagedHiddenSeries.has(s.id)}
                      onChange={event => {
                        const next = new Set(stagedHiddenSeries);
                        if (event.currentTarget.checked) next.delete(s.id);
                        else next.add(s.id);
                        setStagedHiddenSeries(next);
                      }}
                    />
                    <span
                      className="chart-filter-dot"
                      style={{ background: getSeriesColor(chart.paletteId, idx, s.sourceSeries.fillRgb) }}
                    />
                    <span>{s.name}</span>
                  </label>
                ))}

                <div className="chart-filter-section-title">CATEGORIES</div>
                <label className="chart-filter-check">
                  <input
                    type="checkbox"
                    checked={allCatsChecked}
                    onChange={event => {
                      if (event.currentTarget.checked) setStagedHiddenCats(new Set());
                      else setStagedHiddenCats(new Set(data.categories.map((_, i) => i)));
                    }}
                  />
                  <span>(Select All)</span>
                </label>
                {data.categories.map((cat, idx) => (
                  <label key={idx} className="chart-filter-check">
                    <input
                      type="checkbox"
                      checked={!stagedHiddenCats.has(idx)}
                      onChange={event => {
                        const next = new Set(stagedHiddenCats);
                        if (event.currentTarget.checked) next.delete(idx);
                        else next.add(idx);
                        setStagedHiddenCats(next);
                      }}
                    />
                    <span>{cat}</span>
                  </label>
                ))}
              </div>
              <div className="chart-filter-footer">
                <button type="button" className="chart-filter-apply" onClick={applyFilters}>Apply</button>
                <button
                  type="button"
                  className="chart-filter-link"
                  onClick={() => {
                    setOpenTab(undefined);
                    onOpenSelectData(chart.id);
                  }}
                >
                  Select Data...
                </button>
              </div>
            </>
          ) : (
            <>
              <div className="chart-filter-scroll">
                <details open className="chart-names-section">
                  <summary>Series</summary>
                  <div className="chart-names-group">
                    <label className="chart-names-radio">
                      <input
                        type="radio"
                        name="filter-series-names"
                        checked={stagedUseSeriesNames}
                        onChange={() => setStagedUseSeriesNames(true)}
                      />
                      <span>
                        <strong>{seriesSourceLabel}</strong>
                        <small>{data.series.map(s => s.name).join('; ') || 'Revenue'}</small>
                      </span>
                    </label>
                    <label className="chart-names-radio">
                      <input
                        type="radio"
                        name="filter-series-names"
                        checked={!stagedUseSeriesNames}
                        onChange={() => setStagedUseSeriesNames(false)}
                      />
                      <span>(None)</span>
                    </label>
                  </div>
                </details>

                <details open className="chart-names-section">
                  <summary>Categories</summary>
                  <div className="chart-names-group">
                    <label className="chart-names-radio">
                      <input
                        type="radio"
                        name="filter-cat-names"
                        checked={stagedUseCatNames}
                        onChange={() => setStagedUseCatNames(true)}
                      />
                      <span>
                        <strong>{categorySourceLabel}</strong>
                        <small>{data.categories.join('; ')}</small>
                      </span>
                    </label>
                    <label className="chart-names-radio">
                      <input
                        type="radio"
                        name="filter-cat-names"
                        checked={!stagedUseCatNames}
                        onChange={() => setStagedUseCatNames(false)}
                      />
                      <span>(None)</span>
                    </label>
                  </div>
                </details>
              </div>
              <div className="chart-filter-footer">
                <button type="button" className="chart-filter-apply" onClick={applyNames}>Apply</button>
                <button
                  type="button"
                  className="chart-filter-link"
                  onClick={() => {
                    setOpenTab(undefined);
                    onOpenSelectData(chart.id);
                  }}
                >
                  Select Data...
                </button>
              </div>
            </>
          )}
        </div>
      )}
    </div>
  );
}
