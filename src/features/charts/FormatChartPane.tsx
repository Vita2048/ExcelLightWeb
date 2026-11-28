import { useEffect, useState } from 'react';
import {
  chartCategoryOf,
  type ChartAxisConfig,
  type ChartAxisDisplayUnits,
  type ChartBinMode,
  type ChartDataLabelPosition,
  type ChartDocument,
  type ChartLegendPosition,
  type ChartLineDash,
  type ChartPlacement,
  type ChartPlacementMode,
  type ChartQuartileMethod,
  type ChartSeries,
  type ChartTrendline,
  type ChartTrendlineType,
} from '../../core/chart';
import type { ResolvedChartData } from '../../charts/chart-data';
import type { ChartUpdatePatch } from '../../application/chart-commands';
import {
  formatElementTitle,
  listChartElements,
  serializeSelectedElementKey,
  type ChartSelectedElement,
} from './chart-chrome';

type PaneTab = 'fillLine' | 'sizeProps' | 'options';

function FillBucketIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true">
      <path d="M9 2 L14.5 7.5 L8.5 13.5 L3 8 Z" fill="none" stroke="currentColor" strokeWidth="1.4" />
      <path d="M14.5 11 C14.5 12.8 13.4 14.5 14.5 15.5 C15.6 14.5 15.5 12.8 14.5 11 Z" fill="#5cc471" />
    </svg>
  );
}

function SizePropsIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true">
      <rect x="4" y="4" width="10" height="10" fill="none" stroke="currentColor" strokeWidth="1.3" />
      <path d="M2 6 V12 M16 6 V12 M6 2 H12 M6 16 H12" stroke="currentColor" strokeWidth="1.2" />
    </svg>
  );
}

function OptionsBarsIcon() {
  return (
    <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true">
      <rect x="3" y="9" width="3" height="6" fill="currentColor" />
      <rect x="7.5" y="5" width="3" height="10" fill="#4ea1ff" />
      <rect x="12" y="7" width="3" height="8" fill="currentColor" />
    </svg>
  );
}

export function FormatChartPane({
  width,
  chart,
  data,
  selectedElement,
  onSelectElement,
  onUpdateChart,
  onMoveResizeChart,
  onClose,
}: {
  readonly width: number;
  readonly chart: ChartDocument;
  readonly data: ResolvedChartData;
  readonly selectedElement: ChartSelectedElement;
  readonly onSelectElement: (el: ChartSelectedElement) => void;
  readonly onUpdateChart: (patch: ChartUpdatePatch) => void;
  readonly onMoveResizeChart: (placement: Partial<ChartPlacement>) => void;
  readonly onClose: () => void;
}) {
  const supportsFillLine =
    selectedElement.kind === 'chartArea' ||
    selectedElement.kind === 'plotArea' ||
    selectedElement.kind === 'series';
  const supportsSizeProps = selectedElement.kind === 'chartArea';
  const supportsOptions =
    selectedElement.kind === 'series' ||
    selectedElement.kind === 'axis' ||
    selectedElement.kind === 'title' ||
    selectedElement.kind === 'dataLabels' ||
    selectedElement.kind === 'legend' ||
    selectedElement.kind === 'trendline';

  const [tab, setTab] = useState<PaneTab>('fillLine');

  useEffect(() => {
    if (tab === 'fillLine' && !supportsFillLine) {
      setTab('options');
    } else if (tab === 'sizeProps' && !supportsSizeProps) {
      setTab(supportsOptions ? 'options' : 'fillLine');
    } else if (tab === 'options' && !supportsOptions) {
      setTab('fillLine');
    }
  }, [tab, supportsFillLine, supportsSizeProps, supportsOptions]);

  const choices = listChartElements(chart, data);
  const selectedKey = serializeSelectedElementKey(selectedElement);
  const category = chartCategoryOf(chart.chartType);

  const selectedSeries: ChartSeries | undefined = selectedElement.kind === 'series'
    ? (chart.series.find(s => s.id === selectedElement.seriesId) ?? chart.series[0])
    : selectedElement.kind === 'trendline'
      ? chart.series.find(s => s.id === selectedElement.seriesId)
      : undefined;

  const selectedTrendline: ChartTrendline | undefined = selectedElement.kind === 'trendline' && selectedSeries
    ? selectedSeries.trendlines?.find(t => t.id === selectedElement.trendlineId)
    : undefined;

  // Helpers for updating styles/axes/series
  function patchSelectedSeries(fn: (s: ChartSeries) => ChartSeries) {
    if (!selectedSeries) return;
    onUpdateChart({
      series: chart.series.map(s => (s.id === selectedSeries.id ? fn(s) : s)),
    });
  }

  function currentFillInfo(): { readonly mode: 'auto' | 'none' | 'solid'; readonly color: string } {
    if (selectedElement.kind === 'plotArea') {
      const rgb = chart.plotAreaStyle?.fillRgb;
      return rgb === 'none' ? { mode: 'none', color: '#181b20' }
        : rgb ? { mode: 'solid', color: rgb } : { mode: 'auto', color: '#181b20' };
    }
    if (selectedElement.kind === 'series' && selectedSeries) {
      const rgb = selectedSeries.fillRgb;
      return rgb ? { mode: 'solid', color: rgb } : { mode: 'auto', color: '#4ea1ff' };
    }
    const rgb = chart.chartAreaStyle?.fillRgb;
    return rgb === 'none' ? { mode: 'none', color: '#181b20' }
      : rgb ? { mode: 'solid', color: rgb } : { mode: 'auto', color: '#181b20' };
  }

  function currentBorderInfo(): { readonly mode: 'auto' | 'none' | 'solid'; readonly color: string; readonly width: number; readonly dash: ChartLineDash } {
    if (selectedElement.kind === 'plotArea') {
      const rgb = chart.plotAreaStyle?.borderRgb;
      const w = chart.plotAreaStyle?.borderWidth ?? 1;
      const d = chart.plotAreaStyle?.borderDash ?? 'solid';
      return rgb === 'none' ? { mode: 'none', color: '#353b45', width: w, dash: d }
        : rgb ? { mode: 'solid', color: rgb, width: w, dash: d } : { mode: 'auto', color: '#353b45', width: w, dash: d };
    }
    if (selectedElement.kind === 'series' && selectedSeries) {
      const rgb = selectedSeries.borderRgb;
      const w = selectedSeries.lineWidth ?? 2;
      const d = selectedSeries.lineDash ?? 'solid';
      return rgb ? { mode: 'solid', color: rgb, width: w, dash: d } : { mode: 'auto', color: '#4ea1ff', width: w, dash: d };
    }
    const rgb = chart.chartAreaStyle?.borderRgb;
    const w = chart.chartAreaStyle?.borderWidth ?? 1;
    const d = chart.chartAreaStyle?.borderDash ?? 'solid';
    return rgb === 'none' ? { mode: 'none', color: '#353b45', width: w, dash: d }
      : rgb ? { mode: 'solid', color: rgb, width: w, dash: d } : { mode: 'auto', color: '#353b45', width: w, dash: d };
  }

  function setFill(mode: 'auto' | 'none' | 'solid', color?: string) {
    if (selectedElement.kind === 'plotArea') {
      const { fillRgb: _prev, ...rest } = chart.plotAreaStyle ?? {};
      onUpdateChart({
        plotAreaStyle: mode === 'auto' ? rest : { ...rest, fillRgb: mode === 'none' ? 'none' : (color ?? '#1e2229') },
      });
    } else if (selectedElement.kind === 'series' && selectedSeries) {
      patchSelectedSeries(s => {
        const { fillRgb: _prev, ...rest } = s;
        return mode === 'solid' ? { ...rest, fillRgb: color ?? '#4ea1ff' } : rest;
      });
    } else {
      const { fillRgb: _prev, ...rest } = chart.chartAreaStyle ?? {};
      onUpdateChart({
        chartAreaStyle: mode === 'auto' ? rest : { ...rest, fillRgb: mode === 'none' ? 'none' : (color ?? '#181b20') },
      });
    }
  }

  function setBorder(mode: 'auto' | 'none' | 'solid', color?: string, bw?: number, dash?: ChartLineDash) {
    if (selectedElement.kind === 'plotArea') {
      const { borderRgb: _prev, ...rest } = chart.plotAreaStyle ?? {};
      onUpdateChart({
        plotAreaStyle: {
          ...rest,
          ...(mode !== 'auto' && { borderRgb: mode === 'none' ? 'none' : (color ?? '#4ea1ff') }),
          ...(bw !== undefined && { borderWidth: bw }),
          ...(dash !== undefined && { borderDash: dash }),
        },
      });
    } else if (selectedElement.kind === 'series' && selectedSeries) {
      patchSelectedSeries(s => {
        const { borderRgb: _prev, ...rest } = s;
        return {
          ...rest,
          ...(mode === 'solid' && { borderRgb: color ?? '#ffffff' }),
          ...(bw !== undefined && { lineWidth: bw }),
          ...(dash !== undefined && { lineDash: dash }),
        };
      });
    } else {
      const { borderRgb: _prev, ...rest } = chart.chartAreaStyle ?? {};
      onUpdateChart({
        chartAreaStyle: {
          ...rest,
          ...(mode !== 'auto' && { borderRgb: mode === 'none' ? 'none' : (color ?? '#4ea1ff') }),
          ...(bw !== undefined && { borderWidth: bw }),
          ...(dash !== undefined && { borderDash: dash }),
        },
      });
    }
  }

  const fillInfo = currentFillInfo();
  const borderInfo = currentBorderInfo();

  const axisWhich = selectedElement.kind === 'axis' ? selectedElement.axis : 'value';
  const axisCfg: ChartAxisConfig = chart.elements.axes?.[axisWhich] ?? { show: true };

  function updateAxisCfg(patch: Partial<ChartAxisConfig>) {
    const prevAxes = chart.elements.axes ?? {};
    onUpdateChart({
      elements: {
        ...chart.elements,
        axes: {
          ...prevAxes,
          [axisWhich]: { ...axisCfg, ...patch },
        },
      },
    });
  }

  function clearAxisField<K extends keyof ChartAxisConfig>(key: K) {
    const prevAxes = chart.elements.axes ?? {};
    const { [key]: _removed, ...rest } = axisCfg;
    onUpdateChart({
      elements: {
        ...chart.elements,
        axes: {
          ...prevAxes,
          [axisWhich]: rest as ChartAxisConfig,
        },
      },
    });
  }

  const placementMode: ChartPlacementMode = chart.placement.mode ?? 'moveAndSize';

  return (
    <aside className="format-chart-pane" style={{ width }} role="complementary" aria-label="Format Chart">
      <div className="format-chart-header">
        <div className="format-chart-title-row">
          <h2>{formatElementTitle(selectedElement, chart)}</h2>
          <button type="button" className="format-dialog-close" aria-label="Close Format pane" onClick={onClose}>×</button>
        </div>
        <select
          className="format-chart-element-select"
          aria-label="Chart Options"
          value={choices.some(c => c.key === selectedKey) ? selectedKey : 'chartArea'}
          onChange={event => {
            const found = choices.find(c => c.key === event.currentTarget.value);
            if (found) onSelectElement(found.element);
          }}
        >
          {choices.map(c => (
            <option key={c.key} value={c.key}>{c.label}</option>
          ))}
        </select>

        <div className="format-chart-icon-tabs" role="tablist" aria-label="Format categories">
          {supportsFillLine && (
            <button
              type="button"
              role="tab"
              aria-selected={tab === 'fillLine'}
              aria-label="Fill and Line"
              title="Fill & Line"
              className={tab === 'fillLine' ? 'active' : undefined}
              onClick={() => setTab('fillLine')}
            >
              <FillBucketIcon />
            </button>
          )}
          {supportsSizeProps && (
            <button
              type="button"
              role="tab"
              aria-selected={tab === 'sizeProps'}
              aria-label="Size and Properties"
              title="Size & Properties"
              className={tab === 'sizeProps' ? 'active' : undefined}
              onClick={() => setTab('sizeProps')}
            >
              <SizePropsIcon />
            </button>
          )}
          {supportsOptions && (
            <button
              type="button"
              role="tab"
              aria-selected={tab === 'options'}
              aria-label="Element Options"
              title="Element Options"
              className={tab === 'options' ? 'active' : undefined}
              onClick={() => setTab('options')}
            >
              <OptionsBarsIcon />
            </button>
          )}
        </div>
      </div>

      <div className="format-chart-body">
        {tab === 'fillLine' && supportsFillLine && (
          <>
            <details open className="format-chart-section">
              <summary>Fill</summary>
              <div className="format-chart-group">
                {selectedElement.kind !== 'series' && (
                  <label className="format-chart-radio">
                    <input type="radio" name="chart-fill" checked={fillInfo.mode === 'auto'} onChange={() => setFill('auto')} />
                    Automatic
                  </label>
                )}
                {selectedElement.kind !== 'series' && (
                  <label className="format-chart-radio">
                    <input type="radio" name="chart-fill" checked={fillInfo.mode === 'none'} onChange={() => setFill('none')} />
                    No fill
                  </label>
                )}
                {selectedElement.kind === 'series' && (
                  <label className="format-chart-radio">
                    <input type="radio" name="chart-fill" checked={fillInfo.mode === 'auto'} onChange={() => setFill('auto')} />
                    Automatic (Palette)
                  </label>
                )}
                <label className="format-chart-radio">
                  <input type="radio" name="chart-fill" checked={fillInfo.mode === 'solid'} onChange={() => setFill('solid', fillInfo.color)} />
                  Solid fill
                </label>
                <label className="format-chart-field-row">
                  <span>Color</span>
                  <input
                    type="color"
                    aria-label="Fill color"
                    value={fillInfo.color.startsWith('#') ? fillInfo.color : '#181b20'}
                    onChange={event => setFill('solid', event.currentTarget.value)}
                  />
                </label>
              </div>
            </details>

            <details open className="format-chart-section">
              <summary>Border</summary>
              <div className="format-chart-group">
                <label className="format-chart-radio">
                  <input type="radio" name="chart-border" checked={borderInfo.mode === 'auto'} onChange={() => setBorder('auto')} />
                  Automatic
                </label>
                {selectedElement.kind !== 'series' && (
                  <label className="format-chart-radio">
                    <input type="radio" name="chart-border" checked={borderInfo.mode === 'none'} onChange={() => setBorder('none')} />
                    No line
                  </label>
                )}
                <label className="format-chart-radio">
                  <input type="radio" name="chart-border" checked={borderInfo.mode === 'solid'} onChange={() => setBorder('solid', borderInfo.color, borderInfo.width)} />
                  Solid line
                </label>
                <label className="format-chart-field-row">
                  <span>Color</span>
                  <input
                    type="color"
                    aria-label="Border color"
                    value={borderInfo.color.startsWith('#') ? borderInfo.color : '#353b45'}
                    onChange={event => setBorder('solid', event.currentTarget.value, borderInfo.width)}
                  />
                </label>
                <label className="format-chart-field-row">
                  <span>Width (px)</span>
                  <input
                    type="number"
                    min={0.25}
                    max={8}
                    step={0.25}
                    aria-label="Border width"
                    value={borderInfo.width}
                    onChange={event => {
                      const val = Number(event.currentTarget.value);
                      if (Number.isFinite(val) && val > 0) setBorder(borderInfo.mode, borderInfo.color, val, borderInfo.dash);
                    }}
                  />
                </label>
                <label className="format-chart-field-row">
                  <span>Dash type</span>
                  <select
                    aria-label="Dash type"
                    value={borderInfo.dash}
                    onChange={event => setBorder(borderInfo.mode, borderInfo.color, borderInfo.width, event.currentTarget.value as ChartLineDash)}
                  >
                    <option value="solid">Solid</option>
                    <option value="roundDot">Round Dot</option>
                    <option value="squareDot">Square Dot</option>
                    <option value="dash">Dash</option>
                    <option value="dashDot">Dash Dot</option>
                    <option value="longDash">Long Dash</option>
                    <option value="longDashDot">Long Dash Dot</option>
                    <option value="longDashDotDot">Long Dash Dot Dot</option>
                  </select>
                </label>
              </div>
            </details>
          </>
        )}

        {tab === 'sizeProps' && supportsSizeProps && (
          <>
            <details open className="format-chart-section">
              <summary>Size</summary>
              <div className="format-chart-group">
                <label className="format-chart-field-row">
                  <span>Height (px)</span>
                  <input
                    type="number"
                    min={90}
                    max={4096}
                    aria-label="Chart height"
                    value={Math.round(chart.placement.height)}
                    onChange={event => {
                      const v = event.currentTarget.valueAsNumber;
                      if (Number.isFinite(v) && v >= 90 && v <= 4096) onMoveResizeChart({ height: Math.round(v) });
                    }}
                  />
                </label>
                <label className="format-chart-field-row">
                  <span>Width (px)</span>
                  <input
                    type="number"
                    min={120}
                    max={4096}
                    aria-label="Chart width"
                    value={Math.round(chart.placement.width)}
                    onChange={event => {
                      const v = event.currentTarget.valueAsNumber;
                      if (Number.isFinite(v) && v >= 120 && v <= 4096) onMoveResizeChart({ width: Math.round(v) });
                    }}
                  />
                </label>
              </div>
            </details>

            <details open className="format-chart-section">
              <summary>Properties</summary>
              <div className="format-chart-group">
                {([
                  ['moveAndSize', 'Move and size with cells'],
                  ['moveOnly', "Move but don't size with cells"],
                  ['freeFloating', "Don't move or size with cells"],
                ] as readonly [ChartPlacementMode, string][]).map(([mode, label]) => (
                  <label key={mode} className="format-chart-radio">
                    <input
                      type="radio"
                      name="chart-placement-mode"
                      checked={placementMode === mode}
                      onChange={() => onMoveResizeChart({ mode })}
                    />
                    {label}
                  </label>
                ))}
              </div>
            </details>
          </>
        )}

        {tab === 'options' && selectedElement.kind === 'series' && selectedSeries && (
          <details open className="format-chart-section">
            <summary>Series Options</summary>
            <div className="format-chart-group">
              {(category === 'column' || category === 'bar' || category === 'line' || category === 'area' || category === 'combo') && (
                <label className="format-chart-check">
                  <input
                    type="checkbox"
                    checked={!!selectedSeries.secondaryAxis}
                    onChange={e => patchSelectedSeries(s => ({ ...s, secondaryAxis: e.currentTarget.checked }))}
                  />
                  Plot series on Secondary Axis
                </label>
              )}

              {(category === 'line' || category === 'scatter') && (
                <>
                  <label className="format-chart-check">
                    <input
                      type="checkbox"
                      checked={!!selectedSeries.smooth}
                      onChange={e => patchSelectedSeries(s => ({ ...s, smooth: e.currentTarget.checked }))}
                    />
                    Smoothed line
                  </label>
                  <label className="format-chart-check">
                    <input
                      type="checkbox"
                      checked={selectedSeries.markers !== false}
                      onChange={e => patchSelectedSeries(s => ({ ...s, markers: e.currentTarget.checked }))}
                    />
                    Show markers
                  </label>
                </>
              )}

              {category === 'pie' && (
                <>
                  <label className="format-chart-field-row">
                    <span>First slice angle (°)</span>
                    <input
                      type="number"
                      min={0}
                      max={360}
                      value={selectedSeries.firstSliceAngle ?? 0}
                      onChange={e => {
                        const v = e.currentTarget.valueAsNumber;
                        if (Number.isFinite(v)) patchSelectedSeries(s => ({ ...s, firstSliceAngle: Math.max(0, Math.min(360, v)) }));
                      }}
                    />
                  </label>
                  <label className="format-chart-field-row">
                    <span>Point explosion (%)</span>
                    <input
                      type="number"
                      min={0}
                      max={100}
                      value={selectedSeries.explosion ?? 0}
                      onChange={e => {
                        const v = e.currentTarget.valueAsNumber;
                        if (Number.isFinite(v)) patchSelectedSeries(s => ({ ...s, explosion: Math.max(0, Math.min(100, v)) }));
                      }}
                    />
                  </label>
                  {chart.chartType === 'doughnut' && (
                    <label className="format-chart-field-row">
                      <span>Doughnut hole (%)</span>
                      <input
                        type="number"
                        min={10}
                        max={90}
                        value={selectedSeries.holeSize ?? 50}
                        onChange={e => {
                          const v = e.currentTarget.valueAsNumber;
                          if (Number.isFinite(v)) patchSelectedSeries(s => ({ ...s, holeSize: Math.max(10, Math.min(90, v)) }));
                        }}
                      />
                    </label>
                  )}
                </>
              )}

              {(chart.chartType === 'histogram' || chart.chartType === 'pareto') && (
                <>
                  <label className="format-chart-field-row">
                    <span>Bins</span>
                    <select
                      value={selectedSeries.binMode ?? 'auto'}
                      onChange={e => patchSelectedSeries(s => ({ ...s, binMode: e.currentTarget.value as ChartBinMode }))}
                    >
                      <option value="auto">Automatic</option>
                      <option value="byCategory">By Category</option>
                      <option value="binWidth">Bin width</option>
                      <option value="binCount">Number of bins</option>
                    </select>
                  </label>
                  {selectedSeries.binMode === 'binWidth' && (
                    <label className="format-chart-field-row">
                      <span>Bin width</span>
                      <input
                        type="number"
                        min={0.01}
                        step="any"
                        value={selectedSeries.binWidth ?? 10}
                        onChange={e => {
                          const v = e.currentTarget.valueAsNumber;
                          if (Number.isFinite(v) && v > 0) patchSelectedSeries(s => ({ ...s, binWidth: v }));
                        }}
                      />
                    </label>
                  )}
                  {selectedSeries.binMode === 'binCount' && (
                    <label className="format-chart-field-row">
                      <span>Number of bins</span>
                      <input
                        type="number"
                        min={1}
                        max={100}
                        value={selectedSeries.binCount ?? 5}
                        onChange={e => {
                          const v = e.currentTarget.valueAsNumber;
                          if (Number.isFinite(v) && v >= 1) patchSelectedSeries(s => ({ ...s, binCount: Math.round(v) }));
                        }}
                      />
                    </label>
                  )}
                </>
              )}

              {chart.chartType === 'boxWhisker' && (
                <>
                  <label className="format-chart-field-row">
                    <span>Quartile</span>
                    <select
                      value={selectedSeries.quartileMethod ?? 'exclusive'}
                      onChange={e => patchSelectedSeries(s => ({ ...s, quartileMethod: e.currentTarget.value as ChartQuartileMethod }))}
                    >
                      <option value="exclusive">Exclusive median</option>
                      <option value="inclusive">Inclusive median</option>
                    </select>
                  </label>
                  <label className="format-chart-check">
                    <input
                      type="checkbox"
                      checked={!!selectedSeries.showInnerPoints}
                      onChange={e => patchSelectedSeries(s => ({ ...s, showInnerPoints: e.currentTarget.checked }))}
                    />
                    Show inner points
                  </label>
                  <label className="format-chart-check">
                    <input
                      type="checkbox"
                      checked={selectedSeries.showOutlierPoints !== false}
                      onChange={e => patchSelectedSeries(s => ({ ...s, showOutlierPoints: e.currentTarget.checked }))}
                    />
                    Show outlier points
                  </label>
                  <label className="format-chart-check">
                    <input
                      type="checkbox"
                      checked={selectedSeries.showMeanMarkers !== false}
                      onChange={e => patchSelectedSeries(s => ({ ...s, showMeanMarkers: e.currentTarget.checked }))}
                    />
                    Show mean markers
                  </label>
                </>
              )}

              {chart.chartType === 'waterfall' && (
                <>
                  <label className="format-chart-check">
                    <input
                      type="checkbox"
                      checked={selectedSeries.showConnectorLines !== false}
                      onChange={e => patchSelectedSeries(s => ({ ...s, showConnectorLines: e.currentTarget.checked }))}
                    />
                    Show connector lines
                  </label>
                  <div className="format-chart-sublist">
                    <span>Set point as Total:</span>
                    {data.categories.map((cat, idx) => {
                      const isTotal = selectedSeries.pointOverrides?.find(o => o.index === idx)?.isTotal ?? false;
                      return (
                        <label key={idx} className="format-chart-check">
                          <input
                            type="checkbox"
                            checked={isTotal}
                            onChange={e => {
                              const checked = e.currentTarget.checked;
                              const next = [...(selectedSeries.pointOverrides ?? []).filter(o => o.index !== idx)];
                              if (checked) next.push({ index: idx, isTotal: true });
                              patchSelectedSeries(s => ({ ...s, pointOverrides: next }));
                            }}
                          />
                          {cat}
                        </label>
                      );
                    })}
                  </div>
                </>
              )}
            </div>
          </details>
        )}

        {tab === 'options' && selectedElement.kind === 'axis' && (
          <>
            <details open className="format-chart-section">
              <summary>Axis Options</summary>
              <div className="format-chart-group">
                <label className="format-chart-check">
                  <input
                    type="checkbox"
                    checked={axisCfg.show !== false}
                    onChange={e => updateAxisCfg({ show: e.currentTarget.checked })}
                  />
                  Show axis
                </label>
                <label className="format-chart-field-row">
                  <span>Minimum</span>
                  <input
                    type="number"
                    placeholder="Auto"
                    aria-label="Axis minimum"
                    value={axisCfg.min ?? ''}
                    onChange={e => {
                      if (e.currentTarget.value === '') clearAxisField('min');
                      else if (Number.isFinite(e.currentTarget.valueAsNumber)) updateAxisCfg({ min: e.currentTarget.valueAsNumber });
                    }}
                  />
                </label>
                <label className="format-chart-field-row">
                  <span>Maximum</span>
                  <input
                    type="number"
                    placeholder="Auto"
                    aria-label="Axis maximum"
                    value={axisCfg.max ?? ''}
                    onChange={e => {
                      if (e.currentTarget.value === '') clearAxisField('max');
                      else if (Number.isFinite(e.currentTarget.valueAsNumber)) updateAxisCfg({ max: e.currentTarget.valueAsNumber });
                    }}
                  />
                </label>
                <label className="format-chart-field-row">
                  <span>Major unit</span>
                  <input
                    type="number"
                    min={0.0001}
                    step="any"
                    placeholder="Auto"
                    aria-label="Major unit"
                    value={axisCfg.majorUnit ?? ''}
                    onChange={e => {
                      if (e.currentTarget.value === '') clearAxisField('majorUnit');
                      else if (Number.isFinite(e.currentTarget.valueAsNumber) && e.currentTarget.valueAsNumber > 0) {
                        updateAxisCfg({ majorUnit: e.currentTarget.valueAsNumber });
                      }
                    }}
                  />
                </label>
                <label className="format-chart-field-row">
                  <span>Display units</span>
                  <select
                    value={axisCfg.displayUnits ?? 'none'}
                    onChange={e => updateAxisCfg({ displayUnits: e.currentTarget.value as ChartAxisDisplayUnits })}
                  >
                    <option value="none">None</option>
                    <option value="hundreds">Hundreds</option>
                    <option value="thousands">Thousands</option>
                    <option value="millions">Millions</option>
                    <option value="billions">Billions</option>
                  </select>
                </label>
                <label className="format-chart-check">
                  <input
                    type="checkbox"
                    checked={!!axisCfg.logarithmic}
                    onChange={e => updateAxisCfg({ logarithmic: e.currentTarget.checked })}
                  />
                  Logarithmic scale
                </label>
                <label className="format-chart-check">
                  <input
                    type="checkbox"
                    checked={!!axisCfg.reverseOrder}
                    onChange={e => updateAxisCfg({ reverseOrder: e.currentTarget.checked })}
                  />
                  Values in reverse order
                </label>
              </div>
            </details>

            <details open className="format-chart-section">
              <summary>Axis Title</summary>
              <div className="format-chart-group">
                <label className="format-chart-check">
                  <input
                    type="checkbox"
                    checked={!!axisCfg.title?.show}
                    onChange={e => updateAxisCfg({
                      title: { show: e.currentTarget.checked, text: axisCfg.title?.text || 'Axis Title' },
                    })}
                  />
                  Show axis title
                </label>
                <label className="format-chart-field-row">
                  <span>Title text</span>
                  <input
                    type="text"
                    aria-label="Axis title text"
                    value={axisCfg.title?.text ?? ''}
                    onChange={e => updateAxisCfg({ title: { show: true, text: e.currentTarget.value } })}
                  />
                </label>
              </div>
            </details>
          </>
        )}

        {tab === 'options' && selectedElement.kind === 'title' && (
          <details open className="format-chart-section">
            <summary>Title Options</summary>
            <div className="format-chart-group">
              <label className="format-chart-check">
                <input
                  type="checkbox"
                  checked={chart.elements.title?.show !== false}
                  onChange={e => onUpdateChart({
                    elements: {
                      ...chart.elements,
                      title: { ...chart.elements.title, show: e.currentTarget.checked, text: chart.elements.title?.text || 'Chart Title' },
                    },
                  })}
                />
                Show Chart Title
              </label>
              <label className="format-chart-field-row">
                <span>Text</span>
                <input
                  type="text"
                  aria-label="Chart title text"
                  value={chart.elements.title?.text ?? ''}
                  onChange={e => onUpdateChart({
                    elements: {
                      ...chart.elements,
                      title: { show: true, text: e.currentTarget.value },
                    },
                  })}
                />
              </label>
              <label className="format-chart-field-row">
                <span>Cell link</span>
                <input
                  type="text"
                  placeholder="=Sheet1!$A$1"
                  aria-label="Chart title formula link"
                  value={chart.elements.title?.formulaRef ?? ''}
                  onChange={e => {
                    const val = e.currentTarget.value;
                    onUpdateChart({
                      elements: {
                        ...chart.elements,
                        title: val.trim()
                          ? { show: true, formulaRef: val.trim().startsWith('=') ? val.trim() : `=${val.trim()}` }
                          : { show: true, text: chart.elements.title?.text || 'Chart Title' },
                      },
                    });
                  }}
                />
              </label>
            </div>
          </details>
        )}

        {tab === 'options' && selectedElement.kind === 'dataLabels' && (
          <details open className="format-chart-section">
            <summary>Label Options</summary>
            <div className="format-chart-group">
              <label className="format-chart-check">
                <input
                  type="checkbox"
                  checked={!!chart.elements.dataLabels?.show}
                  onChange={e => onUpdateChart({
                    elements: {
                      ...chart.elements,
                      dataLabels: { ...chart.elements.dataLabels, show: e.currentTarget.checked, showValue: chart.elements.dataLabels?.showValue ?? true },
                    },
                  })}
                />
                Show Data Labels
              </label>
              <label className="format-chart-check">
                <input
                  type="checkbox"
                  checked={!!chart.elements.dataLabels?.showSeriesName}
                  onChange={e => onUpdateChart({
                    elements: {
                      ...chart.elements,
                      dataLabels: { show: true, ...chart.elements.dataLabels, showSeriesName: e.currentTarget.checked },
                    },
                  })}
                />
                Series Name
              </label>
              <label className="format-chart-check">
                <input
                  type="checkbox"
                  checked={!!chart.elements.dataLabels?.showCategoryName}
                  onChange={e => onUpdateChart({
                    elements: {
                      ...chart.elements,
                      dataLabels: { show: true, ...chart.elements.dataLabels, showCategoryName: e.currentTarget.checked },
                    },
                  })}
                />
                Category Name
              </label>
              <label className="format-chart-check">
                <input
                  type="checkbox"
                  checked={chart.elements.dataLabels?.showValue !== false}
                  onChange={e => onUpdateChart({
                    elements: {
                      ...chart.elements,
                      dataLabels: { show: true, ...chart.elements.dataLabels, showValue: e.currentTarget.checked },
                    },
                  })}
                />
                Value
              </label>
              {category === 'pie' && (
                <label className="format-chart-check">
                  <input
                    type="checkbox"
                    checked={!!chart.elements.dataLabels?.showPercentage}
                    onChange={e => onUpdateChart({
                      elements: {
                        ...chart.elements,
                        dataLabels: { show: true, ...chart.elements.dataLabels, showPercentage: e.currentTarget.checked },
                      },
                    })}
                  />
                  Percentage
                </label>
              )}
              <label className="format-chart-field-row">
                <span>Position</span>
                <select
                  value={chart.elements.dataLabels?.position ?? 'outsideEnd'}
                  onChange={e => onUpdateChart({
                    elements: {
                      ...chart.elements,
                      dataLabels: { show: true, ...chart.elements.dataLabels, position: e.currentTarget.value as ChartDataLabelPosition },
                    },
                  })}
                >
                  <option value="center">Center</option>
                  <option value="insideEnd">Inside End</option>
                  <option value="insideBase">Inside Base</option>
                  <option value="outsideEnd">Outside End</option>
                  <option value="above">Above</option>
                  <option value="below">Below</option>
                </select>
              </label>
            </div>
          </details>
        )}

        {tab === 'options' && selectedElement.kind === 'legend' && (
          <details open className="format-chart-section">
            <summary>Legend Options</summary>
            <div className="format-chart-group">
              <label className="format-chart-check">
                <input
                  type="checkbox"
                  checked={chart.elements.legend?.show !== false}
                  onChange={e => onUpdateChart({
                    elements: {
                      ...chart.elements,
                      legend: { show: e.currentTarget.checked, position: chart.elements.legend?.position ?? 'bottom' },
                    },
                  })}
                />
                Show Legend
              </label>
              {([
                ['top', 'Top'],
                ['bottom', 'Bottom'],
                ['left', 'Left'],
                ['right', 'Right'],
              ] as readonly [ChartLegendPosition, string][]).map(([pos, label]) => (
                <label key={pos} className="format-chart-radio">
                  <input
                    type="radio"
                    name="legend-pos"
                    checked={(chart.elements.legend?.position ?? 'bottom') === pos}
                    onChange={() => onUpdateChart({
                      elements: { ...chart.elements, legend: { show: true, position: pos } },
                    })}
                  />
                  {label}
                </label>
              ))}
            </div>
          </details>
        )}

        {tab === 'options' && selectedElement.kind === 'trendline' && selectedSeries && selectedTrendline && (
          <details open className="format-chart-section">
            <summary>Trendline Options</summary>
            <div className="format-chart-group">
              {([
                ['exponential', 'Exponential'],
                ['linear', 'Linear'],
                ['logarithmic', 'Logarithmic'],
                ['polynomial', 'Polynomial'],
                ['power', 'Power'],
                ['movingAverage', 'Moving Average'],
              ] as readonly [ChartTrendlineType, string][]).map(([type, label]) => (
                <label key={type} className="format-chart-radio">
                  <input
                    type="radio"
                    name="trendline-type"
                    checked={selectedTrendline.type === type}
                    onChange={() => patchSelectedSeries(s => ({
                      ...s,
                      trendlines: (s.trendlines ?? []).map(t => t.id === selectedTrendline.id ? {
                        ...t,
                        type,
                        ...(type === 'polynomial' && { polynomialOrder: t.polynomialOrder ?? 2 }),
                        ...(type === 'movingAverage' && { period: t.period ?? 2 }),
                      } : t),
                    }))}
                  />
                  {label}
                </label>
              ))}

              {selectedTrendline.type === 'polynomial' && (
                <label className="format-chart-field-row">
                  <span>Order</span>
                  <input
                    type="number"
                    min={2}
                    max={6}
                    value={selectedTrendline.polynomialOrder ?? 2}
                    onChange={e => {
                      const v = Math.max(2, Math.min(6, Math.round(e.currentTarget.valueAsNumber || 2)));
                      patchSelectedSeries(s => ({
                        ...s,
                        trendlines: (s.trendlines ?? []).map(t => t.id === selectedTrendline.id ? { ...t, polynomialOrder: v } : t),
                      }));
                    }}
                  />
                </label>
              )}

              {selectedTrendline.type === 'movingAverage' && (
                <label className="format-chart-field-row">
                  <span>Period</span>
                  <input
                    type="number"
                    min={2}
                    max={12}
                    value={selectedTrendline.period ?? 2}
                    onChange={e => {
                      const v = Math.max(2, Math.round(e.currentTarget.valueAsNumber || 2));
                      patchSelectedSeries(s => ({
                        ...s,
                        trendlines: (s.trendlines ?? []).map(t => t.id === selectedTrendline.id ? { ...t, period: v } : t),
                      }));
                    }}
                  />
                </label>
              )}

              <label className="format-chart-field-row">
                <span>Forward</span>
                <input
                  type="number"
                  min={0}
                  step={0.5}
                  value={selectedTrendline.forward ?? 0}
                  onChange={e => {
                    const v = Math.max(0, e.currentTarget.valueAsNumber || 0);
                    patchSelectedSeries(s => ({
                      ...s,
                      trendlines: (s.trendlines ?? []).map(t => t.id === selectedTrendline.id ? { ...t, forward: v } : t),
                    }));
                  }}
                />
              </label>

              <label className="format-chart-check">
                <input
                  type="checkbox"
                  checked={!!selectedTrendline.displayEquation}
                  onChange={e => patchSelectedSeries(s => ({
                    ...s,
                    trendlines: (s.trendlines ?? []).map(t => t.id === selectedTrendline.id ? { ...t, displayEquation: e.currentTarget.checked } : t),
                  }))}
                />
                Display Equation on chart
              </label>
              <label className="format-chart-check">
                <input
                  type="checkbox"
                  checked={!!selectedTrendline.displayRSquared}
                  onChange={e => patchSelectedSeries(s => ({
                    ...s,
                    trendlines: (s.trendlines ?? []).map(t => t.id === selectedTrendline.id ? { ...t, displayRSquared: e.currentTarget.checked } : t),
                  }))}
                />
                Display R-squared value on chart
              </label>
              <button
                type="button"
                className="format-chart-delete-btn"
                onClick={() => {
                  patchSelectedSeries(s => {
                    const remaining = (s.trendlines ?? []).filter(t => t.id !== selectedTrendline.id);
                    const { trendlines: _prev, ...rest } = s;
                    return remaining.length > 0 ? { ...rest, trendlines: remaining } : rest;
                  });
                  onSelectElement({ kind: 'chartArea' });
                }}
              >
                Delete Trendline
              </button>
            </div>
          </details>
        )}
      </div>
    </aside>
  );
}
