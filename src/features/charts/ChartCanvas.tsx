import { useMemo, useState, type MouseEvent } from 'react';
import { chartCategoryOf, type ChartDataLabelsConfig, type ChartDocument } from '../../core/chart';
import type { ResolvedChartData } from '../../charts/chart-data';
import { computeChartLayout } from '../../charts/layout/advanced-layout';
import { getSeriesColor } from '../../charts/palettes';
import {
  dashArrayFor,
  elementMatches,
  getChartStylePreset,
  type ChartSelectedElement,
} from './chart-chrome';

interface TooltipState {
  readonly x: number;
  readonly y: number;
  readonly title: string;
  readonly detail: string;
}

function formatDataLabel(
  cfg: ChartDataLabelsConfig,
  opts: {
    readonly seriesName: string;
    readonly category: string;
    readonly formatted: string;
    readonly percent?: number;
  },
): string {
  const parts: string[] = [];
  if (cfg.showSeriesName) parts.push(opts.seriesName);
  if (cfg.showCategoryName) parts.push(opts.category);
  const wantVal = cfg.showValue !== false || (!cfg.showCategoryName && !cfg.showSeriesName && !cfg.showPercentage);
  if (wantVal) parts.push(opts.formatted);
  if (cfg.showPercentage && opts.percent !== undefined) {
    parts.push(`${Math.round(opts.percent * 100)}%`);
  }
  return parts.join(cfg.separator ?? ', ');
}

export function ChartCanvas({
  chart,
  data,
  width,
  height,
  svgRef,
  interactive = true,
  selectedElement,
  editingTitle = false,
  onSelectElement,
  onStartTitleEdit,
}: {
  readonly chart: ChartDocument;
  readonly data: ResolvedChartData;
  readonly width: number;
  readonly height: number;
  readonly svgRef?: ((node: SVGSVGElement | null) => void) | undefined;
  readonly interactive?: boolean;
  readonly selectedElement?: ChartSelectedElement | undefined;
  readonly editingTitle?: boolean;
  readonly onSelectElement?: ((el: ChartSelectedElement) => void) | undefined;
  readonly onStartTitleEdit?: (() => void) | undefined;
}) {
  const showDataTable = !!chart.elements.dataTable?.show && data.categories.length > 0 && data.series.some(s => !s.hidden);
  const dataTableRows = useMemo(
    () => (showDataTable ? data.series.filter(s => !s.hidden).slice(0, 6) : []),
    [showDataTable, data.series],
  );
  const dataTableHeight = showDataTable ? 20 + dataTableRows.length * 16 + 6 : 0;
  const effectiveLayoutHeight = Math.max(80, height - dataTableHeight);

  const layout = useMemo(
    () => computeChartLayout(data, chart, width, effectiveLayoutHeight),
    [data, chart, width, effectiveLayoutHeight],
  );
  const [tooltip, setTooltip] = useState<TooltipState | undefined>(undefined);

  const preset = getChartStylePreset(chart.styleId);
  const bgFillRaw = chart.chartAreaStyle?.fillRgb ?? preset.chartAreaBg;
  const bgFill = bgFillRaw === 'none' ? 'transparent' : bgFillRaw;
  const borderRaw = chart.chartAreaStyle?.borderRgb ?? preset.chartAreaBorder;
  const borderColor = borderRaw === 'none' ? 'transparent' : borderRaw;
  const borderWidth = chart.chartAreaStyle?.borderWidth ?? 1;
  const borderDash = dashArrayFor(chart.chartAreaStyle?.borderDash, borderWidth);

  const plotBgRaw = chart.plotAreaStyle?.fillRgb ?? preset.plotAreaBg;
  const plotBg = plotBgRaw === 'none' ? 'transparent' : plotBgRaw;
  const plotBorderRaw = chart.plotAreaStyle?.borderRgb ?? preset.plotAreaBorder;
  const plotBorder = plotBorderRaw && plotBorderRaw !== 'none' ? plotBorderRaw : 'transparent';
  const plotBorderWidth = chart.plotAreaStyle?.borderWidth ?? 1;
  const plotBorderDash = dashArrayFor(chart.plotAreaStyle?.borderDash, plotBorderWidth);

  const totalHeight = Math.max(1, layout.height + dataTableHeight);
  const category = chartCategoryOf(chart.chartType);
  const isBar = category === 'bar';

  // Gridline visibility
  const gridCfg = chart.elements.gridlines;
  const showPrimaryGrid = preset.showGridlines && (
    isBar
      ? gridCfg?.primaryMajorVertical !== false
      : gridCfg?.primaryMajorHorizontal !== false
  );
  const showSecondaryCrossGrid = isBar
    ? !!gridCfg?.primaryMajorHorizontal
    : !!gridCfg?.primaryMajorVertical;
  const showMinorH = !!gridCfg?.primaryMinorHorizontal;
  const showMinorV = !!gridCfg?.primaryMinorVertical;

  // Axis visibility & titles
  const showXAxis = isBar ? chart.elements.axes?.value?.show !== false : chart.elements.axes?.category?.show !== false;
  const showYAxis = isBar ? chart.elements.axes?.category?.show !== false : chart.elements.axes?.value?.show !== false;
  const xTitleCfg = isBar ? chart.elements.axes?.value?.title : chart.elements.axes?.category?.title;
  const yTitleCfg = isBar ? chart.elements.axes?.category?.title : chart.elements.axes?.value?.title;

  function pickElement(event: MouseEvent<SVGElement>, el: ChartSelectedElement) {
    if (!interactive || !onSelectElement) return;
    event.stopPropagation();
    onSelectElement(el);
  }

  const seriesAtIndex = (seriesIndex: number) => data.series[seriesIndex];
  const srcSeriesAtIndex = (seriesIndex: number) => {
    const res = data.series[seriesIndex];
    return res ? chart.series.find(s => s.id === res.id) : undefined;
  };
  const labelConfigFor = (seriesIndex: number): ChartDataLabelsConfig | undefined => {
    const src = srcSeriesAtIndex(seriesIndex);
    if (src?.dataLabels) return src.dataLabels.show ? src.dataLabels : undefined;
    if (chart.elements.dataLabels?.show) return chart.elements.dataLabels;
    if (preset.showDataLabels && chart.elements.dataLabels?.show !== false) {
      return { show: true, showValue: true, position: 'outsideEnd' };
    }
    return undefined;
  };

  return (
    <div className="chart-canvas-wrap">
      <svg
        ref={svgRef}
        className="chart-svg"
        viewBox={`0 0 ${layout.width} ${totalHeight}`}
        role="img"
        aria-label={`${chart.name}: ${layout.title?.text || chart.chartType}`}
        onPointerLeave={() => { if (interactive) setTooltip(undefined); }}
        onClick={event => pickElement(event, { kind: 'chartArea' })}
      >
        <rect
          x={0.5}
          y={0.5}
          width={Math.max(1, layout.width - 1)}
          height={Math.max(1, totalHeight - 1)}
          rx={4}
          fill={bgFill}
          stroke={borderColor}
          strokeWidth={borderWidth}
          strokeDasharray={borderDash}
          className={'chart-area-bg' + (elementMatches(selectedElement, { kind: 'chartArea' }) ? ' chart-el-selected' : '')}
        />

        <rect
          x={layout.plotArea.x}
          y={layout.plotArea.y}
          width={layout.plotArea.width}
          height={layout.plotArea.height}
          fill={plotBg}
          stroke={plotBorder !== 'transparent' ? plotBorder : undefined}
          strokeWidth={plotBorder !== 'transparent' ? plotBorderWidth : 0}
          strokeDasharray={plotBorder !== 'transparent' ? plotBorderDash : undefined}
          className={elementMatches(selectedElement, { kind: 'plotArea' }) ? 'chart-el-selected' : undefined}
          onClick={event => pickElement(event, { kind: 'plotArea' })}
        />

        {showPrimaryGrid && layout.gridlines.map((line, i) => (
          <line
            key={i}
            x1={line.x1}
            y1={line.y1}
            x2={line.x2}
            y2={line.y2}
            className="chart-gridline"
          />
        ))}

        {showSecondaryCrossGrid && (isBar ? layout.yAxis?.ticks : layout.xAxis?.ticks)?.map((t, i) => (
          <line
            key={`cross-grid-${i}`}
            x1={isBar ? layout.plotArea.x : t.x}
            y1={isBar ? t.y - 4 : layout.plotArea.y}
            x2={isBar ? layout.plotArea.x + layout.plotArea.width : t.x}
            y2={isBar ? t.y - 4 : layout.plotArea.y + layout.plotArea.height}
            className="chart-gridline"
          />
        ))}

        {showMinorH && layout.gridlines.slice(0, -1).map((line, i) => {
          const next = layout.gridlines[i + 1]!;
          if (Math.abs(line.y1 - line.y2) > 1) return null;
          const my = (line.y1 + next.y1) / 2;
          return (
            <line
              key={`minor-h-${i}`}
              x1={layout.plotArea.x}
              y1={my}
              x2={layout.plotArea.x + layout.plotArea.width}
              y2={my}
              className="chart-gridline chart-gridline-minor"
            />
          );
        })}

        {showMinorV && (() => {
          const ticks = isBar ? layout.xAxis?.ticks : layout.xAxis?.ticks;
          if (!ticks || ticks.length < 2) return null;
          return ticks.slice(0, -1).map((t, i) => {
            const next = ticks[i + 1]!;
            const mx = (t.x + next.x) / 2;
            return (
              <line
                key={`minor-v-${i}`}
                x1={mx}
                y1={layout.plotArea.y}
                x2={mx}
                y2={layout.plotArea.y + layout.plotArea.height}
                className="chart-gridline chart-gridline-minor"
              />
            );
          });
        })()}

        {layout.radarWeb && (
          <g className="chart-radar-web">
            {layout.radarWeb.rings.map((d, i) => (
              <path key={`r-${i}`} d={d} fill="none" className="chart-gridline" />
            ))}
            {layout.radarWeb.spokes.map((s, i) => (
              <g key={`s-${i}`}>
                <line x1={s.x1} y1={s.y1} x2={s.x2} y2={s.y2} className="chart-gridline" />
                <text x={s.labelX} y={s.labelY} textAnchor={s.anchor} className="chart-axis-label">
                  {s.label}
                </text>
              </g>
            ))}
            {layout.radarWeb.polygons.map((poly, i) => {
              const ser = seriesAtIndex(poly.markers[0]?.seriesIndex ?? i);
              return (
                <g
                  key={`p-${i}`}
                  onClick={event => { if (ser) pickElement(event, { kind: 'series', seriesId: ser.id }); }}
                >
                  <path
                    d={poly.path}
                    fill={poly.fill}
                    fillOpacity={poly.fillOpacity}
                    stroke={poly.stroke}
                    strokeWidth={2}
                  />
                  {poly.markers.map((m, mIdx) => (
                    <circle
                      key={mIdx}
                      className="chart-series-marker"
                      cx={m.x}
                      cy={m.y}
                      r={m.radius}
                      fill={m.fill}
                      stroke={m.stroke}
                      strokeWidth={1.5}
                      onPointerMove={interactive ? () => setTooltip({
                        x: m.x,
                        y: m.y,
                        title: m.seriesName,
                        detail: `${m.category}: ${m.formatted}`,
                      }) : undefined}
                    />
                  ))}
                </g>
              );
            })}
          </g>
        )}

        {layout.areas.map((area, i) => {
          const ser = seriesAtIndex(area.seriesIndex);
          return (
            <path
              key={i}
              d={area.path}
              fill={area.fill}
              fillOpacity={area.opacity}
              onClick={event => { if (ser) pickElement(event, { kind: 'series', seriesId: ser.id }); }}
            />
          );
        })}

        {layout.bars.map((bar, i) => {
          const ser = seriesAtIndex(bar.seriesIndex);
          const srcSer = srcSeriesAtIndex(bar.seriesIndex);
          const isSerSel = ser && elementMatches(selectedElement, { kind: 'series', seriesId: ser.id });
          const dlCfg = labelConfigFor(bar.seriesIndex);
          const barStroke = isSerSel ? '#ffffff' : (srcSer?.borderRgb ?? bar.stroke ?? (preset.outlineBars ? '#d8dee9' : undefined));
          const barStrokeW = isSerSel ? 1.5 : (srcSer?.borderRgb ? (srcSer.lineWidth ?? 1.5) : undefined);
          const barStrokeDash = dashArrayFor(srcSer?.lineDash, barStrokeW ?? 1);
          const dlY = dlCfg?.position === 'insideEnd'
            ? bar.y + 13
            : dlCfg?.position === 'center'
              ? bar.y + bar.height / 2 + 4
              : dlCfg?.position === 'insideBase'
                ? bar.y + bar.height - 4
                : isBar ? bar.y + bar.height / 2 + 4 : bar.y - 5;
          const dlX = isBar && dlCfg?.position !== 'center'
            ? bar.x + bar.width + 4
            : bar.x + bar.width / 2;
          const dlText = dlCfg ? formatDataLabel(dlCfg, {
            seriesName: bar.seriesName,
            category: bar.category,
            formatted: bar.formatted,
          }) : '';
          return (
            <g key={i}>
              <rect
                className={'chart-series-bar' + (isSerSel ? ' chart-el-selected' : '')}
                x={bar.x}
                y={bar.y}
                width={bar.width}
                height={bar.height}
                fill={bar.fill}
                stroke={barStroke}
                strokeWidth={barStrokeW}
                strokeDasharray={barStrokeDash}
                onClick={event => { if (ser) pickElement(event, { kind: 'series', seriesId: ser.id }); }}
                onPointerMove={interactive ? () => setTooltip({
                  x: bar.x + bar.width / 2,
                  y: bar.y,
                  title: bar.seriesName,
                  detail: `${bar.category}: ${bar.formatted}`,
                }) : undefined}
              />
              {dlCfg && bar.width > 4 && bar.height > 2 && (
                <>
                  {dlCfg.position === 'callout' && (
                    <rect
                      x={dlX - Math.max(18, dlText.length * 3.4) - 4}
                      y={dlY - 11}
                      width={Math.max(36, dlText.length * 6.8) + 8}
                      height={15}
                      rx={3}
                      className="chart-data-callout-box"
                    />
                  )}
                  <text
                    x={dlX}
                    y={dlY}
                    textAnchor={isBar && dlCfg.position !== 'center' && dlCfg.position !== 'callout' ? 'start' : 'middle'}
                    className="chart-data-label"
                    onClick={event => pickElement(event, { kind: 'dataLabels', ...(ser && { seriesId: ser.id }) })}
                  >
                    {dlText}
                  </text>
                </>
              )}
            </g>
          );
        })}

        {layout.waterfallConnectors?.map((conn, i) => (
          <line
            key={`wf-${i}`}
            x1={conn.x1}
            y1={conn.y1}
            x2={conn.x2}
            y2={conn.y2}
            className="chart-waterfall-connector"
          />
        ))}

        {layout.boxWhiskers?.map((bw, i) => {
          const cx = bw.x + bw.width / 2;
          const cap = Math.max(4, bw.width * 0.3);
          const ser = data.series.find(s => s.name === bw.seriesName);
          return (
            <g
              key={`bw-${i}`}
              className="chart-box-whisker"
              onClick={event => { if (ser) pickElement(event, { kind: 'series', seriesId: ser.id }); }}
              onPointerMove={interactive ? () => setTooltip({
                x: cx,
                y: bw.yBoxTop,
                title: bw.seriesName,
                detail: bw.summary,
              }) : undefined}
            >
              <line x1={cx} y1={bw.yWhiskerTop} x2={cx} y2={bw.yBoxTop} stroke={bw.stroke} strokeWidth={1.4} />
              <line x1={cx - cap} y1={bw.yWhiskerTop} x2={cx + cap} y2={bw.yWhiskerTop} stroke={bw.stroke} strokeWidth={1.4} />
              <line x1={cx} y1={bw.yBoxBottom} x2={cx} y2={bw.yWhiskerBottom} stroke={bw.stroke} strokeWidth={1.4} />
              <line x1={cx - cap} y1={bw.yBoxBottom} x2={cx + cap} y2={bw.yWhiskerBottom} stroke={bw.stroke} strokeWidth={1.4} />
              <rect
                x={bw.x}
                y={Math.min(bw.yBoxTop, bw.yBoxBottom)}
                width={bw.width}
                height={Math.max(2, Math.abs(bw.yBoxBottom - bw.yBoxTop))}
                fill={bw.fill}
                fillOpacity={0.78}
                stroke={bw.stroke}
                strokeWidth={1.3}
              />
              <line x1={bw.x} y1={bw.yMedian} x2={bw.x + bw.width} y2={bw.yMedian} stroke={bw.stroke} strokeWidth={1.8} />
              {bw.showMeanMarker && (
                <path
                  d={`M ${cx - 3.5} ${bw.yMean - 3.5} L ${cx + 3.5} ${bw.yMean + 3.5} M ${cx + 3.5} ${bw.yMean - 3.5} L ${cx - 3.5} ${bw.yMean + 3.5}`}
                  stroke={bw.stroke}
                  strokeWidth={1.5}
                />
              )}
              {bw.innerPoints.map((pt, pIdx) => (
                <circle key={`in-${pIdx}`} cx={pt.x} cy={pt.y} r={2.2} fill="#d8dee9" />
              ))}
              {bw.outliers.map((pt, oIdx) => (
                <circle key={`out-${oIdx}`} cx={pt.x} cy={pt.y} r={3} fill="none" stroke="#f38c4e" strokeWidth={1.5} />
              ))}
            </g>
          );
        })}

        {layout.treemapRects?.map((rect, i) => (
          <g
            key={`tm-${i}`}
            className="chart-series-bar"
            onPointerMove={interactive ? () => setTooltip({
              x: rect.x + rect.width / 2,
              y: rect.y + rect.height / 2,
              title: rect.label,
              detail: rect.formatted,
            }) : undefined}
          >
            <rect
              x={rect.x}
              y={rect.y}
              width={rect.width}
              height={rect.height}
              fill={rect.fill}
              stroke={rect.stroke}
              strokeWidth={1.5}
            />
            {rect.width > 42 && rect.height > 20 && (
              <text x={rect.x + 5} y={rect.y + 15} className="chart-treemap-label">
                {rect.label.length > Math.floor(rect.width / 7) ? rect.label.slice(0, Math.max(2, Math.floor(rect.width / 7) - 1)) + '…' : rect.label}
              </text>
            )}
          </g>
        ))}

        {layout.stockStems?.map((st, i) => (
          <g
            key={`st-${i}`}
            className="chart-stock-stem"
            onPointerMove={interactive ? () => setTooltip({
              x: st.x,
              y: st.yHigh,
              title: st.category,
              detail: st.summary,
            }) : undefined}
          >
            <line x1={st.x} y1={st.yHigh} x2={st.x} y2={st.yLow} stroke="#d8dee9" strokeWidth={1.5} />
            {st.yOpen !== undefined ? (
              <rect
                x={st.x - st.boxWidth / 2}
                y={Math.min(st.yOpen, st.yClose)}
                width={st.boxWidth}
                height={Math.max(2, Math.abs(st.yClose - st.yOpen))}
                fill={st.isDown ? '#f38c4e' : '#4ea1ff'}
                stroke="#181b20"
                strokeWidth={1}
              />
            ) : (
              <line x1={st.x} y1={st.yClose} x2={st.x + 6} y2={st.yClose} stroke="#4ea1ff" strokeWidth={2} />
            )}
          </g>
        ))}

        {layout.lines.map((series, i) => {
          const ser = seriesAtIndex(series.seriesIndex);
          const srcSer = srcSeriesAtIndex(series.seriesIndex);
          const dlCfg = labelConfigFor(series.seriesIndex);
          const lineDash = dashArrayFor(srcSer?.lineDash, srcSer?.lineWidth ?? series.width);
          return (
            <g
              key={i}
              onClick={event => { if (ser) pickElement(event, { kind: 'series', seriesId: ser.id }); }}
            >
              {series.path && (
                <path
                  d={series.path}
                  fill="none"
                  stroke={srcSer?.borderRgb ?? series.color}
                  strokeWidth={srcSer?.lineWidth ?? series.width}
                  strokeDasharray={lineDash}
                  strokeLinejoin="round"
                  strokeLinecap="round"
                />
              )}
              {series.markers.map((m, mIdx) => (
                <g key={mIdx}>
                  <circle
                    className="chart-series-marker"
                    cx={m.x}
                    cy={m.y}
                    r={m.radius}
                    fill={m.fill}
                    stroke={m.stroke}
                    strokeWidth={1.5}
                    onPointerMove={interactive ? () => setTooltip({
                      x: m.x,
                      y: m.y,
                      title: m.seriesName,
                      detail: `${m.category}: ${m.formatted}`,
                    }) : undefined}
                  />
                  {dlCfg && (
                    <text
                      x={m.x}
                      y={dlCfg.position === 'below' ? m.y + 14 : m.y - 8}
                      textAnchor="middle"
                      className="chart-data-label"
                      onClick={event => pickElement(event, { kind: 'dataLabels', ...(ser && { seriesId: ser.id }) })}
                    >
                      {formatDataLabel(dlCfg, {
                        seriesName: m.seriesName,
                        category: m.category,
                        formatted: m.formatted,
                      })}
                    </text>
                  )}
                </g>
              ))}
            </g>
          );
        })}

        {layout.errorBars?.map((eb, i) => (
          <g key={`eb-${i}`} className="chart-error-bar">
            <line x1={eb.x} y1={eb.yTop} x2={eb.x} y2={eb.yBottom} stroke={eb.color} strokeWidth={1.3} />
            {eb.capWidth > 0 && (
              <>
                <line x1={eb.x - eb.capWidth / 2} y1={eb.yTop} x2={eb.x + eb.capWidth / 2} y2={eb.yTop} stroke={eb.color} strokeWidth={1.3} />
                <line x1={eb.x - eb.capWidth / 2} y1={eb.yBottom} x2={eb.x + eb.capWidth / 2} y2={eb.yBottom} stroke={eb.color} strokeWidth={1.3} />
              </>
            )}
          </g>
        ))}

        {layout.trendlines?.map((tl, i) => {
          const ownerSeries = chart.series.find(s => (s.trendlines?.length ?? 0) > 0);
          const ownerTl = ownerSeries?.trendlines?.[0];
          return (
            <g
              key={`tl-${i}`}
              className="chart-trendline"
              onClick={event => {
                if (ownerSeries && ownerTl) {
                  pickElement(event, { kind: 'trendline', seriesId: ownerSeries.id, trendlineId: ownerTl.id });
                }
              }}
            >
              <path d={tl.path} fill="none" stroke={tl.color} strokeWidth={1.8} strokeDasharray="5 3" />
              {tl.label && tl.labelX !== undefined && tl.labelY !== undefined && (
                <text x={tl.labelX} y={tl.labelY} textAnchor="end" className="chart-axis-label">
                  {tl.label}
                </text>
              )}
            </g>
          );
        })}

        {layout.calloutConnectors.map((conn, i) => (
          <line
            key={i}
            x1={conn.x1}
            y1={conn.y1}
            x2={conn.x2}
            y2={conn.y2}
            className="chart-axis-line"
            strokeDasharray="3 2"
          />
        ))}

        {layout.slices.map((slice, i) => {
          const ser = seriesAtIndex(slice.seriesIndex);
          const dlCfg = labelConfigFor(slice.seriesIndex);
          return (
            <g key={i}>
              <path
                className="chart-series-slice"
                d={slice.path}
                fill={slice.fill}
                stroke={slice.stroke}
                strokeWidth={1.5}
                onClick={event => { if (ser) pickElement(event, { kind: 'series', seriesId: ser.id }); }}
                onPointerMove={interactive ? () => setTooltip({
                  x: slice.labelX,
                  y: slice.labelY,
                  title: slice.category,
                  detail: `${slice.formatted} (${Math.round(slice.percent * 100)}%)`,
                }) : undefined}
              />
              {dlCfg && slice.percent > 0.03 && (
                <text
                  x={slice.labelX}
                  y={slice.labelY + 4}
                  textAnchor="middle"
                  className="chart-data-label chart-data-label-pie"
                  onClick={event => pickElement(event, { kind: 'dataLabels', ...(ser && { seriesId: ser.id }) })}
                >
                  {formatDataLabel(dlCfg, {
                    seriesName: slice.seriesName,
                    category: slice.category,
                    formatted: slice.formatted,
                    percent: slice.percent,
                  })}
                </text>
              )}
            </g>
          );
        })}

        {layout.bubbles.map((b, i) => {
          const ser = seriesAtIndex(b.seriesIndex);
          return (
            <circle
              key={i}
              className="chart-series-marker"
              cx={b.x}
              cy={b.y}
              r={b.r}
              fill={b.fill}
              fillOpacity={0.72}
              stroke={b.stroke}
              strokeWidth={1.5}
              onClick={event => { if (ser) pickElement(event, { kind: 'series', seriesId: ser.id }); }}
              onPointerMove={interactive ? () => setTooltip({
                x: b.x,
                y: b.y - b.r,
                title: b.seriesName,
                detail: `${b.category} · Size: ${b.sizeVal}`,
              }) : undefined}
            />
          );
        })}

        {showXAxis && layout.xAxis && (
          <g
            className={'chart-axis' + (elementMatches(selectedElement, { kind: 'axis', axis: isBar ? 'value' : 'category' }) ? ' chart-el-selected' : '')}
            onClick={event => pickElement(event, { kind: 'axis', axis: isBar ? 'value' : 'category' })}
          >
            <line
              x1={layout.xAxis.line.x1}
              y1={layout.xAxis.line.y1}
              x2={layout.xAxis.line.x2}
              y2={layout.xAxis.line.y2}
              className="chart-axis-line"
            />
            {layout.xAxis.ticks.map((t, i) => (
              <text key={i} x={t.x} y={t.y} textAnchor={t.anchor} className="chart-axis-label">
                {t.label}
              </text>
            ))}
          </g>
        )}

        {xTitleCfg?.show && xTitleCfg.text && (
          <text
            x={layout.plotArea.x + layout.plotArea.width / 2}
            y={layout.plotArea.y + layout.plotArea.height + 27}
            textAnchor="middle"
            data-axis="x"
            className="chart-axis-title"
          >
            {xTitleCfg.text}
          </text>
        )}

        {showYAxis && layout.yAxis && (
          <g
            className={'chart-axis' + (elementMatches(selectedElement, { kind: 'axis', axis: isBar ? 'category' : 'value' }) ? ' chart-el-selected' : '')}
            onClick={event => pickElement(event, { kind: 'axis', axis: isBar ? 'category' : 'value' })}
          >
            <line
              x1={layout.yAxis.line.x1}
              y1={layout.yAxis.line.y1}
              x2={layout.yAxis.line.x2}
              y2={layout.yAxis.line.y2}
              className="chart-axis-line"
            />
            {layout.yAxis.ticks.map((t, i) => (
              <text key={i} x={t.x} y={t.y} textAnchor={t.anchor} className="chart-axis-label">
                {t.label}
              </text>
            ))}
          </g>
        )}

        {yTitleCfg?.show && yTitleCfg.text && (() => {
          const maxYTickChars = layout.yAxis && layout.yAxis.ticks.length > 0
            ? Math.max(2, ...layout.yAxis.ticks.map(t => t.label.length))
            : 4;
          const yTickSpace = Math.min(54, Math.ceil(maxYTickChars * 6.2) + 8);
          const yTitleX = Math.max(14, layout.plotArea.x - yTickSpace - 10);
          const yTitleY = layout.plotArea.y + layout.plotArea.height / 2;
          return (
            <text
              x={yTitleX}
              y={yTitleY}
              textAnchor="middle"
              transform={`rotate(-90 ${yTitleX} ${yTitleY})`}
              data-axis="y"
              className="chart-axis-title"
            >
              {yTitleCfg.text}
            </text>
          );
        })()}

        {layout.secondaryYAxis && (
          <g
            className={'chart-axis' + (elementMatches(selectedElement, { kind: 'axis', axis: 'secondaryValue' }) ? ' chart-el-selected' : '')}
            onClick={event => pickElement(event, { kind: 'axis', axis: 'secondaryValue' })}
          >
            <line
              x1={layout.secondaryYAxis.line.x1}
              y1={layout.secondaryYAxis.line.y1}
              x2={layout.secondaryYAxis.line.x2}
              y2={layout.secondaryYAxis.line.y2}
              className="chart-axis-line"
            />
            {layout.secondaryYAxis.ticks.map((t, i) => (
              <text key={i} x={t.x} y={t.y} textAnchor={t.anchor} className="chart-axis-label">
                {t.label}
              </text>
            ))}
          </g>
        )}

        {layout.title && !editingTitle && (() => {
          const titleY = chart.elements.title?.overlay ? layout.plotArea.y + 16 : layout.title.y;
          return (
            <g
              className={'chart-title-hit' + (elementMatches(selectedElement, { kind: 'title' }) ? ' chart-el-selected' : '')}
              onClick={event => pickElement(event, { kind: 'title' })}
              onDoubleClick={event => {
                if (!interactive || !onStartTitleEdit) return;
                event.stopPropagation();
                onStartTitleEdit();
              }}
            >
              <rect
                x={Math.max(4, layout.title.x - 100)}
                y={Math.max(2, titleY - 16)}
                width={Math.min(layout.width - 8, 200)}
                height={24}
                rx={3}
                fill="transparent"
              />
              <text
                x={layout.title.x}
                y={titleY}
                textAnchor={layout.title.anchor}
                fontSize={layout.title.fontSize}
                fontWeight={layout.title.bold ? 600 : 400}
                className="chart-title-text"
              >
                {layout.title.text}
              </text>
            </g>
          );
        })()}

        {layout.legend && layout.legend.map((item, i) => (
          <g
            key={i}
            className={'chart-legend-item' + (elementMatches(selectedElement, { kind: 'legend' }) ? ' chart-el-selected' : '')}
            onClick={event => pickElement(event, { kind: 'legend' })}
          >
            <rect x={item.x} y={item.y - 8} width={9} height={9} rx={1.5} fill={item.color} />
            <text x={item.x + 13} y={item.y} className="chart-legend-label">
              {item.label.length > 16 ? item.label.slice(0, 15) + '…' : item.label}
            </text>
          </g>
        ))}

        {showDataTable && dataTableRows.length > 0 && (() => {
          const tableY = layout.height - 2;
          const labelColW = Math.min(110, Math.max(58, layout.plotArea.x));
          const catCount = Math.max(1, data.categories.length);
          const colW = (layout.plotArea.x + layout.plotArea.width - labelColW) / catCount;
          const showKeys = chart.elements.dataTable?.showLegendKeys !== false;
          return (
            <g className="chart-data-table" aria-label="Chart data table">
              <rect
                x={labelColW}
                y={tableY}
                width={layout.plotArea.x + layout.plotArea.width - labelColW}
                height={18 + dataTableRows.length * 16}
                fill="none"
                className="chart-gridline"
              />
              <rect
                x={10}
                y={tableY + 18}
                width={Math.max(1, labelColW - 10)}
                height={dataTableRows.length * 16}
                fill="none"
                className="chart-gridline"
              />
              {data.categories.map((cat, cIdx) => {
                const cx = labelColW + (cIdx + 0.5) * colW;
                return (
                  <g key={`dt-h-${cIdx}`}>
                    <line
                      x1={labelColW + cIdx * colW}
                      y1={tableY}
                      x2={labelColW + cIdx * colW}
                      y2={tableY + 18 + dataTableRows.length * 16}
                      className="chart-gridline"
                    />
                    <text x={cx} y={tableY + 13} textAnchor="middle" className="chart-axis-label">
                      {cat.length > 8 ? cat.slice(0, 7) + '…' : cat}
                    </text>
                  </g>
                );
              })}
              {dataTableRows.map((ser, sIdx) => {
                const ry = tableY + 18 + sIdx * 16;
                const origIdx = data.series.findIndex(s => s.id === ser.id);
                const color = getSeriesColor(chart.paletteId, Math.max(0, origIdx), ser.sourceSeries.fillRgb);
                return (
                  <g key={`dt-r-${ser.id}`}>
                    <line
                      x1={10}
                      y1={ry}
                      x2={layout.plotArea.x + layout.plotArea.width}
                      y2={ry}
                      className="chart-gridline"
                    />
                    {showKeys && <rect x={14} y={ry + 4} width={7} height={7} rx={1} fill={color} />}
                    <text x={showKeys ? 25 : 14} y={ry + 12} className="chart-axis-label">
                      {ser.name.length > 10 ? ser.name.slice(0, 9) + '…' : ser.name}
                    </text>
                    {data.categories.map((_, cIdx) => (
                      <text
                        key={`dt-c-${cIdx}`}
                        x={labelColW + (cIdx + 0.5) * colW}
                        y={ry + 12}
                        textAnchor="middle"
                        className="chart-axis-label"
                      >
                        {ser.points[cIdx]?.formatted || '0'}
                      </text>
                    ))}
                  </g>
                );
              })}
            </g>
          );
        })()}

        {layout.isEmpty && (
          <text
            x={layout.width / 2}
            y={layout.height / 2}
            textAnchor="middle"
            className="chart-axis-label"
          >
            Select data with numeric values to plot
          </text>
        )}
      </svg>

      {interactive && tooltip && (
        <div
          className="chart-tooltip"
          style={{
            left: `${Math.min(82, Math.max(8, (tooltip.x / Math.max(1, layout.width)) * 100))}%`,
            top: `${Math.min(82, Math.max(8, (tooltip.y / Math.max(1, totalHeight)) * 100))}%`,
          }}
        >
          <strong>{tooltip.title}</strong>
          <span>{tooltip.detail}</span>
        </div>
      )}
    </div>
  );
}

function triggerDownload(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function downloadChartSvg(svg: SVGSVGElement, name: string): void {
  const clone = svg.cloneNode(true) as SVGSVGElement;
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  const serialized = new XMLSerializer().serializeToString(clone);
  triggerDownload(new Blob([serialized], { type: 'image/svg+xml;charset=utf-8' }), `${name}.svg`);
}

export function downloadChartPng(svg: SVGSVGElement, width: number, height: number, name: string): void {
  const clone = svg.cloneNode(true) as SVGSVGElement;
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  clone.setAttribute('width', String(width));
  clone.setAttribute('height', String(height));
  const serialized = new XMLSerializer().serializeToString(clone);
  const svgBlob = new Blob([serialized], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(svgBlob);
  const image = new Image();
  image.onload = () => {
    const canvas = document.createElement('canvas');
    canvas.width = width * 2;
    canvas.height = height * 2;
    const ctx = canvas.getContext('2d');
    if (ctx) {
      ctx.scale(2, 2);
      ctx.fillStyle = '#181b20';
      ctx.fillRect(0, 0, width, height);
      ctx.drawImage(image, 0, 0, width, height);
      canvas.toBlob(png => {
        if (png) triggerDownload(png, `${name}.png`);
        URL.revokeObjectURL(url);
      }, 'image/png');
    } else {
      URL.revokeObjectURL(url);
    }
  };
  image.onerror = () => URL.revokeObjectURL(url);
  image.src = url;
}
