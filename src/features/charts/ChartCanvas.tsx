import { useMemo, useState } from 'react';
import type { ChartDocument } from '../../core/chart';
import type { ResolvedChartData } from '../../charts/chart-data';
import { computeChartLayout } from '../../charts/layout/advanced-layout';

interface TooltipState {
  readonly x: number;
  readonly y: number;
  readonly title: string;
  readonly detail: string;
}

export function ChartCanvas({
  chart,
  data,
  width,
  height,
  svgRef,
  interactive = true,
}: {
  readonly chart: ChartDocument;
  readonly data: ResolvedChartData;
  readonly width: number;
  readonly height: number;
  readonly svgRef?: ((node: SVGSVGElement | null) => void) | undefined;
  readonly interactive?: boolean;
}) {
  const layout = useMemo(
    () => computeChartLayout(data, chart, width, height),
    [data, chart, width, height],
  );
  const [tooltip, setTooltip] = useState<TooltipState | undefined>(undefined);

  const bgFill = chart.chartAreaStyle?.fillRgb ?? 'var(--chart-bg, #181b20)';
  const borderColor = chart.chartAreaStyle?.borderRgb ?? 'var(--border, #353b45)';
  const plotBg = chart.plotAreaStyle?.fillRgb ?? 'transparent';

  return (
    <div className="chart-canvas-wrap">
      <svg
        ref={svgRef}
        className="chart-svg"
        viewBox={`0 0 ${layout.width} ${layout.height}`}
        role="img"
        aria-label={`${chart.name}: ${layout.title?.text || chart.chartType}`}
        onPointerLeave={() => { if (interactive) setTooltip(undefined); }}
      >
        <rect
          x={0.5}
          y={0.5}
          width={Math.max(1, layout.width - 1)}
          height={Math.max(1, layout.height - 1)}
          rx={4}
          fill={bgFill}
          stroke={borderColor}
          className="chart-area-bg"
        />
        {plotBg !== 'transparent' && (
          <rect
            x={layout.plotArea.x}
            y={layout.plotArea.y}
            width={layout.plotArea.width}
            height={layout.plotArea.height}
            fill={plotBg}
          />
        )}

        {layout.gridlines.map((line, i) => (
          <line
            key={i}
            x1={line.x1}
            y1={line.y1}
            x2={line.x2}
            y2={line.y2}
            className="chart-gridline"
          />
        ))}

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
            {layout.radarWeb.polygons.map((poly, i) => (
              <g key={`p-${i}`}>
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
            ))}
          </g>
        )}

        {layout.areas.map((area, i) => (
          <path
            key={i}
            d={area.path}
            fill={area.fill}
            fillOpacity={area.opacity}
          />
        ))}

        {layout.bars.map((bar, i) => (
          <rect
            key={i}
            className="chart-series-bar"
            x={bar.x}
            y={bar.y}
            width={bar.width}
            height={bar.height}
            fill={bar.fill}
            stroke={bar.stroke}
            onPointerMove={interactive ? () => setTooltip({
              x: bar.x + bar.width / 2,
              y: bar.y,
              title: bar.seriesName,
              detail: `${bar.category}: ${bar.formatted}`,
            }) : undefined}
          />
        ))}

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
          return (
            <g
              key={`bw-${i}`}
              className="chart-box-whisker"
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
              <line x1={cx - cap} y1={bw.yWhiskerBottom} x2={cx + cap} y2={bw.yWhiskerBottom} stroke={bw.stroke} strokeWidth={1.4} />
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

        {layout.lines.map((series, i) => (
          <g key={i}>
            {series.path && (
              <path
                d={series.path}
                fill="none"
                stroke={series.color}
                strokeWidth={series.width}
                strokeLinejoin="round"
                strokeLinecap="round"
              />
            )}
            {series.markers.map((m, mIdx) => (
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
        ))}

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

        {layout.trendlines?.map((tl, i) => (
          <g key={`tl-${i}`} className="chart-trendline">
            <path d={tl.path} fill="none" stroke={tl.color} strokeWidth={1.6} strokeDasharray="5 3" />
            {tl.label && tl.labelX !== undefined && tl.labelY !== undefined && (
              <text x={tl.labelX} y={tl.labelY} textAnchor="end" className="chart-axis-label">
                {tl.label}
              </text>
            )}
          </g>
        ))}

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

        {layout.slices.map((slice, i) => (
          <path
            key={i}
            className="chart-series-slice"
            d={slice.path}
            fill={slice.fill}
            stroke={slice.stroke}
            strokeWidth={1.5}
            onPointerMove={interactive ? () => setTooltip({
              x: slice.labelX,
              y: slice.labelY,
              title: slice.category,
              detail: `${slice.formatted} (${Math.round(slice.percent * 100)}%)`,
            }) : undefined}
          />
        ))}

        {layout.bubbles.map((b, i) => (
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
            onPointerMove={interactive ? () => setTooltip({
              x: b.x,
              y: b.y - b.r,
              title: b.seriesName,
              detail: `${b.category} · Size: ${b.sizeVal}`,
            }) : undefined}
          />
        ))}

        {layout.xAxis && (
          <g className="chart-axis">
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

        {layout.yAxis && (
          <g className="chart-axis">
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

        {layout.secondaryYAxis && (
          <g className="chart-axis">
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

        {layout.title && (
          <text
            x={layout.title.x}
            y={layout.title.y}
            textAnchor={layout.title.anchor}
            fontSize={layout.title.fontSize}
            fontWeight={layout.title.bold ? 600 : 400}
            className="chart-title-text"
          >
            {layout.title.text}
          </text>
        )}

        {layout.legend && layout.legend.map((item, i) => (
          <g key={i} className="chart-legend-item">
            <rect x={item.x} y={item.y - 8} width={9} height={9} rx={1.5} fill={item.color} />
            <text x={item.x + 13} y={item.y} className="chart-legend-label">
              {item.label.length > 16 ? item.label.slice(0, 15) + '…' : item.label}
            </text>
          </g>
        ))}

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
            top: `${Math.min(82, Math.max(8, (tooltip.y / Math.max(1, layout.height)) * 100))}%`,
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
