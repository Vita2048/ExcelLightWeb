import { useMemo, useState } from 'react';
import type { ChartDocument } from '../../core/chart';
import type { ResolvedChartData } from '../../charts/chart-data';
import { computeCore2DChartLayout } from '../../charts/layout/core-2d-layout';

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
    () => computeCore2DChartLayout(data, chart, width, height),
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
