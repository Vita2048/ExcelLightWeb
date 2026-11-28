import { chartCategoryOf, type ChartDocument } from '../../core/chart';
import { formatNumber } from '../../core/formats/cell-format';
import type { ResolvedChartData, ResolvedChartSeries } from '../chart-data';
import { getSeriesColor } from '../palettes';
import { computeNiceAxisScale, formatAxisTick } from './scales';

export interface ChartRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface ChartTextElement {
  readonly x: number;
  readonly y: number;
  readonly text: string;
  readonly anchor: 'start' | 'middle' | 'end';
  readonly fontSize: number;
  readonly bold?: boolean;
  readonly rotateDeg?: number;
}

export interface ChartBarRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly fill: string;
  readonly stroke?: string;
  readonly seriesIndex: number;
  readonly categoryIndex: number;
  readonly seriesName: string;
  readonly category: string;
  readonly value: number;
  readonly formatted: string;
}

export interface ChartMarkerPoint {
  readonly x: number;
  readonly y: number;
  readonly radius: number;
  readonly fill: string;
  readonly stroke: string;
  readonly seriesIndex: number;
  readonly categoryIndex: number;
  readonly seriesName: string;
  readonly category: string;
  readonly value: number;
  readonly formatted: string;
}

export interface ChartLineSeriesLayout {
  readonly path: string;
  readonly fillAreaPath?: string;
  readonly color: string;
  readonly width: number;
  readonly seriesIndex: number;
  readonly markers: readonly ChartMarkerPoint[];
}

export interface ChartPieSlice {
  readonly path: string;
  readonly fill: string;
  readonly stroke: string;
  readonly seriesIndex: number;
  readonly categoryIndex: number;
  readonly seriesName: string;
  readonly category: string;
  readonly value: number;
  readonly formatted: string;
  readonly percent: number;
  readonly centerX: number;
  readonly centerY: number;
  readonly labelX: number;
  readonly labelY: number;
}

export interface ChartBubblePoint {
  readonly x: number;
  readonly y: number;
  readonly r: number;
  readonly fill: string;
  readonly stroke: string;
  readonly seriesIndex: number;
  readonly categoryIndex: number;
  readonly seriesName: string;
  readonly category: string;
  readonly xVal: number;
  readonly yVal: number;
  readonly sizeVal: number;
  readonly formatted: string;
}

export interface ChartAxisTick {
  readonly x: number;
  readonly y: number;
  readonly label: string;
  readonly anchor: 'start' | 'middle' | 'end';
}

export interface ChartAxisLayout {
  readonly line: { readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number };
  readonly ticks: readonly ChartAxisTick[];
  readonly title?: ChartTextElement;
}

export interface ChartLegendItem {
  readonly label: string;
  readonly color: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly seriesIndex: number;
}

export interface ChartBoxWhiskerGeom {
  readonly x: number;
  readonly width: number;
  readonly yBoxTop: number;
  readonly yBoxBottom: number;
  readonly yMedian: number;
  readonly yWhiskerTop: number;
  readonly yWhiskerBottom: number;
  readonly yMean: number;
  readonly fill: string;
  readonly stroke: string;
  readonly showMeanMarker: boolean;
  readonly outliers: readonly { readonly x: number; readonly y: number; readonly value: number }[];
  readonly innerPoints: readonly { readonly x: number; readonly y: number; readonly value: number }[];
  readonly seriesName: string;
  readonly category: string;
  readonly summary: string;
}

export interface ChartTreemapRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly fill: string;
  readonly stroke: string;
  readonly label: string;
  readonly value: number;
  readonly formatted: string;
}

export interface ChartStockGeom {
  readonly x: number;
  readonly yHigh: number;
  readonly yLow: number;
  readonly yOpen?: number;
  readonly yClose: number;
  readonly boxWidth: number;
  readonly isDown: boolean;
  readonly category: string;
  readonly summary: string;
}

export interface ChartRadarWeb {
  readonly rings: readonly string[];
  readonly spokes: readonly {
    readonly x1: number;
    readonly y1: number;
    readonly x2: number;
    readonly y2: number;
    readonly labelX: number;
    readonly labelY: number;
    readonly anchor: 'start' | 'middle' | 'end';
    readonly label: string;
  }[];
  readonly polygons: readonly {
    readonly path: string;
    readonly stroke: string;
    readonly fill: string;
    readonly fillOpacity: number;
    readonly markers: readonly ChartMarkerPoint[];
  }[];
}

export interface ChartTrendlineGeom {
  readonly path: string;
  readonly color: string;
  readonly label?: string;
  readonly labelX?: number;
  readonly labelY?: number;
}

export interface ChartErrorBarGeom {
  readonly x: number;
  readonly yTop: number;
  readonly yBottom: number;
  readonly capWidth: number;
  readonly color: string;
}

export interface ChartLayoutResult {
  readonly width: number;
  readonly height: number;
  readonly plotArea: ChartRect;
  readonly title?: ChartTextElement;
  readonly legend?: readonly ChartLegendItem[];
  readonly gridlines: readonly { readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number }[];
  readonly xAxis?: ChartAxisLayout;
  readonly yAxis?: ChartAxisLayout;
  readonly secondaryYAxis?: ChartAxisLayout;
  readonly bars: readonly ChartBarRect[];
  readonly areas: readonly { readonly path: string; readonly fill: string; readonly opacity: number; readonly seriesIndex: number }[];
  readonly lines: readonly ChartLineSeriesLayout[];
  readonly slices: readonly ChartPieSlice[];
  readonly calloutConnectors: readonly { readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number }[];
  readonly waterfallConnectors?: readonly { readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number }[];
  readonly bubbles: readonly ChartBubblePoint[];
  readonly boxWhiskers?: readonly ChartBoxWhiskerGeom[];
  readonly treemapRects?: readonly ChartTreemapRect[];
  readonly stockStems?: readonly ChartStockGeom[];
  readonly radarWeb?: ChartRadarWeb;
  readonly trendlines?: readonly ChartTrendlineGeom[];
  readonly errorBars?: readonly ChartErrorBarGeom[];
  readonly isEmpty: boolean;
}

function isStackedType(type: ChartDocument['chartType']): boolean {
  return (
    type === 'stackedColumn' || type === 'stackedColumn100' ||
    type === 'stackedBar' || type === 'stackedBar100' ||
    type === 'stackedLine' || type === 'stackedLine100' ||
    type === 'stackedLineMarkers' || type === 'stackedLine100Markers' ||
    type === 'stackedArea' || type === 'stackedArea100'
  );
}

function isPercent100Type(type: ChartDocument['chartType']): boolean {
  return (
    type === 'stackedColumn100' ||
    type === 'stackedBar100' ||
    type === 'stackedLine100' ||
    type === 'stackedLine100Markers' ||
    type === 'stackedArea100'
  );
}

function hasMarkersType(type: ChartDocument['chartType']): boolean {
  return (
    type === 'lineMarkers' ||
    type === 'stackedLineMarkers' ||
    type === 'stackedLine100Markers' ||
    type === 'scatter' ||
    type === 'scatterSmoothLinesMarkers' ||
    type === 'scatterStraightLinesMarkers'
  );
}

function smoothPath(points: readonly { readonly x: number; readonly y: number }[]): string {
  if (points.length === 0) return '';
  if (points.length === 1) return `M ${points[0]!.x} ${points[0]!.y}`;
  if (points.length === 2) return `M ${points[0]!.x} ${points[0]!.y} L ${points[1]!.x} ${points[1]!.y}`;
  let d = `M ${points[0]!.x.toFixed(2)} ${points[0]!.y.toFixed(2)}`;
  for (let i = 0; i < points.length - 1; i++) {
    const p0 = points[Math.max(0, i - 1)]!;
    const p1 = points[i]!;
    const p2 = points[i + 1]!;
    const p3 = points[Math.min(points.length - 1, i + 2)]!;
    const cp1x = p1.x + (p2.x - p0.x) / 6;
    const cp1y = p1.y + (p2.y - p0.y) / 6;
    const cp2x = p2.x - (p3.x - p1.x) / 6;
    const cp2y = p2.y - (p3.y - p1.y) / 6;
    d += ` C ${cp1x.toFixed(2)} ${cp1y.toFixed(2)}, ${cp2x.toFixed(2)} ${cp2y.toFixed(2)}, ${p2.x.toFixed(2)} ${p2.y.toFixed(2)}`;
  }
  return d;
}

export function straightPath(points: readonly { readonly x: number; readonly y: number }[]): string {
  if (points.length === 0) return '';
  return points
    .map((pt, idx) => `${idx === 0 ? 'M' : 'L'} ${pt.x.toFixed(2)} ${pt.y.toFixed(2)}`)
    .join(' ');
}

export function arcSectorPath(
  cx: number,
  cy: number,
  outerR: number,
  innerR: number,
  startAngle: number,
  endAngle: number,
): string {
  const sweep = endAngle - startAngle;
  if (sweep >= Math.PI * 2 - 1e-4) {
    const mid = startAngle + Math.PI;
    return [
      arcSectorPath(cx, cy, outerR, innerR, startAngle, mid),
      arcSectorPath(cx, cy, outerR, innerR, mid, endAngle),
    ].join(' ');
  }
  const largeArc = sweep > Math.PI ? 1 : 0;
  const x1 = cx + outerR * Math.cos(startAngle);
  const y1 = cy + outerR * Math.sin(startAngle);
  const x2 = cx + outerR * Math.cos(endAngle);
  const y2 = cy + outerR * Math.sin(endAngle);
  if (innerR <= 0.5) {
    return `M ${cx.toFixed(2)} ${cy.toFixed(2)} L ${x1.toFixed(2)} ${y1.toFixed(2)} A ${outerR.toFixed(2)} ${outerR.toFixed(2)} 0 ${largeArc} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z`;
  }
  const ix1 = cx + innerR * Math.cos(endAngle);
  const iy1 = cy + innerR * Math.sin(endAngle);
  const ix2 = cx + innerR * Math.cos(startAngle);
  const iy2 = cy + innerR * Math.sin(startAngle);
  return `M ${x1.toFixed(2)} ${y1.toFixed(2)} A ${outerR.toFixed(2)} ${outerR.toFixed(2)} 0 ${largeArc} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} L ${ix1.toFixed(2)} ${iy1.toFixed(2)} A ${innerR.toFixed(2)} ${innerR.toFixed(2)} 0 ${largeArc} 0 ${ix2.toFixed(2)} ${iy2.toFixed(2)} Z`;
}

export function computeLegendItems(
  seriesList: readonly { readonly label: string; readonly color: string; readonly seriesIndex: number }[],
  chartWidth: number,
  chartHeight: number,
  position: NonNullable<ChartDocument['elements']['legend']>['position'],
): { readonly items: readonly ChartLegendItem[]; readonly padTop: number; readonly padBottom: number; readonly padRight: number; readonly padLeft: number } {
  if (seriesList.length === 0) {
    return { items: [], padTop: 0, padBottom: 0, padRight: 0, padLeft: 0 };
  }
  if (position === 'right' || position === 'overlayRight' || position === 'left') {
    const colWidth = 108;
    const x = position === 'left' ? 12 : Math.max(12, chartWidth - colWidth - 8);
    const startY = Math.max(38, chartHeight / 2 - (seriesList.length * 18) / 2);
    const items = seriesList.map((entry, idx) => ({
      ...entry,
      x,
      y: startY + idx * 18,
      width: colWidth,
    }));
    return {
      items,
      padTop: 0,
      padBottom: 0,
      padLeft: position === 'left' ? colWidth + 10 : 0,
      padRight: position === 'right' ? colWidth + 10 : 0,
    };
  }

  // Horizontal bottom / top legend
  const items: ChartLegendItem[] = [];
  const itemWidths = seriesList.map(s => Math.min(140, Math.max(54, s.label.length * 6.8 + 22)));
  const totalWidth = itemWidths.reduce((a, b) => a + b, 0);
  let cursorX = Math.max(14, (chartWidth - totalWidth) / 2);
  const y = position === 'top' ? 34 : chartHeight - 16;
  for (let i = 0; i < seriesList.length; i++) {
    const w = itemWidths[i]!;
    if (cursorX + w > chartWidth - 10 && i > 0) break;
    items.push({ ...seriesList[i]!, x: cursorX, y, width: w });
    cursorX += w + 8;
  }
  return {
    items,
    padTop: position === 'top' ? 22 : 0,
    padBottom: position === 'bottom' ? 26 : 0,
    padLeft: 0,
    padRight: 0,
  };
}

export function visibleSeriesList(data: ResolvedChartData): readonly { readonly series: ResolvedChartSeries; readonly index: number }[] {
  return data.series
    .map((series, index) => ({ series, index }))
    .filter(entry => !entry.series.hidden);
}

export function computeCore2DChartLayout(
  data: ResolvedChartData,
  chart: ChartDocument,
  width: number,
  height: number,
): ChartLayoutResult {
  const safeW = Math.max(120, width);
  const safeH = Math.max(90, height);
  const activeSeries = visibleSeriesList(data);
  const category = chartCategoryOf(chart.chartType);
  const titleCfg = chart.elements.title;
  const showTitle = !!titleCfg?.show && (data.title.trim().length > 0 || !!titleCfg.text);
  const titleText = data.title.trim() || titleCfg?.text || '';
  const titleOffset = showTitle ? 32 : 12;

  const legendCfg = chart.elements.legend;
  const showLegend = legendCfg?.show !== false && activeSeries.length > 0;
  const legendInputs = category === 'pie'
    ? data.categories.map((cat, idx) => ({
      label: cat,
      color: getSeriesColor(chart.paletteId, idx),
      seriesIndex: idx,
    }))
    : activeSeries.map(entry => ({
      label: entry.series.name,
      color: getSeriesColor(chart.paletteId, entry.index, entry.series.sourceSeries.fillRgb),
      seriesIndex: entry.index,
    }));

  const legendLayout = showLegend
    ? computeLegendItems(legendInputs, safeW, safeH, legendCfg?.position ?? 'bottom')
    : { items: [], padTop: 0, padBottom: 0, padLeft: 0, padRight: 0 };

  const title: ChartTextElement | undefined = showTitle
    ? { x: safeW / 2, y: 20, text: titleText, anchor: 'middle', fontSize: 13.5, bold: true }
    : undefined;

  if (activeSeries.length === 0 || (data.categories.length === 0 && category !== 'scatter')) {
    return {
      width: safeW,
      height: safeH,
      plotArea: { x: 24, y: titleOffset, width: Math.max(40, safeW - 48), height: Math.max(30, safeH - titleOffset - 24) },
      ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines: [],
      bars: [],
      areas: [],
      lines: [],
      slices: [],
      calloutConnectors: [],
      bubbles: [],
      isEmpty: true,
    };
  }

  // Pie / Doughnut / Pie of Pie / Bar of Pie
  if (category === 'pie') {
    const plotArea: ChartRect = {
      x: 16 + legendLayout.padLeft,
      y: titleOffset + legendLayout.padTop,
      width: Math.max(60, safeW - 32 - legendLayout.padLeft - legendLayout.padRight),
      height: Math.max(60, safeH - titleOffset - legendLayout.padTop - legendLayout.padBottom - 12),
    };
    const primary = activeSeries[0]!;
    const rawPoints = primary.series.points.map((pt, idx) => ({
      idx,
      category: data.categories[idx] ?? String(idx + 1),
      value: Math.max(0, pt.value ?? 0),
      formatted: pt.formatted || String(pt.value ?? 0),
    }));
    const total = rawPoints.reduce((acc, p) => acc + p.value, 0);
    if (total <= 0) {
      return {
        width: safeW, height: safeH, plotArea, ...(title && { title }),
        ...(showLegend && { legend: legendLayout.items }),
        gridlines: [], bars: [], areas: [], lines: [], slices: [], calloutConnectors: [], bubbles: [], isEmpty: true,
      };
    }

    const isCallout = chart.chartType === 'pieOfPie' || chart.chartType === 'barOfPie';
    const mainCx = isCallout ? plotArea.x + plotArea.width * 0.34 : plotArea.x + plotArea.width / 2;
    const mainCy = plotArea.y + plotArea.height / 2;
    const outerR = Math.max(24, (isCallout ? Math.min(plotArea.width * 0.3, plotArea.height * 0.44) : Math.min(plotArea.width, plotArea.height) * 0.43));
    const holePct = chart.chartType === 'doughnut' ? (primary.series.sourceSeries.holeSize ?? 50) / 100 : 0;
    const innerR = outerR * holePct;
    const startDeg = primary.series.sourceSeries.firstSliceAngle ?? 0;
    let angle = ((startDeg - 90) * Math.PI) / 180;

    const slices: ChartPieSlice[] = [];
    const bars: ChartBarRect[] = [];
    const calloutConnectors: { x1: number; y1: number; x2: number; y2: number }[] = [];

    if (isCallout && rawPoints.length > 2) {
      const splitCount = Math.max(1, Math.min(rawPoints.length - 1, Math.floor(rawPoints.length / 3)));
      const mainPoints = rawPoints.slice(0, rawPoints.length - splitCount);
      const otherPoints = rawPoints.slice(rawPoints.length - splitCount);
      const otherTotal = otherPoints.reduce((s, p) => s + p.value, 0);
      const combined = [
        ...mainPoints,
        { idx: -1, category: 'Other', value: otherTotal, formatted: formatNumber(otherTotal).text },
      ];
      // Center the "Other" wedge facing right (angle 0)
      const otherSweep = (otherTotal / total) * Math.PI * 2;
      angle = otherSweep / 2;
      let otherStart = 0;
      let otherEnd = 0;
      for (const pt of combined) {
        const sweep = (pt.value / total) * Math.PI * 2;
        const end = angle + sweep;
        const mid = (angle + end) / 2;
        if (pt.idx === -1) {
          // Place "Other" wedge centered around angle 0
          otherStart = -otherSweep / 2;
          otherEnd = otherSweep / 2;
          const midR = outerR * 0.65;
          slices.push({
            path: arcSectorPath(mainCx, mainCy, outerR, 0, otherStart, otherEnd),
            fill: '#676d77',
            stroke: '#181b20',
            seriesIndex: primary.index,
            categoryIndex: mainPoints.length,
            seriesName: primary.series.name,
            category: 'Other',
            value: pt.value,
            formatted: pt.formatted,
            percent: pt.value / total,
            centerX: mainCx,
            centerY: mainCy,
            labelX: mainCx + midR * Math.cos(0),
            labelY: mainCy + midR * Math.sin(0),
          });
        } else {
          const sliceColor = getSeriesColor(chart.paletteId, pt.idx);
          const midR = outerR * 0.65;
          slices.push({
            path: arcSectorPath(mainCx, mainCy, outerR, 0, angle, end),
            fill: sliceColor,
            stroke: '#181b20',
            seriesIndex: primary.index,
            categoryIndex: pt.idx,
            seriesName: primary.series.name,
            category: pt.category,
            value: pt.value,
            formatted: pt.formatted,
            percent: pt.value / total,
            centerX: mainCx,
            centerY: mainCy,
            labelX: mainCx + midR * Math.cos(mid),
            labelY: mainCy + midR * Math.sin(mid),
          });
          angle = end;
        }
      }

      const subCx = plotArea.x + plotArea.width * 0.78;
      const subR = outerR * 0.56;
      if (chart.chartType === 'pieOfPie') {
        let subAngle = -Math.PI / 2;
        for (const pt of otherPoints) {
          const sweep = otherTotal > 0 ? (pt.value / otherTotal) * Math.PI * 2 : 0;
          const end = subAngle + sweep;
          const mid = (subAngle + end) / 2;
          slices.push({
            path: arcSectorPath(subCx, mainCy, subR, 0, subAngle, end),
            fill: getSeriesColor(chart.paletteId, pt.idx),
            stroke: '#181b20',
            seriesIndex: primary.index,
            categoryIndex: pt.idx,
            seriesName: primary.series.name,
            category: pt.category,
            value: pt.value,
            formatted: pt.formatted,
            percent: pt.value / total,
            centerX: subCx,
            centerY: mainCy,
            labelX: subCx + subR * 0.6 * Math.cos(mid),
            labelY: mainCy + subR * 0.6 * Math.sin(mid),
          });
          subAngle = end;
        }
        calloutConnectors.push(
          { x1: mainCx + outerR * Math.cos(otherStart), y1: mainCy + outerR * Math.sin(otherStart), x2: subCx, y2: mainCy - subR },
          { x1: mainCx + outerR * Math.cos(otherEnd), y1: mainCy + outerR * Math.sin(otherEnd), x2: subCx, y2: mainCy + subR },
        );
      } else {
        const barW = Math.max(24, subR * 1.1);
        const barH = subR * 2;
        const barX = subCx - barW / 2;
        let curY = mainCy - barH / 2;
        for (const pt of otherPoints) {
          const h = otherTotal > 0 ? (pt.value / otherTotal) * barH : 0;
          bars.push({
            x: barX,
            y: curY,
            width: barW,
            height: h,
            fill: getSeriesColor(chart.paletteId, pt.idx),
            stroke: '#181b20',
            seriesIndex: primary.index,
            categoryIndex: pt.idx,
            seriesName: primary.series.name,
            category: pt.category,
            value: pt.value,
            formatted: pt.formatted,
          });
          curY += h;
        }
        calloutConnectors.push(
          { x1: mainCx + outerR * Math.cos(otherStart), y1: mainCy + outerR * Math.sin(otherStart), x2: barX, y2: mainCy - barH / 2 },
          { x1: mainCx + outerR * Math.cos(otherEnd), y1: mainCy + outerR * Math.sin(otherEnd), x2: barX, y2: mainCy + barH / 2 },
        );
      }
    } else {
      for (const pt of rawPoints) {
        const sweep = (pt.value / total) * Math.PI * 2;
        const end = angle + sweep;
        const mid = (angle + end) / 2;
        const explosion = (primary.series.sourceSeries.pointOverrides?.find(o => o.index === pt.idx)?.explosion ??
          primary.series.sourceSeries.explosion ?? 0) / 100 * outerR * 0.25;
        const scx = mainCx + explosion * Math.cos(mid);
        const scy = mainCy + explosion * Math.sin(mid);
        const labelR = (outerR + innerR) * 0.55;
        slices.push({
          path: arcSectorPath(scx, scy, outerR, innerR, angle, end),
          fill: getSeriesColor(chart.paletteId, pt.idx),
          stroke: '#181b20',
          seriesIndex: primary.index,
          categoryIndex: pt.idx,
          seriesName: primary.series.name,
          category: pt.category,
          value: pt.value,
          formatted: pt.formatted,
          percent: pt.value / total,
          centerX: scx,
          centerY: scy,
          labelX: scx + labelR * Math.cos(mid),
          labelY: scy + labelR * Math.sin(mid),
        });
        angle = end;
      }
    }

    return {
      width: safeW,
      height: safeH,
      plotArea,
      ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines: [],
      bars,
      areas: [],
      lines: [],
      slices,
      calloutConnectors,
      bubbles: [],
      isEmpty: false,
    };
  }

  // Cartesian plots (Column, Bar, Line, Area, Scatter, Bubble)
  const isBar = category === 'bar';
  const leftPad = (isBar ? 74 : 52) + legendLayout.padLeft;
  const rightPad = 18 + legendLayout.padRight;
  const topPad = titleOffset + legendLayout.padTop + 6;
  const bottomPad = 32 + legendLayout.padBottom;
  const plotArea: ChartRect = {
    x: leftPad,
    y: topPad,
    width: Math.max(48, safeW - leftPad - rightPad),
    height: Math.max(40, safeH - topPad - bottomPad),
  };

  // Scatter & Bubble
  if (category === 'scatter') {
    let minX = Infinity;
    let maxX = -Infinity;
    let minY = Infinity;
    let maxY = -Infinity;
    let maxSize = 0;

    const scatterPointsBySeries = activeSeries.map(entry => {
      const pts: { x: number; y: number; size: number; idx: number; formatted: string }[] = [];
      for (let i = 0; i < entry.series.points.length; i++) {
        const pt = entry.series.points[i]!;
        const xVal = pt.x ?? i + 1;
        const yVal = pt.value;
        if (yVal === null || !Number.isFinite(xVal) || !Number.isFinite(yVal)) continue;
        const sizeVal = Math.max(0.01, Math.abs(pt.size ?? 1));
        minX = Math.min(minX, xVal);
        maxX = Math.max(maxX, xVal);
        minY = Math.min(minY, yVal);
        maxY = Math.max(maxY, yVal);
        maxSize = Math.max(maxSize, sizeVal);
        pts.push({ x: xVal, y: yVal, size: sizeVal, idx: i, formatted: pt.formatted || String(yVal) });
      }
      return { ...entry, pts };
    });

    if (!Number.isFinite(minX)) {
      return {
        width: safeW, height: safeH, plotArea, ...(title && { title }),
        ...(showLegend && { legend: legendLayout.items }),
        gridlines: [], bars: [], areas: [], lines: [], slices: [], calloutConnectors: [], bubbles: [], isEmpty: true,
      };
    }

    const xCfg = chart.elements.axes?.category;
    const yCfg = chart.elements.axes?.value;
    const xScale = computeNiceAxisScale(minX, maxX, {
      includeZero: true,
      ...(xCfg?.min !== undefined && { min: xCfg.min }),
      ...(xCfg?.max !== undefined && { max: xCfg.max }),
      ...(xCfg?.majorUnit !== undefined && { majorUnit: xCfg.majorUnit }),
    });
    const yScale = computeNiceAxisScale(minY, maxY, {
      includeZero: true,
      ...(yCfg?.min !== undefined && { min: yCfg.min }),
      ...(yCfg?.max !== undefined && { max: yCfg.max }),
      ...(yCfg?.majorUnit !== undefined && { majorUnit: yCfg.majorUnit }),
    });

    const toPxX = (v: number) => plotArea.x + ((v - xScale.min) / Math.max(1e-9, xScale.max - xScale.min)) * plotArea.width;
    const toPxY = (v: number) => plotArea.y + plotArea.height - ((v - yScale.min) / Math.max(1e-9, yScale.max - yScale.min)) * plotArea.height;

    const gridlines: { x1: number; y1: number; x2: number; y2: number }[] = [];
    for (const tv of yScale.ticks) {
      const py = toPxY(tv);
      gridlines.push({ x1: plotArea.x, y1: py, x2: plotArea.x + plotArea.width, y2: py });
    }
    for (const tx of xScale.ticks) {
      const px = toPxX(tx);
      gridlines.push({ x1: px, y1: plotArea.y, x2: px, y2: plotArea.y + plotArea.height });
    }

    const lines: ChartLineSeriesLayout[] = [];
    const bubbles: ChartBubblePoint[] = [];
    const isBubble = chart.chartType === 'bubble';
    const drawLines = !isBubble && chart.chartType !== 'scatter';
    const isSmooth = chart.chartType === 'scatterSmoothLines' || chart.chartType === 'scatterSmoothLinesMarkers';
    const showMarkers = hasMarkersType(chart.chartType);

    for (const entry of scatterPointsBySeries) {
      const color = getSeriesColor(chart.paletteId, entry.index, entry.series.sourceSeries.fillRgb);
      if (isBubble) {
        for (const p of entry.pts) {
          const r = 6 + Math.sqrt(p.size / Math.max(1e-9, maxSize)) * Math.min(26, Math.min(plotArea.width, plotArea.height) * 0.14);
          bubbles.push({
            x: toPxX(p.x),
            y: toPxY(p.y),
            r,
            fill: color,
            stroke: '#181b20',
            seriesIndex: entry.index,
            categoryIndex: p.idx,
            seriesName: entry.series.name,
            category: `(${formatNumber(p.x).text}, ${formatNumber(p.y).text})`,
            xVal: p.x,
            yVal: p.y,
            sizeVal: p.size,
            formatted: p.formatted,
          });
        }
      } else {
        const pxPts = entry.pts.map(p => ({ ...p, px: toPxX(p.x), py: toPxY(p.y) }));
        const path = drawLines
          ? (isSmooth ? smoothPath(pxPts.map(p => ({ x: p.px, y: p.py }))) : straightPath(pxPts.map(p => ({ x: p.px, y: p.py }))))
          : '';
        lines.push({
          path,
          color,
          width: entry.series.sourceSeries.lineWidth ?? 2,
          seriesIndex: entry.index,
          markers: showMarkers
            ? pxPts.map(p => ({
              x: p.px,
              y: p.py,
              radius: 4,
              fill: color,
              stroke: '#181b20',
              seriesIndex: entry.index,
              categoryIndex: p.idx,
              seriesName: entry.series.name,
              category: formatNumber(p.x).text,
              value: p.y,
              formatted: p.formatted,
            }))
            : [],
        });
      }
    }

    return {
      width: safeW,
      height: safeH,
      plotArea,
      ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines,
      xAxis: {
        line: { x1: plotArea.x, y1: plotArea.y + plotArea.height, x2: plotArea.x + plotArea.width, y2: plotArea.y + plotArea.height },
        ticks: xScale.ticks.map(tx => ({
          x: toPxX(tx),
          y: plotArea.y + plotArea.height + 14,
          label: formatAxisTick(tx),
          anchor: 'middle',
        })),
      },
      yAxis: {
        line: { x1: plotArea.x, y1: plotArea.y, x2: plotArea.x, y2: plotArea.y + plotArea.height },
        ticks: yScale.ticks.map(ty => ({
          x: plotArea.x - 6,
          y: toPxY(ty) + 4,
          label: formatAxisTick(ty, activeSeries[0]?.series.valuesFormat),
          anchor: 'end',
        })),
      },
      bars: [],
      areas: [],
      lines,
      slices: [],
      calloutConnectors: [],
      bubbles,
      isEmpty: false,
    };
  }

  // Category-based Column, Bar, Line, Area
  const catCount = data.categories.length;
  const stacked = isStackedType(chart.chartType);
  const percent100 = isPercent100Type(chart.chartType);

  let minVal = Infinity;
  let maxVal = -Infinity;

  if (stacked) {
    for (let c = 0; c < catCount; c++) {
      let posSum = 0;
      let negSum = 0;
      let absSum = 0;
      for (const entry of activeSeries) {
        const v = entry.series.points[c]?.value ?? 0;
        if (v >= 0) posSum += v;
        else negSum += v;
        absSum += Math.abs(v);
      }
      if (percent100) {
        if (absSum > 0) {
          minVal = Math.min(minVal, negSum / absSum);
          maxVal = Math.max(maxVal, posSum / absSum);
        }
      } else {
        minVal = Math.min(minVal, negSum);
        maxVal = Math.max(maxVal, posSum);
      }
    }
  } else {
    for (const entry of activeSeries) {
      for (let c = 0; c < catCount; c++) {
        const v = entry.series.points[c]?.value;
        if (v === null || v === undefined) continue;
        minVal = Math.min(minVal, v);
        maxVal = Math.max(maxVal, v);
      }
    }
  }

  if (!Number.isFinite(minVal)) {
    minVal = 0;
    maxVal = 1;
  }

  const valAxisCfg = chart.elements.axes?.value;
  const valScale = computeNiceAxisScale(minVal, maxVal, {
    includeZero: category === 'column' || category === 'bar' || category === 'area',
    isPercent100: percent100,
    ...(valAxisCfg?.min !== undefined && { min: valAxisCfg.min }),
    ...(valAxisCfg?.max !== undefined && { max: valAxisCfg.max }),
    ...(valAxisCfg?.majorUnit !== undefined && { majorUnit: valAxisCfg.majorUnit }),
  });

  const valSpan = Math.max(1e-9, valScale.max - valScale.min);
  const gridlines: { x1: number; y1: number; x2: number; y2: number }[] = [];
  const bars: ChartBarRect[] = [];
  const areas: { path: string; fill: string; opacity: number; seriesIndex: number }[] = [];
  const lines: ChartLineSeriesLayout[] = [];

  if (isBar) {
    const valToX = (v: number) => plotArea.x + ((v - valScale.min) / valSpan) * plotArea.width;
    const baselineX = valToX(Math.max(valScale.min, Math.min(valScale.max, 0)));
    for (const tick of valScale.ticks) {
      const px = valToX(tick);
      gridlines.push({ x1: px, y1: plotArea.y, x2: px, y2: plotArea.y + plotArea.height });
    }

    const bandH = plotArea.height / catCount;
    const posStacks = new Array<number>(catCount).fill(0);
    const negStacks = new Array<number>(catCount).fill(0);
    const totals = new Array<number>(catCount).fill(0);
    if (percent100) {
      for (let c = 0; c < catCount; c++) {
        totals[c] = activeSeries.reduce((sum, s) => sum + Math.abs(s.series.points[c]?.value ?? 0), 0);
      }
    }

    for (let sIdx = 0; sIdx < activeSeries.length; sIdx++) {
      const entry = activeSeries[sIdx]!;
      const color = getSeriesColor(chart.paletteId, entry.index, entry.series.sourceSeries.fillRgb);
      for (let c = 0; c < catCount; c++) {
        const pt = entry.series.points[c];
        if (!pt || pt.value === null) continue;
        const rawV = pt.value;
        const plotted = percent100 ? (totals[c]! > 0 ? rawV / totals[c]! : 0) : rawV;
        if (stacked) {
          const prev = plotted >= 0 ? posStacks[c]! : negStacks[c]!;
          const next = prev + plotted;
          if (plotted >= 0) posStacks[c] = next;
          else negStacks[c] = next;
          const x1 = valToX(Math.min(prev, next));
          const x2 = valToX(Math.max(prev, next));
          const h = Math.max(4, bandH * 0.62);
          const y = plotArea.y + c * bandH + (bandH - h) / 2;
          bars.push({
            x: x1, y, width: Math.max(1, x2 - x1), height: h, fill: color,
            seriesIndex: entry.index, categoryIndex: c,
            seriesName: entry.series.name, category: data.categories[c] ?? '',
            value: rawV, formatted: pt.formatted || String(rawV),
          });
        } else {
          const groupPad = bandH * 0.16;
          const innerH = (bandH - groupPad * 2) / activeSeries.length;
          const h = Math.max(3, innerH * 0.84);
          const y = plotArea.y + c * bandH + groupPad + sIdx * innerH + (innerH - h) / 2;
          const vx = valToX(plotted);
          bars.push({
            x: Math.min(baselineX, vx),
            y,
            width: Math.max(1, Math.abs(vx - baselineX)),
            height: h,
            fill: color,
            seriesIndex: entry.index,
            categoryIndex: c,
            seriesName: entry.series.name,
            category: data.categories[c] ?? '',
            value: rawV,
            formatted: pt.formatted || String(rawV),
          });
        }
      }
    }

    return {
      width: safeW,
      height: safeH,
      plotArea,
      ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines,
      xAxis: {
        line: { x1: plotArea.x, y1: plotArea.y + plotArea.height, x2: plotArea.x + plotArea.width, y2: plotArea.y + plotArea.height },
        ticks: valScale.ticks.map(tick => ({
          x: valToX(tick),
          y: plotArea.y + plotArea.height + 14,
          label: formatAxisTick(tick, activeSeries[0]?.series.valuesFormat, percent100),
          anchor: 'middle',
        })),
      },
      yAxis: {
        line: { x1: baselineX, y1: plotArea.y, x2: baselineX, y2: plotArea.y + plotArea.height },
        ticks: data.categories.map((cat, idx) => ({
          x: plotArea.x - 6,
          y: plotArea.y + (idx + 0.5) * bandH + 4,
          label: cat.length > 12 ? cat.slice(0, 11) + '…' : cat,
          anchor: 'end',
        })),
      },
      bars,
      areas,
      lines,
      slices: [],
      calloutConnectors: [],
      bubbles: [],
      isEmpty: false,
    };
  }

  // Column, Line, Area
  const valToY = (v: number) => plotArea.y + plotArea.height - ((v - valScale.min) / valSpan) * plotArea.height;
  const baselineY = valToY(Math.max(valScale.min, Math.min(valScale.max, 0)));
  for (const tick of valScale.ticks) {
    const py = valToY(tick);
    gridlines.push({ x1: plotArea.x, y1: py, x2: plotArea.x + plotArea.width, y2: py });
  }

  const bandW = plotArea.width / catCount;
  const totals = new Array<number>(catCount).fill(0);
  if (percent100) {
    for (let c = 0; c < catCount; c++) {
      totals[c] = activeSeries.reduce((sum, s) => sum + Math.abs(s.series.points[c]?.value ?? 0), 0);
    }
  }

  if (category === 'column') {
    const posStacks = new Array<number>(catCount).fill(0);
    const negStacks = new Array<number>(catCount).fill(0);
    for (let sIdx = 0; sIdx < activeSeries.length; sIdx++) {
      const entry = activeSeries[sIdx]!;
      const color = getSeriesColor(chart.paletteId, entry.index, entry.series.sourceSeries.fillRgb);
      for (let c = 0; c < catCount; c++) {
        const pt = entry.series.points[c];
        if (!pt || pt.value === null) continue;
        const rawV = pt.value;
        const plotted = percent100 ? (totals[c]! > 0 ? rawV / totals[c]! : 0) : rawV;
        if (stacked) {
          const prev = plotted >= 0 ? posStacks[c]! : negStacks[c]!;
          const next = prev + plotted;
          if (plotted >= 0) posStacks[c] = next;
          else negStacks[c] = next;
          const y1 = valToY(Math.max(prev, next));
          const y2 = valToY(Math.min(prev, next));
          const w = Math.max(4, bandW * 0.58);
          const x = plotArea.x + c * bandW + (bandW - w) / 2;
          bars.push({
            x, y: y1, width: w, height: Math.max(1, y2 - y1), fill: color,
            seriesIndex: entry.index, categoryIndex: c,
            seriesName: entry.series.name, category: data.categories[c] ?? '',
            value: rawV, formatted: pt.formatted || String(rawV),
          });
        } else {
          const groupPad = bandW * 0.16;
          const innerW = (bandW - groupPad * 2) / activeSeries.length;
          const w = Math.max(3, innerW * 0.84);
          const x = plotArea.x + c * bandW + groupPad + sIdx * innerW + (innerW - w) / 2;
          const vy = valToY(plotted);
          bars.push({
            x,
            y: Math.min(baselineY, vy),
            width: w,
            height: Math.max(1, Math.abs(vy - baselineY)),
            fill: color,
            seriesIndex: entry.index,
            categoryIndex: c,
            seriesName: entry.series.name,
            category: data.categories[c] ?? '',
            value: rawV,
            formatted: pt.formatted || String(rawV),
          });
        }
      }
    }
  } else {
    // Line & Area
    const showMarkers = category === 'line' && hasMarkersType(chart.chartType);
    const prevTops = new Array<number>(catCount).fill(0);
    for (let sIdx = 0; sIdx < activeSeries.length; sIdx++) {
      const entry = activeSeries[sIdx]!;
      const color = getSeriesColor(chart.paletteId, entry.index, entry.series.sourceSeries.fillRgb);
      const segments: { x: number; y: number; baseY: number; idx: number; rawV: number; formatted: string }[][] = [];
      let currentSeg: { x: number; y: number; baseY: number; idx: number; rawV: number; formatted: string }[] = [];

      for (let c = 0; c < catCount; c++) {
        const pt = entry.series.points[c];
        const px = plotArea.x + (catCount === 1 ? plotArea.width / 2 : (c + 0.5) * bandW);
        if (!pt || pt.value === null) {
          if (!entry.series.connectGaps && currentSeg.length > 0) {
            segments.push(currentSeg);
            currentSeg = [];
          }
          continue;
        }
        const rawV = pt.value;
        const plotted = percent100 ? (totals[c]! > 0 ? rawV / totals[c]! : 0) : rawV;
        const topVal = stacked ? prevTops[c]! + plotted : plotted;
        const baseVal = stacked ? prevTops[c]! : 0;
        if (stacked) prevTops[c] = topVal;
        currentSeg.push({
          x: px,
          y: valToY(topVal),
          baseY: valToY(baseVal),
          idx: c,
          rawV,
          formatted: pt.formatted || String(rawV),
        });
      }
      if (currentSeg.length > 0) segments.push(currentSeg);

      const linePath = segments.map(seg => straightPath(seg)).join(' ');
      if (category === 'area') {
        for (const seg of segments) {
          if (seg.length === 0) continue;
          const upper = straightPath(seg);
          const lower = [...seg].reverse().map(p => `L ${p.x.toFixed(2)} ${p.baseY.toFixed(2)}`).join(' ');
          areas.push({
            path: `${upper} ${lower} Z`,
            fill: color,
            opacity: stacked ? 0.78 : 0.45,
            seriesIndex: entry.index,
          });
        }
      }

      const markers: ChartMarkerPoint[] = showMarkers
        ? segments.flatMap(seg => seg.map(p => ({
          x: p.x,
          y: p.y,
          radius: 4,
          fill: color,
          stroke: '#181b20',
          seriesIndex: entry.index,
          categoryIndex: p.idx,
          seriesName: entry.series.name,
          category: data.categories[p.idx] ?? '',
          value: p.rawV,
          formatted: p.formatted,
        })))
        : [];

      lines.push({
        path: linePath,
        color,
        width: entry.series.sourceSeries.lineWidth ?? 2,
        seriesIndex: entry.index,
        markers,
      });
    }
  }

  const labelStep = Math.max(1, Math.ceil(catCount / Math.max(2, Math.floor(plotArea.width / 54))));
  const xTicks: ChartAxisTick[] = [];
  for (let i = 0; i < catCount; i += labelStep) {
    const cat = data.categories[i] ?? '';
    xTicks.push({
      x: plotArea.x + (i + 0.5) * bandW,
      y: plotArea.y + plotArea.height + 14,
      label: cat.length > 12 ? cat.slice(0, 11) + '…' : cat,
      anchor: 'middle',
    });
  }

  return {
    width: safeW,
    height: safeH,
    plotArea,
    ...(title && { title }),
    ...(showLegend && { legend: legendLayout.items }),
    gridlines,
    xAxis: {
      line: { x1: plotArea.x, y1: baselineY, x2: plotArea.x + plotArea.width, y2: baselineY },
      ticks: xTicks,
    },
    yAxis: {
      line: { x1: plotArea.x, y1: plotArea.y, x2: plotArea.x, y2: plotArea.y + plotArea.height },
      ticks: valScale.ticks.map(tick => ({
        x: plotArea.x - 6,
        y: valToY(tick) + 4,
        label: formatAxisTick(tick, activeSeries[0]?.series.valuesFormat, percent100),
        anchor: 'end',
      })),
    },
    bars,
    areas,
    lines,
    slices: [],
    calloutConnectors: [],
    bubbles: [],
    isEmpty: false,
  };
}
