import { chartCategoryOf, type ChartDocument, type ChartType } from '../../core/chart';
import { formatNumber } from '../../core/formats/cell-format';
import {
  computeBoxWhiskerStats,
  computeErrorBars,
  computeHistogramBins,
  computeParetoBins,
  computeTrendline,
} from '../analytics';
import type { ResolvedChartData, ResolvedChartSeries } from '../chart-data';
import { getSeriesColor } from '../palettes';
import {
  axisTitlePadding,
  arcSectorPath,
  computeCore2DChartLayout,
  computeLegendItems,
  straightPath,
  visibleSeriesList,
  type ChartAxisTick,
  type ChartBarRect,
  type ChartBoxWhiskerGeom,
  type ChartErrorBarGeom,
  type ChartLayoutResult,
  type ChartLineSeriesLayout,
  type ChartMarkerPoint,
  type ChartPieSlice,
  type ChartRadarWeb,
  type ChartRect,
  type ChartStockGeom,
  type ChartTextElement,
  type ChartTreemapRect,
  type ChartTrendlineGeom,
} from './core-2d-layout';
import { computeNiceAxisScale, formatAxisTick } from './scales';

function emptyLayout(
  width: number,
  height: number,
  title: ChartTextElement | undefined,
  legend: ChartLayoutResult['legend'],
  plotArea: ChartRect,
): ChartLayoutResult {
  return {
    width,
    height,
    plotArea,
    ...(title && { title }),
    ...(legend && { legend }),
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

function squarifyTreemap(
  items: readonly { readonly label: string; readonly value: number; readonly formatted: string; readonly index: number }[],
  rect: ChartRect,
  paletteId: string,
): readonly ChartTreemapRect[] {
  const positive = items.filter(i => i.value > 0).sort((a, b) => b.value - a.value);
  const total = positive.reduce((s, i) => s + i.value, 0);
  if (total <= 0 || positive.length === 0) return [];

  const result: ChartTreemapRect[] = [];
  let rx = rect.x;
  let ry = rect.y;
  let rw = rect.width;
  let rh = rect.height;
  let remainingTotal = total;

  let i = 0;
  while (i < positive.length && rw > 1 && rh > 1) {
    const horizontal = rw >= rh;
    // Take a strip of items so aspect ratio stays balanced
    let stripSum = 0;
    const strip: typeof positive = [];
    const targetShare = remainingTotal * 0.42;
    while (i < positive.length) {
      const next = positive[i]!;
      strip.push(next);
      stripSum += next.value;
      i++;
      if (stripSum >= targetShare || strip.length >= 3) break;
    }

    const frac = stripSum / Math.max(1e-9, remainingTotal);
    if (horizontal) {
      const stripW = rw * frac;
      let curY = ry;
      for (const item of strip) {
        const h = rh * (item.value / Math.max(1e-9, stripSum));
        result.push({
          x: rx,
          y: curY,
          width: Math.max(1, stripW),
          height: Math.max(1, h),
          fill: getSeriesColor(paletteId, item.index),
          stroke: '#181b20',
          label: item.label,
          value: item.value,
          formatted: item.formatted,
        });
        curY += h;
      }
      rx += stripW;
      rw -= stripW;
    } else {
      const stripH = rh * frac;
      let curX = rx;
      for (const item of strip) {
        const w = rw * (item.value / Math.max(1e-9, stripSum));
        result.push({
          x: curX,
          y: ry,
          width: Math.max(1, w),
          height: Math.max(1, stripH),
          fill: getSeriesColor(paletteId, item.index),
          stroke: '#181b20',
          label: item.label,
          value: item.value,
          formatted: item.formatted,
        });
        curX += w;
      }
      ry += stripH;
      rh -= stripH;
    }
    remainingTotal -= stripSum;
  }

  return result;
}

function attachCartesianOverlays(
  base: ChartLayoutResult,
  data: ResolvedChartData,
  chart: ChartDocument,
): ChartLayoutResult {
  const activeSeries = visibleSeriesList(data);
  if (activeSeries.length === 0 || base.isEmpty) return base;
  const catCount = data.categories.length;
  const isScatter = chartCategoryOf(chart.chartType) === 'scatter';
  const hasAnyOverlay = activeSeries.some(
    s => (s.series.sourceSeries.trendlines?.length ?? 0) > 0 || s.series.sourceSeries.errorBars !== undefined,
  );
  if (!hasAnyOverlay) return base;

  const yAxis = base.yAxis;
  const xAxis = base.xAxis;
  if (!yAxis || !xAxis) return base;

  // Recover linear mapping from axis tick positions or min/max values
  const primaryValues: number[] = [];
  for (const entry of activeSeries) {
    if (entry.series.secondaryAxis) continue;
    for (const pt of entry.series.points) {
      if (pt.value !== null) primaryValues.push(pt.value);
    }
  }
  const minVal = primaryValues.length > 0 ? Math.min(...primaryValues) : 0;
  const maxVal = primaryValues.length > 0 ? Math.max(...primaryValues) : 1;
  const valScale = computeNiceAxisScale(minVal, maxVal, {
    includeZero: chartCategoryOf(chart.chartType) !== 'line',
    ...(chart.elements.axes?.value?.min !== undefined && { min: chart.elements.axes.value.min }),
    ...(chart.elements.axes?.value?.max !== undefined && { max: chart.elements.axes.value.max }),
    ...(chart.elements.axes?.value?.majorUnit !== undefined && { majorUnit: chart.elements.axes.value.majorUnit }),
  });
  const valSpan = Math.max(1e-9, valScale.max - valScale.min);
  const valToY = (v: number) =>
    base.plotArea.y + base.plotArea.height - ((v - valScale.min) / valSpan) * base.plotArea.height;

  const xValues: number[] = [];
  if (isScatter) {
    for (const entry of activeSeries) {
      for (let i = 0; i < entry.series.points.length; i++) {
        xValues.push(entry.series.points[i]?.x ?? i + 1);
      }
    }
  }
  const xScale = isScatter
    ? computeNiceAxisScale(Math.min(...xValues, 0), Math.max(...xValues, 1), { includeZero: true })
    : undefined;
  const xToPx = (xVal: number) => {
    if (isScatter && xScale) {
      return base.plotArea.x + ((xVal - xScale.min) / Math.max(1e-9, xScale.max - xScale.min)) * base.plotArea.width;
    }
    const bandW = base.plotArea.width / Math.max(1, catCount);
    return base.plotArea.x + (xVal - 0.5) * bandW;
  };

  const trendlines: ChartTrendlineGeom[] = [];
  const errorBars: ChartErrorBarGeom[] = [];

  for (const entry of activeSeries) {
    const color = getSeriesColor(chart.paletteId, entry.index, entry.series.sourceSeries.fillRgb);
    const pts = entry.series.points;

    for (const trend of entry.series.sourceSeries.trendlines ?? []) {
      const numericPts = pts
        .map((p, idx) => ({
          x: isScatter ? (p.x ?? idx + 1) : idx + 1,
          y: p.value,
        }))
        .filter((p): p is { x: number; y: number } => p.y !== null && Number.isFinite(p.x) && Number.isFinite(p.y));
      const res = computeTrendline(numericPts, trend);
      if (res.points.length >= 2) {
        const pxPts = res.points.map(p => ({ x: xToPx(p.x), y: valToY(p.y) }));
        const last = pxPts[pxPts.length - 1]!;
        const parts: string[] = [];
        if (trend.displayEquation && res.equationText) parts.push(res.equationText);
        if (trend.displayRSquared && res.rSquared !== undefined) parts.push(`R² = ${res.rSquared.toFixed(4)}`);
        trendlines.push({
          path: straightPath(pxPts),
          color: trend.colorRgb ?? color,
          ...(parts.length > 0 && {
            label: parts.join('   '),
            labelX: Math.min(base.plotArea.x + base.plotArea.width - 8, last.x),
            labelY: Math.max(base.plotArea.y + 12, last.y - 6),
          }),
        });
      }
    }

    if (entry.series.sourceSeries.errorBars) {
      const eb = entry.series.sourceSeries.errorBars;
      const deltas = computeErrorBars(pts.map(p => p.value), eb);
      for (let i = 0; i < pts.length; i++) {
        const v = pts[i]?.value;
        if (v === null || v === undefined) continue;
        const delta = deltas[i]!;
        const xVal = isScatter ? (pts[i]?.x ?? i + 1) : i + 1;
        const px = xToPx(xVal);
        errorBars.push({
          x: px,
          yTop: valToY(v + delta.plus),
          yBottom: valToY(v - delta.minus),
          capWidth: eb.cap === false ? 0 : 6,
          color: '#d8dee9',
        });
      }
    }
  }

  return {
    ...base,
    ...(trendlines.length > 0 && { trendlines }),
    ...(errorBars.length > 0 && { errorBars }),
  };
}

function buildHeaderAndLegend(
  data: ResolvedChartData,
  chart: ChartDocument,
  safeW: number,
  safeH: number,
  legendInputs: readonly { readonly label: string; readonly color: string; readonly seriesIndex: number }[],
) {
  const titleCfg = chart.elements.title;
  const showTitle = !!titleCfg?.show && (data.title.trim().length > 0 || !!titleCfg.text);
  const titleText = data.title.trim() || titleCfg?.text || '';
  const titleOffset = showTitle ? 32 : 12;
  const title: ChartTextElement | undefined = showTitle
    ? { x: safeW / 2, y: 20, text: titleText, anchor: 'middle', fontSize: 13.5, bold: true }
    : undefined;
  const legendCfg = chart.elements.legend;
  const showLegend = legendCfg?.show !== false && legendInputs.length > 0;
  const legendLayout = showLegend
    ? computeLegendItems(legendInputs, safeW, safeH, legendCfg?.position ?? 'bottom')
    : { items: [], padTop: 0, padBottom: 0, padLeft: 0, padRight: 0 };
  return { title, titleOffset, showLegend, legendLayout };
}

function computeComboLayout(
  data: ResolvedChartData,
  chart: ChartDocument,
  safeW: number,
  safeH: number,
): ChartLayoutResult {
  const activeSeries = visibleSeriesList(data);
  const { title, titleOffset, showLegend, legendLayout } = buildHeaderAndLegend(
    data,
    chart,
    safeW,
    safeH,
    activeSeries.map(e => ({
      label: e.series.name,
      color: getSeriesColor(chart.paletteId, e.index, e.series.sourceSeries.fillRgb),
      seriesIndex: e.index,
    })),
  );
  const hasSecondary = activeSeries.some(s => s.series.secondaryAxis);
  const comboTitlePad = axisTitlePadding(chart, false);
  const leftPad = 52 + legendLayout.padLeft + comboTitlePad.left;
  const rightPad = (hasSecondary ? 52 : 18) + legendLayout.padRight;
  const topPad = titleOffset + legendLayout.padTop + 6;
  const bottomPad = 32 + legendLayout.padBottom + comboTitlePad.bottom;
  const plotArea: ChartRect = {
    x: leftPad,
    y: topPad,
    width: Math.max(48, safeW - leftPad - rightPad),
    height: Math.max(40, safeH - topPad - bottomPad),
  };

  if (activeSeries.length === 0 || data.categories.length === 0) {
    return emptyLayout(safeW, safeH, title, showLegend ? legendLayout.items : undefined, plotArea);
  }

  const catCount = data.categories.length;
  const primarySeries = activeSeries.filter(s => !s.series.secondaryAxis);
  const secondarySeries = activeSeries.filter(s => s.series.secondaryAxis);

  const rangeFor = (list: readonly { readonly series: ResolvedChartSeries }[]) => {
    let lo = Infinity;
    let hi = -Infinity;
    for (const s of list) {
      for (const p of s.series.points) {
        if (p.value === null) continue;
        lo = Math.min(lo, p.value);
        hi = Math.max(hi, p.value);
      }
    }
    return Number.isFinite(lo) ? { lo, hi } : { lo: 0, hi: 1 };
  };

  const pRange = rangeFor(primarySeries.length > 0 ? primarySeries : activeSeries);
  const pScale = computeNiceAxisScale(pRange.lo, pRange.hi, {
    includeZero: true,
    ...(chart.elements.axes?.value?.min !== undefined && { min: chart.elements.axes.value.min }),
    ...(chart.elements.axes?.value?.max !== undefined && { max: chart.elements.axes.value.max }),
  });

  const sRange = hasSecondary ? rangeFor(secondarySeries) : pRange;
  const sScale = hasSecondary
    ? computeNiceAxisScale(sRange.lo, sRange.hi, {
      includeZero: true,
      ...(chart.elements.axes?.secondaryValue?.min !== undefined && { min: chart.elements.axes.secondaryValue.min }),
      ...(chart.elements.axes?.secondaryValue?.max !== undefined && { max: chart.elements.axes.secondaryValue.max }),
    })
    : pScale;

  const pToY = (v: number) =>
    plotArea.y + plotArea.height - ((v - pScale.min) / Math.max(1e-9, pScale.max - pScale.min)) * plotArea.height;
  const sToY = (v: number) =>
    plotArea.y + plotArea.height - ((v - sScale.min) / Math.max(1e-9, sScale.max - sScale.min)) * plotArea.height;

  const gridlines = pScale.ticks.map(t => ({
    x1: plotArea.x,
    y1: pToY(t),
    x2: plotArea.x + plotArea.width,
    y2: pToY(t),
  }));

  const bandW = plotArea.width / catCount;
  const bars: ChartBarRect[] = [];
  const areas: { path: string; fill: string; opacity: number; seriesIndex: number }[] = [];
  const lines: ChartLineSeriesLayout[] = [];

  const columnEntries = activeSeries.filter(s => {
    const t: ChartType = s.series.chartType === 'combo' ? 'clusteredColumn' : s.series.chartType;
    return t === 'clusteredColumn' || t === 'stackedColumn' || t === 'stackedColumn100';
  });

  for (let sIdx = 0; sIdx < activeSeries.length; sIdx++) {
    const entry = activeSeries[sIdx]!;
    const subType: ChartType = entry.series.chartType === 'combo'
      ? (sIdx < Math.ceil(activeSeries.length / 2) ? 'clusteredColumn' : 'line')
      : entry.series.chartType;
    const color = getSeriesColor(chart.paletteId, entry.index, entry.series.sourceSeries.fillRgb);
    const valToY = entry.series.secondaryAxis ? sToY : pToY;
    const scaleUsed = entry.series.secondaryAxis ? sScale : pScale;
    const baselineY = valToY(Math.max(scaleUsed.min, Math.min(scaleUsed.max, 0)));

    if (subType === 'clusteredColumn' || subType === 'stackedColumn' || subType === 'stackedColumn100') {
      const colPos = Math.max(0, columnEntries.indexOf(entry));
      const colCount = Math.max(1, columnEntries.length);
      const groupPad = bandW * 0.16;
      const innerW = (bandW - groupPad * 2) / colCount;
      const w = Math.max(3, innerW * 0.84);
      for (let c = 0; c < catCount; c++) {
        const pt = entry.series.points[c];
        if (!pt || pt.value === null) continue;
        const x = plotArea.x + c * bandW + groupPad + colPos * innerW + (innerW - w) / 2;
        const vy = valToY(pt.value);
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
          value: pt.value,
          formatted: pt.formatted || String(pt.value),
        });
      }
    } else {
      const showMarkers = subType.includes('Markers');
      const pxPts: { x: number; y: number; idx: number; rawV: number; formatted: string }[] = [];
      for (let c = 0; c < catCount; c++) {
        const pt = entry.series.points[c];
        if (!pt || pt.value === null) continue;
        pxPts.push({
          x: plotArea.x + (c + 0.5) * bandW,
          y: valToY(pt.value),
          idx: c,
          rawV: pt.value,
          formatted: pt.formatted || String(pt.value),
        });
      }
      const linePath = straightPath(pxPts);
      if (subType === 'area' || subType === 'stackedArea' || subType === 'stackedArea100') {
        if (pxPts.length > 0) {
          areas.push({
            path: `${linePath} L ${pxPts[pxPts.length - 1]!.x.toFixed(2)} ${baselineY.toFixed(2)} L ${pxPts[0]!.x.toFixed(2)} ${baselineY.toFixed(2)} Z`,
            fill: color,
            opacity: 0.45,
            seriesIndex: entry.index,
          });
        }
      }
      lines.push({
        path: linePath,
        color,
        width: entry.series.sourceSeries.lineWidth ?? 2,
        seriesIndex: entry.index,
        markers: showMarkers
          ? pxPts.map(p => ({
            x: p.x, y: p.y, radius: 4, fill: color, stroke: '#181b20',
            seriesIndex: entry.index, categoryIndex: p.idx,
            seriesName: entry.series.name, category: data.categories[p.idx] ?? '',
            value: p.rawV, formatted: p.formatted,
          }))
          : [],
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

  return attachCartesianOverlays(
    {
      width: safeW,
      height: safeH,
      plotArea,
      ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines,
      xAxis: {
        line: { x1: plotArea.x, y1: pToY(Math.max(pScale.min, Math.min(pScale.max, 0))), x2: plotArea.x + plotArea.width, y2: pToY(Math.max(pScale.min, Math.min(pScale.max, 0))) },
        ticks: xTicks,
      },
      yAxis: {
        line: { x1: plotArea.x, y1: plotArea.y, x2: plotArea.x, y2: plotArea.y + plotArea.height },
        ticks: pScale.ticks.map(t => ({
          x: plotArea.x - 6,
          y: pToY(t) + 4,
          label: formatAxisTick(t, primarySeries[0]?.series.valuesFormat),
          anchor: 'end',
        })),
      },
      ...(hasSecondary && {
        secondaryYAxis: {
          line: { x1: plotArea.x + plotArea.width, y1: plotArea.y, x2: plotArea.x + plotArea.width, y2: plotArea.y + plotArea.height },
          ticks: sScale.ticks.map(t => ({
            x: plotArea.x + plotArea.width + 6,
            y: sToY(t) + 4,
            label: formatAxisTick(t, secondarySeries[0]?.series.valuesFormat),
            anchor: 'start',
          })),
        },
      }),
      bars,
      areas,
      lines,
      slices: [],
      calloutConnectors: [],
      bubbles: [],
      isEmpty: false,
    },
    data,
    chart,
  );
}

export function computeChartLayout(
  data: ResolvedChartData,
  chart: ChartDocument,
  width: number,
  height: number,
): ChartLayoutResult {
  const safeW = Math.max(120, width);
  const safeH = Math.max(90, height);
  const cat = chartCategoryOf(chart.chartType);

  if (
    cat === 'column' || cat === 'bar' || cat === 'line' ||
    cat === 'area' || cat === 'pie' || cat === 'scatter'
  ) {
    const core = computeCore2DChartLayout(data, chart, safeW, safeH);
    return cat === 'pie' ? core : attachCartesianOverlays(core, data, chart);
  }

  if (cat === 'combo') {
    return computeComboLayout(data, chart, safeW, safeH);
  }

  const activeSeries = visibleSeriesList(data);

  // Waterfall
  if (chart.chartType === 'waterfall') {
    const legendInputs = [
      { label: 'Increase', color: '#4ea1ff', seriesIndex: 0 },
      { label: 'Decrease', color: '#f38c4e', seriesIndex: 1 },
      { label: 'Total', color: '#8c95a3', seriesIndex: 2 },
    ];
    const { title, titleOffset, showLegend, legendLayout } = buildHeaderAndLegend(data, chart, safeW, safeH, legendInputs);
    const wfTitlePad = axisTitlePadding(chart, false);
    const leftPad = 52 + legendLayout.padLeft + wfTitlePad.left;
    const rightPad = 18 + legendLayout.padRight;
    const topPad = titleOffset + legendLayout.padTop + 6;
    const bottomPad = 32 + legendLayout.padBottom + wfTitlePad.bottom;
    const plotArea: ChartRect = {
      x: leftPad, y: topPad,
      width: Math.max(48, safeW - leftPad - rightPad),
      height: Math.max(40, safeH - topPad - bottomPad),
    };
    const primary = activeSeries[0];
    if (!primary || data.categories.length === 0) {
      return emptyLayout(safeW, safeH, title, showLegend ? legendLayout.items : undefined, plotArea);
    }

    const overrides = new Set(
      (primary.series.sourceSeries.pointOverrides ?? []).filter(o => o.isTotal).map(o => o.index),
    );
    const steps: { start: number; end: number; value: number; isTotal: boolean; formatted: string; category: string; idx: number }[] = [];
    let running = 0;
    let minV = 0;
    let maxV = 0;
    for (let i = 0; i < data.categories.length; i++) {
      const pt = primary.series.points[i];
      const val = pt?.value ?? 0;
      const isTotal = overrides.has(i) || (overrides.size === 0 && i === data.categories.length - 1 && /total/iu.test(data.categories[i] ?? ''));
      const start = isTotal ? 0 : running;
      const end = isTotal ? val : running + val;
      if (!isTotal) running = end;
      else running = end;
      minV = Math.min(minV, start, end);
      maxV = Math.max(maxV, start, end);
      steps.push({
        start,
        end,
        value: val,
        isTotal,
        formatted: pt?.formatted || String(val),
        category: data.categories[i] ?? String(i + 1),
        idx: i,
      });
    }

    const scale = computeNiceAxisScale(minV, maxV, { includeZero: true });
    const valToY = (v: number) =>
      plotArea.y + plotArea.height - ((v - scale.min) / Math.max(1e-9, scale.max - scale.min)) * plotArea.height;
    const bandW = plotArea.width / steps.length;
    const barW = Math.max(6, bandW * 0.58);
    const bars: ChartBarRect[] = [];
    const connectors: { x1: number; y1: number; x2: number; y2: number }[] = [];

    for (let i = 0; i < steps.length; i++) {
      const s = steps[i]!;
      const x = plotArea.x + i * bandW + (bandW - barW) / 2;
      const y1 = valToY(Math.max(s.start, s.end));
      const y2 = valToY(Math.min(s.start, s.end));
      const fill = s.isTotal ? '#8c95a3' : s.value >= 0 ? '#4ea1ff' : '#f38c4e';
      bars.push({
        x, y: y1, width: barW, height: Math.max(2, y2 - y1), fill,
        seriesIndex: primary.index, categoryIndex: s.idx,
        seriesName: s.isTotal ? 'Total' : s.value >= 0 ? 'Increase' : 'Decrease',
        category: s.category, value: s.value, formatted: s.formatted,
      });
      if (i + 1 < steps.length && primary.series.sourceSeries.showConnectorLines !== false) {
        const nextX = plotArea.x + (i + 1) * bandW + (bandW - barW) / 2;
        const connY = valToY(s.end);
        connectors.push({ x1: x + barW, y1: connY, x2: nextX, y2: connY });
      }
    }

    return {
      width: safeW, height: safeH, plotArea, ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines: scale.ticks.map(t => ({ x1: plotArea.x, y1: valToY(t), x2: plotArea.x + plotArea.width, y2: valToY(t) })),
      xAxis: {
        line: { x1: plotArea.x, y1: valToY(0), x2: plotArea.x + plotArea.width, y2: valToY(0) },
        ticks: steps.map((s, i) => ({
          x: plotArea.x + (i + 0.5) * bandW,
          y: plotArea.y + plotArea.height + 14,
          label: s.category.length > 10 ? s.category.slice(0, 9) + '…' : s.category,
          anchor: 'middle',
        })),
      },
      yAxis: {
        line: { x1: plotArea.x, y1: plotArea.y, x2: plotArea.x, y2: plotArea.y + plotArea.height },
        ticks: scale.ticks.map(t => ({
          x: plotArea.x - 6, y: valToY(t) + 4,
          label: formatAxisTick(t, primary.series.valuesFormat),
          anchor: 'end',
        })),
      },
      bars,
      areas: [],
      lines: [],
      slices: [],
      calloutConnectors: [],
      waterfallConnectors: connectors,
      bubbles: [],
      isEmpty: false,
    };
  }

  // Funnel
  if (chart.chartType === 'funnel') {
    const { title, titleOffset, showLegend, legendLayout } = buildHeaderAndLegend(
      data, chart, safeW, safeH,
      activeSeries.slice(0, 1).map(e => ({
        label: e.series.name,
        color: getSeriesColor(chart.paletteId, e.index, e.series.sourceSeries.fillRgb),
        seriesIndex: e.index,
      })),
    );
    const leftPad = 78 + legendLayout.padLeft;
    const rightPad = 24 + legendLayout.padRight;
    const topPad = titleOffset + legendLayout.padTop + 6;
    const bottomPad = 18 + legendLayout.padBottom;
    const plotArea: ChartRect = {
      x: leftPad, y: topPad,
      width: Math.max(48, safeW - leftPad - rightPad),
      height: Math.max(40, safeH - topPad - bottomPad),
    };
    const primary = activeSeries[0];
    if (!primary || data.categories.length === 0) {
      return emptyLayout(safeW, safeH, title, showLegend ? legendLayout.items : undefined, plotArea);
    }
    const vals = primary.series.points.map(p => Math.max(0, p.value ?? 0));
    const maxVal = Math.max(1e-9, ...vals);
    const bandH = plotArea.height / vals.length;
    const fill = getSeriesColor(chart.paletteId, primary.index, primary.series.sourceSeries.fillRgb);
    const bars: ChartBarRect[] = vals.map((v, i) => {
      const w = Math.max(4, (v / maxVal) * plotArea.width);
      const h = Math.max(4, bandH * 0.68);
      return {
        x: plotArea.x + (plotArea.width - w) / 2,
        y: plotArea.y + i * bandH + (bandH - h) / 2,
        width: w,
        height: h,
        fill,
        seriesIndex: primary.index,
        categoryIndex: i,
        seriesName: primary.series.name,
        category: data.categories[i] ?? String(i + 1),
        value: v,
        formatted: primary.series.points[i]?.formatted || String(v),
      };
    });
    return {
      width: safeW, height: safeH, plotArea, ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines: [],
      yAxis: {
        line: { x1: plotArea.x, y1: plotArea.y, x2: plotArea.x, y2: plotArea.y },
        ticks: data.categories.map((cat, i) => ({
          x: plotArea.x - 8,
          y: plotArea.y + (i + 0.5) * bandH + 4,
          label: cat.length > 12 ? cat.slice(0, 11) + '…' : cat,
          anchor: 'end',
        })),
      },
      bars,
      areas: [],
      lines: [],
      slices: [],
      calloutConnectors: [],
      bubbles: [],
      isEmpty: false,
    };
  }

  // Histogram & Pareto
  if (chart.chartType === 'histogram' || chart.chartType === 'pareto') {
    const primary = activeSeries[0];
    const isPareto = chart.chartType === 'pareto';
    const legendInputs = primary
      ? [
        { label: primary.series.name, color: getSeriesColor(chart.paletteId, 0, primary.series.sourceSeries.fillRgb), seriesIndex: 0 },
        ...(isPareto ? [{ label: 'Cumulative %', color: '#f38c4e', seriesIndex: 1 }] : []),
      ]
      : [];
    const { title, titleOffset, showLegend, legendLayout } = buildHeaderAndLegend(data, chart, safeW, safeH, legendInputs);
    const histTitlePad = axisTitlePadding(chart, false);
    const leftPad = 48 + legendLayout.padLeft + histTitlePad.left;
    const rightPad = (isPareto ? 48 : 18) + legendLayout.padRight;
    const topPad = titleOffset + legendLayout.padTop + 6;
    const bottomPad = 32 + legendLayout.padBottom + histTitlePad.bottom;
    const plotArea: ChartRect = {
      x: leftPad, y: topPad,
      width: Math.max(48, safeW - leftPad - rightPad),
      height: Math.max(40, safeH - topPad - bottomPad),
    };
    const rawVals = (primary?.series.points ?? [])
      .map(p => p.value)
      .filter((v): v is number => v !== null && Number.isFinite(v));
    if (!primary || rawVals.length === 0) {
      return emptyLayout(safeW, safeH, title, showLegend ? legendLayout.items : undefined, plotArea);
    }

    const ss = primary.series.sourceSeries;
    const binOpts = {
      ...(ss.binMode !== undefined && { binMode: ss.binMode }),
      ...(ss.binWidth !== undefined && { binWidth: ss.binWidth }),
      ...(ss.binCount !== undefined && { binCount: ss.binCount }),
      ...(ss.overflowBin !== undefined && { overflowBin: ss.overflowBin }),
      ...(ss.underflowBin !== undefined && { underflowBin: ss.underflowBin }),
    };
    const paretoBins = isPareto ? computeParetoBins(rawVals, binOpts) : [];
    const bins = isPareto ? paretoBins : computeHistogramBins(rawVals, binOpts);
    const maxCount = Math.max(1, ...bins.map(b => b.count));
    const scale = computeNiceAxisScale(0, maxCount, { includeZero: true });
    const valToY = (v: number) =>
      plotArea.y + plotArea.height - ((v - scale.min) / Math.max(1e-9, scale.max - scale.min)) * plotArea.height;
    const pctToY = (pct: number) => plotArea.y + plotArea.height - pct * plotArea.height;
    const bandW = plotArea.width / Math.max(1, bins.length);
    const barColor = getSeriesColor(chart.paletteId, 0, ss.fillRgb);
    const bars: ChartBarRect[] = bins.map((b, i) => {
      const topY = valToY(b.count);
      const baseY = valToY(0);
      return {
        x: plotArea.x + i * bandW + 1,
        y: topY,
        width: Math.max(4, bandW - 2),
        height: Math.max(1, baseY - topY),
        fill: barColor,
        stroke: '#181b20',
        seriesIndex: 0,
        categoryIndex: i,
        seriesName: primary.series.name,
        category: b.label,
        value: b.count,
        formatted: String(b.count),
      };
    });

    const lines: ChartLineSeriesLayout[] = [];
    if (isPareto) {
      const pctPts = paretoBins.map((b, i) => ({
        x: plotArea.x + (i + 0.5) * bandW,
        y: pctToY(b.cumulativePercent),
        pct: b.cumulativePercent,
        label: b.label,
        i,
      }));
      lines.push({
        path: straightPath(pctPts),
        color: '#f38c4e',
        width: 2,
        seriesIndex: 1,
        markers: pctPts.map(p => ({
          x: p.x, y: p.y, radius: 3.5, fill: '#f38c4e', stroke: '#181b20',
          seriesIndex: 1, categoryIndex: p.i, seriesName: 'Cumulative %',
          category: p.label, value: p.pct * 100, formatted: `${Math.round(p.pct * 100)}%`,
        })),
      });
    }

    return {
      width: safeW, height: safeH, plotArea, ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines: scale.ticks.map(t => ({ x1: plotArea.x, y1: valToY(t), x2: plotArea.x + plotArea.width, y2: valToY(t) })),
      xAxis: {
        line: { x1: plotArea.x, y1: plotArea.y + plotArea.height, x2: plotArea.x + plotArea.width, y2: plotArea.y + plotArea.height },
        ticks: bins.map((b, i) => ({
          x: plotArea.x + (i + 0.5) * bandW,
          y: plotArea.y + plotArea.height + 14,
          label: b.label.length > 14 ? b.label.slice(0, 13) + '…' : b.label,
          anchor: 'middle',
        })),
      },
      yAxis: {
        line: { x1: plotArea.x, y1: plotArea.y, x2: plotArea.x, y2: plotArea.y + plotArea.height },
        ticks: scale.ticks.map(t => ({ x: plotArea.x - 6, y: valToY(t) + 4, label: formatAxisTick(t), anchor: 'end' })),
      },
      ...(isPareto && {
        secondaryYAxis: {
          line: { x1: plotArea.x + plotArea.width, y1: plotArea.y, x2: plotArea.x + plotArea.width, y2: plotArea.y + plotArea.height },
          ticks: [0, 0.2, 0.4, 0.6, 0.8, 1].map(p => ({
            x: plotArea.x + plotArea.width + 6, y: pctToY(p) + 4, label: `${Math.round(p * 100)}%`, anchor: 'start',
          })),
        },
      }),
      bars,
      areas: [],
      lines,
      slices: [],
      calloutConnectors: [],
      bubbles: [],
      isEmpty: false,
    };
  }

  // Box & Whisker
  if (chart.chartType === 'boxWhisker') {
    const { title, titleOffset, showLegend, legendLayout } = buildHeaderAndLegend(
      data, chart, safeW, safeH,
      activeSeries.map(e => ({
        label: e.series.name,
        color: getSeriesColor(chart.paletteId, e.index, e.series.sourceSeries.fillRgb),
        seriesIndex: e.index,
      })),
    );
    const leftPad = 52 + legendLayout.padLeft;
    const rightPad = 18 + legendLayout.padRight;
    const topPad = titleOffset + legendLayout.padTop + 6;
    const bottomPad = 32 + legendLayout.padBottom;
    const plotArea: ChartRect = {
      x: leftPad, y: topPad,
      width: Math.max(48, safeW - leftPad - rightPad),
      height: Math.max(40, safeH - topPad - bottomPad),
    };
    const statsList = activeSeries
      .map(entry => {
        const vals = entry.series.points.map(p => p.value).filter((v): v is number => v !== null);
        const stats = computeBoxWhiskerStats(vals, entry.series.sourceSeries.quartileMethod ?? 'exclusive');
        return stats ? { entry, stats } : undefined;
      })
      .filter((x): x is NonNullable<typeof x> => x !== undefined);

    if (statsList.length === 0) {
      return emptyLayout(safeW, safeH, title, showLegend ? legendLayout.items : undefined, plotArea);
    }

    const lo = Math.min(...statsList.map(s => s.stats.min));
    const hi = Math.max(...statsList.map(s => s.stats.max));
    const scale = computeNiceAxisScale(lo, hi, { includeZero: false });
    const valToY = (v: number) =>
      plotArea.y + plotArea.height - ((v - scale.min) / Math.max(1e-9, scale.max - scale.min)) * plotArea.height;

    const bandW = plotArea.width / statsList.length;
    const boxW = Math.max(12, Math.min(48, bandW * 0.45));
    const boxWhiskers: ChartBoxWhiskerGeom[] = statsList.map(({ entry, stats }, i) => {
      const cx = plotArea.x + (i + 0.5) * bandW;
      const color = getSeriesColor(chart.paletteId, entry.index, entry.series.sourceSeries.fillRgb);
      return {
        x: cx - boxW / 2,
        width: boxW,
        yBoxTop: valToY(stats.q3),
        yBoxBottom: valToY(stats.q1),
        yMedian: valToY(stats.median),
        yWhiskerTop: valToY(stats.whiskerHigh),
        yWhiskerBottom: valToY(stats.whiskerLow),
        yMean: valToY(stats.mean),
        fill: color,
        stroke: '#e6ebf2',
        showMeanMarker: entry.series.sourceSeries.showMeanMarkers !== false,
        outliers: entry.series.sourceSeries.showOutlierPoints !== false
          ? stats.outliers.map(v => ({ x: cx, y: valToY(v), value: v }))
          : [],
        innerPoints: entry.series.sourceSeries.showInnerPoints
          ? stats.innerPoints.map(v => ({ x: cx, y: valToY(v), value: v }))
          : [],
        seriesName: entry.series.name,
        category: entry.series.name,
        summary: `Q1: ${formatNumber(stats.q1).text} · Med: ${formatNumber(stats.median).text} · Q3: ${formatNumber(stats.q3).text}`,
      };
    });

    return {
      width: safeW, height: safeH, plotArea, ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines: scale.ticks.map(t => ({ x1: plotArea.x, y1: valToY(t), x2: plotArea.x + plotArea.width, y2: valToY(t) })),
      xAxis: {
        line: { x1: plotArea.x, y1: plotArea.y + plotArea.height, x2: plotArea.x + plotArea.width, y2: plotArea.y + plotArea.height },
        ticks: statsList.map((s, i) => ({
          x: plotArea.x + (i + 0.5) * bandW,
          y: plotArea.y + plotArea.height + 14,
          label: s.entry.series.name,
          anchor: 'middle',
        })),
      },
      yAxis: {
        line: { x1: plotArea.x, y1: plotArea.y, x2: plotArea.x, y2: plotArea.y + plotArea.height },
        ticks: scale.ticks.map(t => ({ x: plotArea.x - 6, y: valToY(t) + 4, label: formatAxisTick(t), anchor: 'end' })),
      },
      bars: [],
      areas: [],
      lines: [],
      slices: [],
      calloutConnectors: [],
      bubbles: [],
      boxWhiskers,
      isEmpty: false,
    };
  }

  // Hierarchy: Treemap & Sunburst
  if (chart.chartType === 'treemap' || chart.chartType === 'sunburst') {
    const primary = activeSeries[0];
    const items = (primary?.series.points ?? []).map((p, idx) => ({
      label: data.categories[idx] ?? String(idx + 1),
      value: Math.max(0, p.value ?? 0),
      formatted: p.formatted || String(p.value ?? 0),
      index: idx,
    }));
    const { title, titleOffset, showLegend, legendLayout } = buildHeaderAndLegend(
      data, chart, safeW, safeH,
      items.map(it => ({ label: it.label, color: getSeriesColor(chart.paletteId, it.index), seriesIndex: it.index })),
    );
    const plotArea: ChartRect = {
      x: 16 + legendLayout.padLeft,
      y: titleOffset + legendLayout.padTop + 4,
      width: Math.max(48, safeW - 32 - legendLayout.padLeft - legendLayout.padRight),
      height: Math.max(48, safeH - titleOffset - legendLayout.padTop - legendLayout.padBottom - 12),
    };
    const total = items.reduce((s, it) => s + it.value, 0);
    if (!primary || total <= 0) {
      return emptyLayout(safeW, safeH, title, showLegend ? legendLayout.items : undefined, plotArea);
    }

    if (chart.chartType === 'treemap') {
      return {
        width: safeW, height: safeH, plotArea, ...(title && { title }),
        ...(showLegend && { legend: legendLayout.items }),
        gridlines: [],
        bars: [], areas: [], lines: [], slices: [], calloutConnectors: [], bubbles: [],
        treemapRects: squarifyTreemap(items, plotArea, chart.paletteId),
        isEmpty: false,
      };
    }

    // Sunburst concentric rings (outer series rings + inner category ring)
    const cx = plotArea.x + plotArea.width / 2;
    const cy = plotArea.y + plotArea.height / 2;
    const maxR = Math.max(24, Math.min(plotArea.width, plotArea.height) * 0.46);
    const slices: ChartPieSlice[] = [];
    const rings = Math.max(1, Math.min(3, activeSeries.length));
    for (let rIdx = 0; rIdx < rings; rIdx++) {
      const sEntry = activeSeries[rIdx]!;
      const rInner = maxR * (0.24 + (rIdx / rings) * 0.72);
      const rOuter = maxR * (0.24 + ((rIdx + 1) / rings) * 0.72);
      const ringTotal = sEntry.series.points.reduce((sum, p) => sum + Math.max(0, p.value ?? 0), 0);
      if (ringTotal <= 0) continue;
      let angle = -Math.PI / 2;
      for (let i = 0; i < sEntry.series.points.length; i++) {
        const pt = sEntry.series.points[i]!;
        const v = Math.max(0, pt.value ?? 0);
        if (v <= 0) continue;
        const sweep = (v / ringTotal) * Math.PI * 2;
        const end = angle + sweep;
        const mid = (angle + end) / 2;
        const midR = (rInner + rOuter) / 2;
        slices.push({
          path: arcSectorPath(cx, cy, rOuter, rInner, angle, end),
          fill: getSeriesColor(chart.paletteId, i),
          stroke: '#181b20',
          seriesIndex: sEntry.index,
          categoryIndex: i,
          seriesName: sEntry.series.name,
          category: data.categories[i] ?? String(i + 1),
          value: v,
          formatted: pt.formatted || String(v),
          percent: v / ringTotal,
          centerX: cx,
          centerY: cy,
          labelX: cx + midR * Math.cos(mid),
          labelY: cy + midR * Math.sin(mid),
        });
        angle = end;
      }
    }
    return {
      width: safeW, height: safeH, plotArea, ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines: [],
      bars: [], areas: [], lines: [], slices, calloutConnectors: [], bubbles: [],
      isEmpty: slices.length === 0,
    };
  }

  // Stock (HLC, OHLC, VHLC, VOHLC)
  if (cat === 'stock') {
    const hasVol = chart.chartType === 'stockVhlc' || chart.chartType === 'stockVohlc';
    const hasOpen = chart.chartType === 'stockOhlc' || chart.chartType === 'stockVohlc';
    const minSeries = (hasVol ? 1 : 0) + (hasOpen ? 4 : 3);
    const { title, titleOffset, showLegend, legendLayout } = buildHeaderAndLegend(
      data, chart, safeW, safeH,
      activeSeries.map(e => ({ label: e.series.name, color: getSeriesColor(chart.paletteId, e.index), seriesIndex: e.index })),
    );
    const leftPad = 52 + legendLayout.padLeft;
    const rightPad = (hasVol ? 52 : 18) + legendLayout.padRight;
    const topPad = titleOffset + legendLayout.padTop + 6;
    const bottomPad = 32 + legendLayout.padBottom;
    const plotArea: ChartRect = {
      x: leftPad, y: topPad,
      width: Math.max(48, safeW - leftPad - rightPad),
      height: Math.max(40, safeH - topPad - bottomPad),
    };
    if (activeSeries.length < minSeries || data.categories.length === 0) {
      return emptyLayout(safeW, safeH, title, showLegend ? legendLayout.items : undefined, plotArea);
    }

    const offset = hasVol ? 1 : 0;
    const volSeries = hasVol ? activeSeries[0]! : undefined;
    const openSeries = hasOpen ? activeSeries[offset]! : undefined;
    const highSeries = activeSeries[offset + (hasOpen ? 1 : 0)]!;
    const lowSeries = activeSeries[offset + (hasOpen ? 2 : 1)]!;
    const closeSeries = activeSeries[offset + (hasOpen ? 3 : 2)]!;

    const priceVals: number[] = [];
    const volVals: number[] = [];
    for (let i = 0; i < data.categories.length; i++) {
      for (const s of [openSeries, highSeries, lowSeries, closeSeries]) {
        const v = s?.series.points[i]?.value;
        if (v !== null && v !== undefined) priceVals.push(v);
      }
      const vv = volSeries?.series.points[i]?.value;
      if (vv !== null && vv !== undefined) volVals.push(vv);
    }
    if (priceVals.length === 0) {
      return emptyLayout(safeW, safeH, title, showLegend ? legendLayout.items : undefined, plotArea);
    }

    const pScale = computeNiceAxisScale(Math.min(...priceVals), Math.max(...priceVals), { includeZero: false });
    const pToY = (v: number) =>
      plotArea.y + plotArea.height - ((v - pScale.min) / Math.max(1e-9, pScale.max - pScale.min)) * plotArea.height;
    const vScale = hasVol && volVals.length > 0
      ? computeNiceAxisScale(0, Math.max(...volVals) * 2.5, { includeZero: true })
      : undefined;
    const vToY = vScale
      ? (v: number) => plotArea.y + plotArea.height - ((v - vScale.min) / Math.max(1e-9, vScale.max - vScale.min)) * plotArea.height
      : undefined;

    const bandW = plotArea.width / data.categories.length;
    const bars: ChartBarRect[] = [];
    const stockStems: ChartStockGeom[] = [];

    for (let i = 0; i < data.categories.length; i++) {
      const cx = plotArea.x + (i + 0.5) * bandW;
      if (volSeries && vToY) {
        const vv = volSeries.series.points[i]?.value;
        if (vv !== null && vv !== undefined) {
          const topY = vToY(vv);
          const baseY = vToY(0);
          const w = Math.max(4, bandW * 0.45);
          bars.push({
            x: cx - w / 2, y: topY, width: w, height: Math.max(1, baseY - topY),
            fill: '#676d77', seriesIndex: 0, categoryIndex: i,
            seriesName: volSeries.series.name, category: data.categories[i] ?? '',
            value: vv, formatted: volSeries.series.points[i]?.formatted || String(vv),
          });
        }
      }
      const h = highSeries.series.points[i]?.value;
      const l = lowSeries.series.points[i]?.value;
      const c = closeSeries.series.points[i]?.value;
      const o = openSeries?.series.points[i]?.value;
      if (h === null || h === undefined || l === null || l === undefined || c === null || c === undefined) continue;
      const isDown = o !== null && o !== undefined ? c < o : false;
      stockStems.push({
        x: cx,
        yHigh: pToY(h),
        yLow: pToY(l),
        ...(o !== null && o !== undefined && { yOpen: pToY(o) }),
        yClose: pToY(c),
        boxWidth: Math.max(6, Math.min(18, bandW * 0.4)),
        isDown,
        category: data.categories[i] ?? String(i + 1),
        summary: `${o !== undefined && o !== null ? `O: ${o} · ` : ''}H: ${h} · L: ${l} · C: ${c}`,
      });
    }

    return {
      width: safeW, height: safeH, plotArea, ...(title && { title }),
      ...(showLegend && { legend: legendLayout.items }),
      gridlines: pScale.ticks.map(t => ({ x1: plotArea.x, y1: pToY(t), x2: plotArea.x + plotArea.width, y2: pToY(t) })),
      xAxis: {
        line: { x1: plotArea.x, y1: plotArea.y + plotArea.height, x2: plotArea.x + plotArea.width, y2: plotArea.y + plotArea.height },
        ticks: data.categories.map((c, i) => ({
          x: plotArea.x + (i + 0.5) * bandW, y: plotArea.y + plotArea.height + 14, label: c, anchor: 'middle',
        })),
      },
      yAxis: {
        line: { x1: plotArea.x, y1: plotArea.y, x2: plotArea.x, y2: plotArea.y + plotArea.height },
        ticks: pScale.ticks.map(t => ({ x: plotArea.x - 6, y: pToY(t) + 4, label: formatAxisTick(t), anchor: 'end' })),
      },
      ...(vScale && vToY && {
        secondaryYAxis: {
          line: { x1: plotArea.x + plotArea.width, y1: plotArea.y, x2: plotArea.x + plotArea.width, y2: plotArea.y + plotArea.height },
          ticks: vScale.ticks.map(t => ({ x: plotArea.x + plotArea.width + 6, y: vToY(t) + 4, label: formatAxisTick(t), anchor: 'start' })),
        },
      }),
      bars,
      areas: [],
      lines: [],
      slices: [],
      calloutConnectors: [],
      bubbles: [],
      stockStems,
      isEmpty: stockStems.length === 0,
    };
  }

  // Radar (radar, radarMarkers, filledRadar)
  const { title, titleOffset, showLegend, legendLayout } = buildHeaderAndLegend(
    data, chart, safeW, safeH,
    activeSeries.map(e => ({
      label: e.series.name,
      color: getSeriesColor(chart.paletteId, e.index, e.series.sourceSeries.fillRgb),
      seriesIndex: e.index,
    })),
  );
  const plotArea: ChartRect = {
    x: 20 + legendLayout.padLeft,
    y: titleOffset + legendLayout.padTop + 6,
    width: Math.max(60, safeW - 40 - legendLayout.padLeft - legendLayout.padRight),
    height: Math.max(60, safeH - titleOffset - legendLayout.padTop - legendLayout.padBottom - 16),
  };
  const n = data.categories.length;
  if (activeSeries.length === 0 || n < 3) {
    return emptyLayout(safeW, safeH, title, showLegend ? legendLayout.items : undefined, plotArea);
  }
  const allVals = activeSeries.flatMap(s => s.series.points.map(p => p.value).filter((v): v is number => v !== null));
  const scale = computeNiceAxisScale(0, Math.max(1, ...allVals), { includeZero: true });
  const cx = plotArea.x + plotArea.width / 2;
  const cy = plotArea.y + plotArea.height / 2;
  const maxR = Math.max(20, Math.min(plotArea.width, plotArea.height) * 0.39);
  const ringCount = Math.max(3, Math.min(6, scale.ticks.length - 1));

  const angleAt = (idx: number) => -Math.PI / 2 + (idx / n) * Math.PI * 2;
  const rings: string[] = [];
  for (let r = 1; r <= ringCount; r++) {
    const rad = (r / ringCount) * maxR;
    const pts = Array.from({ length: n }, (_, i) => ({
      x: cx + rad * Math.cos(angleAt(i)),
      y: cy + rad * Math.sin(angleAt(i)),
    }));
    rings.push(`${straightPath(pts)} Z`);
  }

  const spokes = data.categories.map((cat, i) => {
    const a = angleAt(i);
    const cos = Math.cos(a);
    const sin = Math.sin(a);
    return {
      x1: cx,
      y1: cy,
      x2: cx + maxR * cos,
      y2: cy + maxR * sin,
      labelX: cx + (maxR + 12) * cos,
      labelY: cy + (maxR + 12) * sin + 4,
      anchor: (Math.abs(cos) < 0.2 ? 'middle' : cos > 0 ? 'start' : 'end') as 'start' | 'middle' | 'end',
      label: cat.length > 10 ? cat.slice(0, 9) + '…' : cat,
    };
  });

  const isFilled = chart.chartType === 'filledRadar';
  const showMarkers = chart.chartType === 'radarMarkers';
  const polygons = activeSeries.map(entry => {
    const color = getSeriesColor(chart.paletteId, entry.index, entry.series.sourceSeries.fillRgb);
    const pts = data.categories.map((_, i) => {
      const v = entry.series.points[i]?.value ?? scale.min;
      const frac = Math.max(0, Math.min(1, (v - scale.min) / Math.max(1e-9, scale.max - scale.min)));
      const rad = frac * maxR;
      const a = angleAt(i);
      return {
        x: cx + rad * Math.cos(a),
        y: cy + rad * Math.sin(a),
        value: v,
        formatted: entry.series.points[i]?.formatted || String(v),
        i,
      };
    });
    const markers: ChartMarkerPoint[] = showMarkers
      ? pts.map(p => ({
        x: p.x, y: p.y, radius: 3.5, fill: color, stroke: '#181b20',
        seriesIndex: entry.index, categoryIndex: p.i,
        seriesName: entry.series.name, category: data.categories[p.i] ?? '',
        value: p.value, formatted: p.formatted,
      }))
      : [];
    return {
      path: `${straightPath(pts)} Z`,
      stroke: color,
      fill: color,
      fillOpacity: isFilled ? 0.42 : 0,
      markers,
    };
  });

  const radarWeb: ChartRadarWeb = { rings, spokes, polygons };
  return {
    width: safeW, height: safeH, plotArea, ...(title && { title }),
    ...(showLegend && { legend: legendLayout.items }),
    gridlines: [],
    bars: [], areas: [], lines: [], slices: [], calloutConnectors: [], bubbles: [],
    radarWeb,
    isEmpty: false,
  };
}
