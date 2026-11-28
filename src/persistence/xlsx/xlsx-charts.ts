import { formatAbsoluteAreaRef, resolveFormulaExtent } from '../../charts/chart-data';
import {
  defaultChartElements,
  nextChartName,
  type ChartAxesConfig,
  type ChartAxisConfig,
  type ChartDataLabelPosition,
  type ChartDataLabelsConfig,
  type ChartDocument,
  type ChartErrorBarDirection,
  type ChartErrorBars,
  type ChartErrorBarType,
  type ChartGridlinesConfig,
  type ChartHiddenEmptyMode,
  type ChartLegendPosition,
  type ChartPlacement,
  type ChartPlacementMode,
  type ChartSeries,
  type ChartTitleConfig,
  type ChartTrendline,
  type ChartTrendlineType,
  type ChartType,
} from '../../core/chart';
import { documentLimits, type DocumentId, type IdFactory, type WorkbookDocument } from '../../core/workbook';
import {
  attr,
  childElements,
  descendantElements,
  firstChild,
  parseXmlPart,
  relationshipNamespace,
  textOf,
  type Element,
} from './xlsx-xml';

export interface ChartOoxmlRel {
  readonly id: string;
  readonly type: string;
  readonly target: string;
}

export interface ChartSheetMeta {
  readonly id: DocumentId;
  readonly name: string;
  readonly targetPath: string;
  readonly root: Element;
  readonly rels: readonly ChartOoxmlRel[];
}

export interface ImportChartsResult {
  readonly document: WorkbookDocument;
  readonly importedChartParts: ReadonlySet<string>;
  readonly importedDrawingParts: ReadonlySet<string>;
  readonly hasUnsupportedCharts: boolean;
  readonly hasUnimportedDrawings: boolean;
}

const defaultColWidthPx = 96;
const defaultRowHeightPx = 26;
const emuPerPx = 9525;

const unsupportedPlotTags = new Set([
  'bar3DChart',
  'line3DChart',
  'area3DChart',
  'pie3DChart',
  'surfaceChart',
  'surface3DChart',
  'doughnut3DChart',
]);

const supportedPlotTags = new Set([
  'barChart',
  'lineChart',
  'areaChart',
  'pieChart',
  'doughnutChart',
  'ofPieChart',
  'scatterChart',
  'bubbleChart',
  'radarChart',
  'stockChart',
]);

const nonChartDrawingTags = new Set(['sp', 'pic', 'cxnSp', 'grpSp']);

function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash >= 0 ? partPath.slice(0, slash) : '';
  const file = slash >= 0 ? partPath.slice(slash + 1) : partPath;
  return `${dir ? `${dir}/` : ''}_rels/${file}.rels`;
}

function resolveRelativeTarget(basePartPath: string, target: string): string | undefined {
  const text = target.replaceAll('\\', '/').trim();
  if (!text || text.includes(':')) return undefined;
  const segments = text.startsWith('/') ? [] : basePartPath.split('/').slice(0, -1);
  for (const part of text.replace(/^\/+/u, '').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') {
      if (segments.length === 0) return undefined;
      segments.pop();
      continue;
    }
    segments.push(part);
  }
  return segments.join('/');
}

function readRels(parts: ReadonlyMap<string, Uint8Array>, partPath: string): readonly ChartOoxmlRel[] {
  const bytes = parts.get(relsPathFor(partPath).toLowerCase());
  if (!bytes) return [];
  const doc = parseXmlPart(bytes, `Relationships for ${partPath}`);
  const rels: ChartOoxmlRel[] = [];
  for (const rel of childElements(doc.root, 'Relationship')) {
    if (attr(rel, 'TargetMode') === 'External') continue;
    const id = attr(rel, 'Id');
    const type = attr(rel, 'Type') ?? '';
    const rawTarget = attr(rel, 'Target');
    if (!id || !rawTarget) continue;
    const target = resolveRelativeTarget(partPath, rawTarget);
    if (target) rels.push({ id, type, target });
  }
  return rels;
}

function relId(el: Element): string | undefined {
  const qualified = el.getAttributeNS(relationshipNamespace, 'id');
  return qualified === null || qualified === '' ? attr(el, 'r:id') : qualified;
}

function valAttr(parent: Element | undefined, childName: string): string | undefined {
  const el = firstChild(parent, childName);
  return el ? attr(el, 'val') : undefined;
}

function numVal(parent: Element | undefined, childName: string): number | undefined {
  const raw = valAttr(parent, childName);
  if (raw === undefined) return undefined;
  const n = Number(raw);
  return Number.isFinite(n) ? n : undefined;
}

function normalizeOoxmlFormula(raw: string | undefined): string | undefined {
  const trimmed = raw?.trim();
  if (!trimmed) return undefined;
  return trimmed.startsWith('=') ? trimmed : `=${trimmed}`;
}

function clampInt(n: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, Math.round(n)));
}

function clampOffset(n: number): number {
  return Math.max(-4096, Math.min(4096, Math.round(n)));
}

function readAnchorMarker(marker: Element | undefined): {
  readonly row: number;
  readonly column: number;
  readonly offsetX: number;
  readonly offsetY: number;
} {
  const colRaw = Number(textOf(firstChild(marker, 'col')) ?? '0');
  const rowRaw = Number(textOf(firstChild(marker, 'row')) ?? '0');
  const colOffRaw = Number(textOf(firstChild(marker, 'colOff')) ?? '0');
  const rowOffRaw = Number(textOf(firstChild(marker, 'rowOff')) ?? '0');
  return {
    row: Number.isFinite(rowRaw) ? clampInt(rowRaw + 1, 1, documentLimits.rows) : 1,
    column: Number.isFinite(colRaw) ? clampInt(colRaw + 1, 1, documentLimits.columns) : 1,
    offsetX: Number.isFinite(colOffRaw) ? clampOffset(colOffRaw / emuPerPx) : 0,
    offsetY: Number.isFinite(rowOffRaw) ? clampOffset(rowOffRaw / emuPerPx) : 0,
  };
}

function readAnchorPlacement(anchorEl: Element, sheetId: DocumentId): ChartPlacement {
  const editAs = attr(anchorEl, 'editAs');
  const mode: ChartPlacementMode = editAs === 'absolute'
    ? 'freeFloating'
    : editAs === 'oneCell'
      ? 'moveOnly'
      : 'moveAndSize';

  const fromEl = firstChild(anchorEl, 'from');
  const toEl = firstChild(anchorEl, 'to');
  const extEl = firstChild(anchorEl, 'ext');

  const from = fromEl ? readAnchorMarker(fromEl) : { row: 1, column: 1, offsetX: 0, offsetY: 0 };
  let width = 480;
  let height = 300;
  if (toEl) {
    const to = readAnchorMarker(toEl);
    const w = (to.column - from.column) * defaultColWidthPx + (to.offsetX - from.offsetX);
    const h = (to.row - from.row) * defaultRowHeightPx + (to.offsetY - from.offsetY);
    if (w >= 40) width = w;
    if (h >= 40) height = h;
  } else if (extEl) {
    const cx = Number(attr(extEl, 'cx') ?? '');
    const cy = Number(attr(extEl, 'cy') ?? '');
    if (Number.isFinite(cx) && cx > 0) width = cx / emuPerPx;
    if (Number.isFinite(cy) && cy > 0) height = cy / emuPerPx;
  }

  return {
    sheetId,
    row: from.row,
    column: from.column,
    offsetX: from.offsetX,
    offsetY: from.offsetY,
    width: Math.max(120, Math.min(4096, Math.round(width))),
    height: Math.max(90, Math.min(4096, Math.round(height))),
    mode,
  };
}

function readRichText(parent: Element | undefined): string | undefined {
  if (!parent) return undefined;
  const parts = descendantElements(parent, 't')
    .map(el => textOf(el) ?? '')
    .filter(t => t.length > 0);
  return parts.length > 0 ? parts.join('') : undefined;
}

function readTitleConfig(titleEl: Element | undefined, autoDeleted: boolean): ChartTitleConfig | undefined {
  if (autoDeleted) return { show: false };
  if (!titleEl) return undefined;
  const tx = firstChild(titleEl, 'tx');
  const formulaRef = normalizeOoxmlFormula(textOf(firstChild(firstChild(tx, 'strRef'), 'f')));
  const richText = readRichText(firstChild(tx, 'rich'));
  const overlay = valAttr(titleEl, 'overlay') === '1';
  if (formulaRef) {
    return { show: true, formulaRef, ...(overlay && { overlay: true }) };
  }
  return {
    show: true,
    ...(richText && { text: richText }),
    ...(overlay && { overlay: true }),
  };
}

function readFillHex(spPr: Element | undefined): string | undefined {
  const srgb = firstChild(firstChild(spPr, 'solidFill'), 'srgbClr');
  const hex = attr(srgb, 'val')?.trim();
  return hex && /^[0-9A-Fa-f]{6}$/u.test(hex) ? `#${hex.toUpperCase()}` : undefined;
}

function mapLabelPos(val: string | undefined): ChartDataLabelPosition | undefined {
  switch (val) {
    case 'ctr': return 'center';
    case 'inEnd': return 'insideEnd';
    case 'inBase': return 'insideBase';
    case 'outEnd': return 'outsideEnd';
    case 'bestFit': return 'bestFit';
    case 'l': return 'left';
    case 'r': return 'right';
    case 't': return 'above';
    case 'b': return 'below';
    default: return undefined;
  }
}

function readDataLabels(el: Element | undefined): ChartDataLabelsConfig | undefined {
  if (!el) return undefined;
  const showVal = valAttr(el, 'showVal') === '1';
  const showCategoryName = valAttr(el, 'showCatName') === '1';
  const showSeriesName = valAttr(el, 'showSerName') === '1';
  const showPercentage = valAttr(el, 'showPercent') === '1';
  const position = mapLabelPos(valAttr(el, 'dLblPos'));
  const show = showVal || showCategoryName || showSeriesName || showPercentage;
  if (!show) return undefined;
  return {
    show: true,
    ...(position && { position }),
    ...(showVal && { showValue: true }),
    ...(showCategoryName && { showCategoryName: true }),
    ...(showSeriesName && { showSeriesName: true }),
    ...(showPercentage && { showPercentage: true }),
  };
}

function readTrendlines(serEl: Element, makeId: IdFactory): readonly ChartTrendline[] | undefined {
  const els = childElements(serEl, 'trendline');
  if (els.length === 0) return undefined;
  const result: ChartTrendline[] = [];
  for (const tl of els) {
    const rawType = valAttr(tl, 'trendlineType');
    const type: ChartTrendlineType = rawType === 'exp'
      ? 'exponential'
      : rawType === 'log'
        ? 'logarithmic'
        : rawType === 'poly'
          ? 'polynomial'
          : rawType === 'power'
            ? 'power'
            : rawType === 'movingAvg'
              ? 'movingAverage'
              : 'linear';
    const order = numVal(tl, 'order');
    const period = numVal(tl, 'period');
    const forward = numVal(tl, 'forward');
    const backward = numVal(tl, 'backward');
    const intercept = numVal(tl, 'intercept');
    const name = textOf(firstChild(tl, 'name'))?.trim();
    const displayEquation = valAttr(tl, 'dispEq') === '1';
    const displayRSquared = valAttr(tl, 'dispRSqr') === '1';
    result.push({
      id: makeId(),
      type,
      ...(order !== undefined && order >= 2 && order <= 6 && { polynomialOrder: Math.round(order) }),
      ...(period !== undefined && period >= 2 && { period: Math.round(period) }),
      ...(forward !== undefined && { forward }),
      ...(backward !== undefined && { backward }),
      ...(intercept !== undefined && { intercept }),
      ...(displayEquation && { displayEquation: true }),
      ...(displayRSquared && { displayRSquared: true }),
      ...(name && { name }),
    });
  }
  return result.length > 0 ? result : undefined;
}

function readErrorBars(serEl: Element): ChartErrorBars | undefined {
  const eb = firstChild(serEl, 'errBars');
  if (!eb) return undefined;
  const dirRaw = valAttr(eb, 'errBarType');
  const direction: ChartErrorBarDirection = dirRaw === 'plus' || dirRaw === 'minus' ? dirRaw : 'both';
  const typeRaw = valAttr(eb, 'errValType');
  const type: ChartErrorBarType = typeRaw === 'fixedVal'
    ? 'fixed'
    : typeRaw === 'percentage'
      ? 'percentage'
      : typeRaw === 'stdDev'
        ? 'stdDev'
        : typeRaw === 'stdErr'
          ? 'stdErr'
          : typeRaw === 'cust'
            ? 'custom'
            : 'percentage';
  const val = numVal(eb, 'val');
  const plusRef = normalizeOoxmlFormula(textOf(firstChild(firstChild(firstChild(eb, 'plus'), 'numRef'), 'f')));
  const minusRef = normalizeOoxmlFormula(textOf(firstChild(firstChild(firstChild(eb, 'minus'), 'numRef'), 'f')));
  const cap = valAttr(eb, 'noEndCap') !== '1';
  return {
    direction,
    type,
    ...(val !== undefined && { value: val }),
    ...(plusRef && { plusRef }),
    ...(minusRef && { minusRef }),
    cap,
  };
}

function mapPlotType(plotEl: Element): ChartType {
  const tag = plotEl.localName;
  const grouping = valAttr(plotEl, 'grouping');
  if (tag === 'barChart') {
    const isBar = valAttr(plotEl, 'barDir') === 'bar';
    if (grouping === 'stacked') return isBar ? 'stackedBar' : 'stackedColumn';
    if (grouping === 'percentStacked') return isBar ? 'stackedBar100' : 'stackedColumn100';
    return isBar ? 'clusteredBar' : 'clusteredColumn';
  }
  if (tag === 'lineChart') {
    const hasMarkers = childElements(plotEl, 'ser').some(ser => {
      const symbol = valAttr(firstChild(ser, 'marker'), 'symbol');
      return symbol !== undefined && symbol !== 'none';
    }) || valAttr(plotEl, 'marker') === '1';
    if (grouping === 'stacked') return hasMarkers ? 'stackedLineMarkers' : 'stackedLine';
    if (grouping === 'percentStacked') return hasMarkers ? 'stackedLine100Markers' : 'stackedLine100';
    return hasMarkers ? 'lineMarkers' : 'line';
  }
  if (tag === 'areaChart') {
    if (grouping === 'stacked') return 'stackedArea';
    if (grouping === 'percentStacked') return 'stackedArea100';
    return 'area';
  }
  if (tag === 'pieChart') return 'pie';
  if (tag === 'doughnutChart') return 'doughnut';
  if (tag === 'ofPieChart') return valAttr(plotEl, 'ofPieType') === 'bar' ? 'barOfPie' : 'pieOfPie';
  if (tag === 'scatterChart') {
    const style = valAttr(plotEl, 'scatterStyle');
    if (style === 'smoothMarker') return 'scatterSmoothLinesMarkers';
    if (style === 'smooth') return 'scatterSmoothLines';
    if (style === 'lineMarker') return 'scatterStraightLinesMarkers';
    if (style === 'line') return 'scatterStraightLines';
    return 'scatter';
  }
  if (tag === 'bubbleChart') return 'bubble';
  if (tag === 'radarChart') {
    const style = valAttr(plotEl, 'radarStyle');
    if (style === 'filled') return 'filledRadar';
    if (style === 'marker') return 'radarMarkers';
    return 'radar';
  }
  const serCount = childElements(plotEl, 'ser').length;
  return serCount >= 4 ? 'stockOhlc' : 'stockHlc';
}

function readAxisConfig(axEl: Element | undefined): ChartAxisConfig | undefined {
  if (!axEl) return undefined;
  const show = valAttr(axEl, 'delete') !== '1';
  const scaling = firstChild(axEl, 'scaling');
  const min = numVal(scaling, 'min');
  const max = numVal(scaling, 'max');
  const logBase = numVal(scaling, 'logBase');
  const reverseOrder = valAttr(scaling, 'orientation') === 'maxMin';
  const majorUnit = numVal(axEl, 'majorUnit');
  const minorUnit = numVal(axEl, 'minorUnit');
  const title = readTitleConfig(firstChild(axEl, 'title'), false);
  return {
    show,
    sourceLinked: true,
    ...(min !== undefined && { min }),
    ...(max !== undefined && { max }),
    ...(logBase !== undefined && { logarithmic: true, logBase }),
    ...(reverseOrder && { reverseOrder: true }),
    ...(majorUnit !== undefined && { majorUnit }),
    ...(minorUnit !== undefined && { minorUnit }),
    ...(title && { title }),
  };
}

function parseChartXml(
  document: WorkbookDocument,
  chartBytes: Uint8Array,
  placement: ChartPlacement,
  makeId: IdFactory,
): ChartDocument | undefined {
  let root: Element;
  try {
    root = parseXmlPart(chartBytes, 'Chart').root;
  } catch {
    return undefined;
  }
  if (root.localName !== 'chartSpace') return undefined;
  const chartEl = firstChild(root, 'chart');
  const plotArea = firstChild(chartEl, 'plotArea');
  if (!chartEl || !plotArea) return undefined;

  const plots: Element[] = [];
  for (const child of childElements(plotArea)) {
    if (child.localName && unsupportedPlotTags.has(child.localName)) return undefined;
    if (child.localName && supportedPlotTags.has(child.localName)) plots.push(child);
  }
  if (plots.length === 0) return undefined;

  const isCombo = plots.length > 1;
  const chartType: ChartType = isCombo ? 'combo' : mapPlotType(plots[0]!);

  const valAxEls = childElements(plotArea, 'valAx');
  const primaryValAxId = valAxEls[0] ? valAttr(valAxEls[0], 'axId') : undefined;
  const catAxEl = firstChild(plotArea, 'catAx') ?? firstChild(plotArea, 'dateAx');
  const primaryValAxEl = valAxEls[0];
  const secondaryValAxEl = valAxEls.length > 1 ? valAxEls[1] : undefined;

  const series: ChartSeries[] = [];
  let plotDataLabels: ChartDataLabelsConfig | undefined;

  for (const plot of plots) {
    const plotType = mapPlotType(plot);
    const plotAxIds = childElements(plot, 'axId').map(el => attr(el, 'val')).filter((v): v is string => Boolean(v));
    const usesSecondaryValAx = isCombo &&
      secondaryValAxEl !== undefined &&
      primaryValAxId !== undefined &&
      plotAxIds.some(id => id !== primaryValAxId && id === valAttr(secondaryValAxEl, 'axId'));

    const pLabels = readDataLabels(firstChild(plot, 'dLbls'));
    if (!plotDataLabels && pLabels) plotDataLabels = pLabels;

    for (const ser of childElements(plot, 'ser')) {
      const tx = firstChild(ser, 'tx');
      const nameRef = normalizeOoxmlFormula(textOf(firstChild(firstChild(tx, 'strRef'), 'f')));
      const nameLit = textOf(firstChild(tx, 'v'))?.trim();

      const catEl = firstChild(ser, 'cat');
      const categoriesRef = normalizeOoxmlFormula(
        textOf(firstChild(firstChild(catEl, 'strRef'), 'f')) ??
        textOf(firstChild(firstChild(catEl, 'numRef'), 'f')) ??
        textOf(firstChild(firstChild(catEl, 'multiLvlStrRef'), 'f')),
      );

      const xValEl = firstChild(ser, 'xVal');
      const xValuesRef = normalizeOoxmlFormula(
        textOf(firstChild(firstChild(xValEl, 'numRef'), 'f')) ??
        textOf(firstChild(firstChild(xValEl, 'strRef'), 'f')),
      );

      const valContainer = firstChild(ser, 'val') ?? firstChild(ser, 'yVal');
      const valuesRef = normalizeOoxmlFormula(textOf(firstChild(firstChild(valContainer, 'numRef'), 'f')));
      if (!valuesRef) continue;

      const bubbleSizeEl = firstChild(ser, 'bubbleSize');
      const bubbleSizesRef = normalizeOoxmlFormula(textOf(firstChild(firstChild(bubbleSizeEl, 'numRef'), 'f')));
      const fillRgb = readFillHex(firstChild(ser, 'spPr'));
      const serLabels = readDataLabels(firstChild(ser, 'dLbls'));
      const trendlines = readTrendlines(ser, makeId);
      const errorBars = readErrorBars(ser);

      series.push({
        id: makeId(),
        ...(nameRef ? { nameRef } : nameLit ? { name: nameLit } : {}),
        ...(categoriesRef && { categoriesRef }),
        valuesRef,
        ...(xValuesRef && { xValuesRef }),
        ...(bubbleSizesRef && { bubbleSizesRef }),
        ...(isCombo && { chartType: plotType }),
        ...(usesSecondaryValAx && { secondaryAxis: true }),
        ...(fillRgb && { fillRgb }),
        ...(serLabels && { dataLabels: serLabels }),
        ...(trendlines && { trendlines }),
        ...(errorBars && { errorBars }),
      });
    }
  }

  if (series.length === 0) return undefined;

  // Derive overarching ChartSource formula from series bounds (falling back to first valuesRef)
  let sourceFormula: string = series[0]!.valuesRef;
  let minRow = Number.POSITIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let maxCol = 0;
  let commonSheetId: DocumentId | undefined;
  let bounded = true;
  for (const s of series) {
    for (const ref of [s.categoriesRef, s.valuesRef, s.xValuesRef]) {
      if (!ref) continue;
      const area = resolveFormulaExtent(document, placement.sheetId, ref);
      if (!area) { bounded = false; break; }
      if (commonSheetId === undefined) commonSheetId = area.sheetId;
      else if (commonSheetId !== area.sheetId) { bounded = false; break; }
      minRow = Math.min(minRow, area.row);
      minCol = Math.min(minCol, area.column);
      maxRow = Math.max(maxRow, area.row + area.rows - 1);
      maxCol = Math.max(maxCol, area.column + area.columns - 1);
    }
  }
  if (bounded && commonSheetId !== undefined && maxRow >= minRow && maxCol >= minCol) {
    const srcSheet = document.sheets.find(s => s.id === commonSheetId);
    if (srcSheet) {
      sourceFormula = formatAbsoluteAreaRef(srcSheet.name, minRow, minCol, maxRow - minRow + 1, maxCol - minCol + 1);
    }
  }

  const baseElements = defaultChartElements(chartType);
  const autoTitleDeleted = valAttr(chartEl, 'autoTitleDeleted') === '1';
  const parsedTitle = readTitleConfig(firstChild(chartEl, 'title'), autoTitleDeleted);

  const legendEl = firstChild(chartEl, 'legend');
  const legendPosRaw = valAttr(legendEl, 'legendPos');
  const legendPosition: ChartLegendPosition = legendPosRaw === 't'
    ? 'top'
    : legendPosRaw === 'l'
      ? 'left'
      : legendPosRaw === 'r'
        ? 'right'
        : 'bottom';

  const catAxCfg = readAxisConfig(catAxEl);
  const valAxCfg = readAxisConfig(primaryValAxEl);
  const secValAxCfg = readAxisConfig(secondaryValAxEl);
  const axes: ChartAxesConfig | undefined = catAxCfg || valAxCfg || secValAxCfg
    ? {
      ...(catAxCfg ? { category: catAxCfg } : baseElements.axes?.category ? { category: baseElements.axes.category } : {}),
      ...(valAxCfg ? { value: valAxCfg } : baseElements.axes?.value ? { value: baseElements.axes.value } : {}),
      ...(secValAxCfg && { secondaryValue: secValAxCfg }),
    }
    : baseElements.axes;

  const hasAxisGrid = Boolean(catAxEl || primaryValAxEl);
  const gridlines: ChartGridlinesConfig | undefined = hasAxisGrid
    ? {
      primaryMajorHorizontal: Boolean(firstChild(primaryValAxEl, 'majorGridlines')),
      primaryMinorHorizontal: Boolean(firstChild(primaryValAxEl, 'minorGridlines')),
      primaryMajorVertical: Boolean(firstChild(catAxEl, 'majorGridlines')),
      primaryMinorVertical: Boolean(firstChild(catAxEl, 'minorGridlines')),
    }
    : baseElements.gridlines;

  const dTableEl = firstChild(plotArea, 'dTable');
  const dispBlanks = valAttr(chartEl, 'dispBlanksAs');
  const hiddenEmptyMode: ChartHiddenEmptyMode = dispBlanks === 'zero'
    ? 'zero'
    : dispBlanks === 'span'
      ? 'connect'
      : 'gaps';
  const showHiddenCells = valAttr(chartEl, 'plotVisOnly') === '0';

  return {
    id: makeId(),
    name: nextChartName(document),
    chartType,
    placement,
    source: { kind: 'range', formula: sourceFormula },
    seriesBy: 'columns',
    series,
    elements: {
      ...baseElements,
      ...(parsedTitle && { title: parsedTitle }),
      legend: legendEl ? { show: true, position: legendPosition } : { show: false, position: 'bottom' },
      ...(axes && { axes }),
      ...(gridlines && { gridlines }),
      ...(plotDataLabels && { dataLabels: plotDataLabels }),
      ...(dTableEl && { dataTable: { show: true, showLegendKeys: valAttr(dTableEl, 'showKeys') !== '0' } }),
    },
    hiddenEmptyMode,
    showHiddenCells,
    styleId: 1,
    paletteId: 'colorful-1',
  };
}

export function importCharts(
  document: WorkbookDocument,
  sheetMetas: readonly ChartSheetMeta[],
  parts: ReadonlyMap<string, Uint8Array>,
  makeId: IdFactory,
): ImportChartsResult {
  const importedChartParts = new Set<string>();
  const importedDrawingParts = new Set<string>();
  let hasUnsupportedCharts = false;
  let hasUnimportedDrawings = false;
  let nextDoc = document;

  for (const sheetMeta of sheetMetas) {
    const referencedDrawingRids = new Set(
      childElements(sheetMeta.root, 'drawing')
        .map(el => relId(el))
        .filter((id): id is string => Boolean(id)),
    );
    const seenDrawings = new Set<string>();
    for (const rel of sheetMeta.rels) {
      const isDrawing = rel.type.toLowerCase().endsWith('/drawing') || referencedDrawingRids.has(rel.id);
      if (!isDrawing) continue;
      const drawingKey = rel.target.toLowerCase();
      if (seenDrawings.has(drawingKey)) continue;
      seenDrawings.add(drawingKey);

      const drawingBytes = parts.get(drawingKey);
      if (!drawingBytes) continue;

      let drawingRoot: Element;
      let drawingRels: readonly ChartOoxmlRel[];
      try {
        drawingRoot = parseXmlPart(drawingBytes, `Drawing ${rel.target}`).root;
        drawingRels = readRels(parts, rel.target);
      } catch {
        hasUnimportedDrawings = true;
        continue;
      }

      const relMap = new Map(drawingRels.map(r => [r.id, r]));
      const anchors = childElements(drawingRoot).filter(el =>
        el.localName === 'twoCellAnchor' || el.localName === 'oneCellAnchor' || el.localName === 'absoluteAnchor');

      let drawingHasNonChartShape = false;
      let drawingImportedAChart = false;

      for (const anchorEl of anchors) {
        for (const child of childElements(anchorEl)) {
          if (child.localName && nonChartDrawingTags.has(child.localName)) {
            drawingHasNonChartShape = true;
          }
        }
        const chartEls = descendantElements(anchorEl, 'chart');
        for (const chartRefEl of chartEls) {
          const rid = relId(chartRefEl);
          const targetRel = rid ? relMap.get(rid) : undefined;
          if (!targetRel) {
            hasUnsupportedCharts = true;
            continue;
          }
          const chartKey = targetRel.target.toLowerCase();
          const chartBytes = parts.get(chartKey);
          if (!chartBytes) {
            hasUnsupportedCharts = true;
            continue;
          }
          const placement = readAnchorPlacement(anchorEl, sheetMeta.id);
          const parsed = parseChartXml(nextDoc, chartBytes, placement, makeId);
          if (!parsed) {
            hasUnsupportedCharts = true;
            continue;
          }
          const existingCharts = nextDoc.charts ?? [];
          nextDoc = {
            ...nextDoc,
            charts: [...existingCharts, {
              ...parsed,
              placement: {
                ...parsed.placement,
                zIndex: existingCharts.filter(c => c.placement.sheetId === sheetMeta.id).length + 1,
              },
            }],
          };
          importedChartParts.add(chartKey);
          drawingImportedAChart = true;
        }
      }

      if (drawingHasNonChartShape) hasUnimportedDrawings = true;
      if (drawingImportedAChart && !drawingHasNonChartShape) {
        importedDrawingParts.add(drawingKey);
      }
    }
  }

  return {
    document: nextDoc,
    importedChartParts,
    importedDrawingParts,
    hasUnsupportedCharts,
    hasUnimportedDrawings,
  };
}
