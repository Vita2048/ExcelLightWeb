import type { Element } from '@xmldom/xmldom';
import type { CellBorder, CellFill, CellStyle } from '../../core/workbook';
import { borderStyles } from '../../core/formats/borders';
import { fillPatterns, type FillShading } from '../../core/formats/appearance';
import { attr, childElements, escapeAttribute, firstChild, parseInvariantNumber, parseXmlPart } from './xlsx-xml';
import { exchangeCodes, XlsxReport } from './xlsx-report';

/** ECMA/ISO built-in number formats. Localized IDs without an explicit code are reported on import. */
const builtInFormats: Readonly<Record<number, string>> = Object.freeze({
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
  20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss',
  47: 'mmss.0', 48: '##0.0E+0', 49: '@',
});
const builtInByCode = new Map(Object.entries(builtInFormats).map(([id, code]) => [code, Number(id)]));
const ooxmlBorderStyles = new Set(Object.keys(borderStyles).filter(name => name !== 'none'));

type ColorRead = { readonly rgb: string } | 'default' | 'loss';

function readColor(element: Element | undefined, report: XlsxReport, code: string, detail: string,
  defaultTheme: boolean): ColorRead {
  if (!element) return 'default';
  const rgb = attr(element, 'rgb');
  if (rgb) {
    const match = /^(?:[0-9a-fA-F]{2})?([0-9a-fA-F]{6})$/u.exec(rgb);
    if (match?.[1]) return { rgb: match[1].toUpperCase() };
    report.add(code, detail);
    return 'loss';
  }
  if (attr(element, 'auto')) return 'default';
  const theme = attr(element, 'theme');
  const indexed = attr(element, 'indexed');
  if ((theme === '1' && defaultTheme) || indexed === '64') return 'default';
  if (theme !== undefined || indexed !== undefined) {
    report.add(code, detail);
    return 'loss';
  }
  return 'default';
}

function flag(element: Element | undefined): boolean {
  if (!element) return false;
  const value = attr(element, 'val');
  return value !== '0' && value !== 'false';
}

type UnderlineStyle = 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

interface FontRecord {
  readonly bold: boolean;
  readonly italic: boolean;
  readonly strikethrough: boolean;
  readonly underline: UnderlineStyle;
  readonly script?: 'superscript' | 'subscript';
  readonly family?: string;
  readonly size?: number;
  readonly color?: string;
  readonly shadow: boolean;
}

function readFont(element: Element, report: XlsxReport): FontRecord {
  const underlineText = attr(firstChild(element, 'u'), 'val') ?? (firstChild(element, 'u') ? 'single' : 'none');
  const underline: UnderlineStyle = ['single', 'double', 'singleAccounting', 'doubleAccounting'].includes(underlineText)
    ? underlineText as UnderlineStyle : underlineText === 'none' ? 'none' : 'single';
  const vertAlign = attr(firstChild(element, 'vertAlign'), 'val');
  const rawName = attr(firstChild(element, 'name'), 'val');
  const rawSize = attr(firstChild(element, 'sz'), 'val');
  let size: number | undefined;
  if (rawSize !== undefined) {
    try {
      size = parseInvariantNumber(rawSize, 'Font size is invalid.');
    } catch {
      size = undefined;
    }
    if (size !== undefined && (size < 6 || size > 72)) {
      report.add(exchangeCodes.fontSize, 'Font sizes outside 6-72 points were clamped.');
      size = Math.min(72, Math.max(6, size));
    }
  }
  const color = readColor(firstChild(element, 'color'), report, exchangeCodes.themeFont,
    'Non-default theme/indexed font colors are omitted; explicit RGB colors are preserved.', true);
  return {
    bold: flag(firstChild(element, 'b')),
    italic: flag(firstChild(element, 'i')),
    strikethrough: flag(firstChild(element, 'strike')),
    underline,
    ...(vertAlign === 'superscript' || vertAlign === 'subscript' ? { script: vertAlign } : {}),
    ...(rawName && rawName !== 'Calibri' ? { family: rawName } : {}),
    ...(size !== undefined && Math.abs(size - 11) >= .001 ? { size } : {}),
    ...(color === 'default' || color === 'loss' ? {} : { color: color.rgb }),
    shadow: flag(firstChild(element, 'shadow')),
  };
}

interface FillRecord {
  readonly fillRgb?: string;
  readonly fillDetails?: CellFill;
  readonly gradient?: { readonly fillRgb: string; readonly fillDetails: CellFill };
}

function stopRgb(stop: Element): string | undefined {
  const value = attr(firstChild(stop, 'color'), 'rgb');
  if (!value) return undefined;
  const match = /^(?:[0-9a-fA-F]{2})?([0-9a-fA-F]{6})$/u.exec(value);
  return match?.[1]?.toUpperCase();
}

/** Desktop XlsxGradients.ReadGradientCells: only the canonical 2/3-stop shapes are preserved. */
function readGradient(element: Element, report: XlsxReport): { fillRgb: string; fillDetails: CellFill } | undefined {
  const stops = childElements(element, 'stop');
  const colors = stops.map(stopRgb);
  const positions = stops.map(stop => {
    const value = attr(stop, 'position');
    if (value === undefined) return 0;
    try { return parseInvariantNumber(value, 'Gradient stop position is invalid.'); } catch { return NaN; }
  });
  if ((stops.length !== 2 && stops.length !== 3) || colors.some(color => !color) ||
    positions[0] !== 0 || positions.at(-1) !== 1 ||
    (stops.length === 3 && (positions[1] !== .5 || colors[2] !== colors[0]))) {
    report.add(exchangeCodes.gradientFill, 'Unsupported gradient stops/colors were omitted.');
    return undefined;
  }
  const color = colors[0]!;
  const second = colors[1]!;
  let shading: FillShading;
  let variant: number;
  if (attr(element, 'type') === 'path') {
    const left = numericAttribute(element, 'left', 0);
    const top = numericAttribute(element, 'top', 0);
    shading = (left === .5 || left === .25) && left === top ? 'center' : 'corner';
    variant = shading === 'center' ? (left === .25 ? 2 : 0) : (left >= .5 ? 1 : 0) + (top >= .5 ? 2 : 0);
  } else {
    const degree = ((numericAttribute(element, 'degree', 0) % 360) + 360) % 360;
    const direction = degree % 180;
    shading = direction === 90 ? 'horizontal' : direction === 0 ? 'vertical' :
      direction === 45 ? 'diagonalDown' : direction === 135 ? 'diagonalUp' : 'vertical';
    if (direction !== 0 && direction !== 45 && direction !== 90 && direction !== 135) {
      report.add(exchangeCodes.gradientFill, 'Nonstandard gradient angle was approximated.');
    }
    variant = (degree >= 180 ? 1 : 0) + (stops.length === 3 ? 2 : 0);
  }
  return { fillRgb: color, fillDetails: { pattern: 'none', gradientRgb: second, shading, variant } };
}

function numericAttribute(element: Element, name: string, fallback: number): number {
  const value = attr(element, name);
  if (value === undefined) return fallback;
  try { return parseInvariantNumber(value, 'XLSX style attribute is invalid.'); } catch { return fallback; }
}

function readFill(element: Element, report: XlsxReport): FillRecord {
  const gradient = firstChild(element, 'gradientFill');
  if (gradient) {
    const parsed = readGradient(gradient, report);
    return parsed ? { fillRgb: parsed.fillRgb, fillDetails: parsed.fillDetails, gradient: parsed } : {};
  }
  const pattern = firstChild(element, 'patternFill');
  if (!pattern) return {};
  const patternType = attr(pattern, 'patternType');
  if (!patternType || patternType === 'none') return {};
  const foreground = readColor(firstChild(pattern, 'fgColor'), report, exchangeCodes.themeFill,
    'A non-RGB fill was omitted.', false);
  const background = readColor(firstChild(pattern, 'bgColor'), report, exchangeCodes.themeFill,
    'A non-RGB fill was omitted.', false);
  if (patternType === 'solid') {
    const solid = foreground !== 'default' && foreground !== 'loss' ? foreground.rgb :
      background !== 'default' && background !== 'loss' ? background.rgb : undefined;
    return solid ? { fillRgb: solid } : {};
  }
  if (!Object.hasOwn(fillPatterns, patternType)) {
    report.add(exchangeCodes.patternFill, 'Unsupported fill patterns were omitted.');
    return {};
  }
  const patternRgb = foreground !== 'default' && foreground !== 'loss' ? foreground.rgb : undefined;
  const backgroundRgb = background !== 'default' && background !== 'loss' ? background.rgb : undefined;
  if (!patternRgb && !backgroundRgb) return {};
  return {
    ...(backgroundRgb ? { fillRgb: backgroundRgb } : {}),
    fillDetails: { pattern: patternType, ...(patternRgb ? { patternRgb } : {}) },
  };
}

interface BorderRecord {
  readonly edges: Readonly<Partial<Record<'top' | 'right' | 'bottom' | 'left', CellBorder>>>;
  readonly diagonal?: CellBorder;
  readonly diagonalUp: boolean;
  readonly diagonalDown: boolean;
}

function readBorder(element: Element, report: XlsxReport): BorderRecord {
  const edges: Partial<Record<'top' | 'right' | 'bottom' | 'left', CellBorder>> = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const record = firstChild(element, side);
    const pen = record ? readPen(record, report) : undefined;
    if (pen) edges[side] = pen;
  }
  const diagonalRecord = firstChild(element, 'diagonal');
  const diagonal = diagonalRecord ? readPen(diagonalRecord, report) : undefined;
  return {
    edges,
    ...(diagonal ? { diagonal } : {}),
    diagonalUp: attr(element, 'diagonalUp') === '1' || attr(element, 'diagonalUp') === 'true',
    diagonalDown: attr(element, 'diagonalDown') === '1' || attr(element, 'diagonalDown') === 'true',
  };
}

function readPen(element: Element, report: XlsxReport): CellBorder | undefined {
  const style = attr(element, 'style');
  if (style === undefined || style === 'none') return undefined;
  if (!ooxmlBorderStyles.has(style)) {
    report.add(exchangeCodes.borderStyles, 'Unsupported border line styles were omitted.');
    return undefined;
  }
  const color = readColor(firstChild(element, 'color'), report, exchangeCodes.borderColors,
    'Theme/indexed border colors are replaced by Automatic; explicit RGB colors are retained.', true);
  return { style, ...(color !== 'default' && color !== 'loss' ? { rgb: color.rgb } : {}) };
}

interface XfRecord {
  readonly fontId: number;
  readonly fillId: number;
  readonly borderId: number;
  readonly numFmtId: number;
  readonly alignment?: Element;
}

export interface ReadStyles {
  cellStyle(index: number | undefined): CellStyle | undefined;
  isDateFormatted(index: number | undefined): boolean;
  numberFormatForId(numFmtId: number | undefined): string | undefined;
}

const emptyStyles: ReadStyles = Object.freeze({
  cellStyle: () => undefined,
  isDateFormatted: () => false,
  numberFormatForId: () => undefined,
});

export function readStyles(bytes: Uint8Array | undefined, report: XlsxReport): ReadStyles {
  if (!bytes) return emptyStyles;
  const { root } = parseXmlPart(bytes, 'The XLSX styles part');
  const numFmts = new Map<number, string>();
  for (const record of childElements(firstChild(root, 'numFmts'), 'numFmt')) {
    const id = Number(attr(record, 'numFmtId'));
    const code = attr(record, 'formatCode');
    if (Number.isInteger(id) && code !== undefined) numFmts.set(id, code);
  }
  const fonts = childElements(firstChild(root, 'fonts'), 'font').map(font => readFont(font, report));
  const fills = childElements(firstChild(root, 'fills'), 'fill').map(fill => readFill(fill, report));
  const borders = childElements(firstChild(root, 'borders'), 'border').map(border => readBorder(border, report));
  const xfs: XfRecord[] = childElements(firstChild(root, 'cellXfs'), 'xf').map(xf => ({
    fontId: numericAttribute(xf, 'fontId', 0),
    fillId: numericAttribute(xf, 'fillId', 0),
    borderId: numericAttribute(xf, 'borderId', 0),
    numFmtId: numericAttribute(xf, 'numFmtId', 0),
    ...(firstChild(xf, 'alignment') ? { alignment: firstChild(xf, 'alignment')! } : {}),
  }));

  const cached = new Map<number, CellStyle | undefined>();
  const numberFormatForId = (numFmtId: number | undefined, warn = false): string | undefined => {
    if (numFmtId === undefined || numFmtId === 0) return undefined;
    const explicit = numFmts.get(numFmtId);
    if (explicit !== undefined) return explicit === 'General' ? undefined : explicit;
    const builtIn = builtInFormats[numFmtId];
    if (builtIn === undefined) {
      if (warn) {
        report.add(exchangeCodes.numberFormat, 'An unsupported localized built-in number format was omitted; raw numeric values are retained.');
      }
      return undefined;
    }
    return builtIn === 'General' ? undefined : builtIn;
  };
  const formatFor = (index: number | undefined): string | undefined => {
    const xf = index === undefined ? undefined : xfs[index];
    return xf ? numberFormatForId(xf.numFmtId, true) : undefined;
  };

  function cellStyle(index: number | undefined): CellStyle | undefined {
    if (index === undefined) return undefined;
    if (cached.has(index)) return cached.get(index);
    const xf = xfs[index];
    if (!xf) return undefined;
    const font = fonts[xf.fontId] ?? fonts[0];
    const fill = fills[xf.fillId];
    const border = borders[xf.borderId];
    const alignment = xf.alignment;
    const numberFormat = formatFor(index);
    const horizontal = attr(alignment, 'horizontal');
    const vertical = attr(alignment, 'vertical');
    const readingOrder = attr(alignment, 'readingOrder');
    const textRotationRaw = alignment ? numericAttribute(alignment, 'textRotation', 0) : 0;
    const textRotation = textRotationRaw === 255 ? 255 : textRotationRaw > 90 ? 90 - textRotationRaw : textRotationRaw;
    const indent = alignment ? numericAttribute(alignment, 'indent', 0) : 0;
    const legacyBorders = border ? legacyBorderText(border, font?.color) : undefined;
    const style: CellStyle = {
      ...(font?.bold ? { bold: true } : {}),
      ...(numberFormat ? { numberFormat } : {}),
      ...(fill?.fillRgb ? { fillRgb: fill.fillRgb } : {}),
      ...(font?.family ? { fontFamily: font.family } : {}),
      ...(font?.size !== undefined ? { fontSize: font.size } : {}),
      ...(font?.italic ? { italic: true } : {}),
      ...(font && font.underline !== 'none' ? { underline: true } : {}),
      ...(font?.color ? { textRgb: font.color } : {}),
      ...(legacyBorders ? { borders: legacyBorders } : {}),
      ...(horizontal && ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'].includes(horizontal)
        ? { alignment: horizontal } : {}),
      ...(attr(alignment, 'wrapText') === '1' || attr(alignment, 'wrapText') === 'true' ? { wrap: true } : {}),
      ...(font?.underline === 'double' || font?.underline === 'doubleAccounting' ? { doubleUnderline: true } : {}),
      ...(vertical && ['top', 'center', 'justify', 'distributed'].includes(vertical)
        ? { verticalAlignment: vertical } : {}),
      ...(font?.strikethrough ? { strikethrough: true } : {}),
      ...(font?.script ? { script: font.script } : {}),
      ...(font?.underline === 'singleAccounting' || font?.underline === 'doubleAccounting'
        ? { underlineKind: font.underline } : {}),
      ...(fill?.fillDetails ? { fillDetails: fill.fillDetails } : {}),
      ...(!legacyBorders && border && Object.keys(borderDetails(border)).length > 0
        ? { borderDetails: borderDetails(border) } : {}),
      ...(textRotation !== 0 ? { textRotation } : {}),
      ...(indent !== 0 ? { indent } : {}),
      ...(attr(alignment, 'shrinkToFit') === '1' || attr(alignment, 'shrinkToFit') === 'true' ? { shrinkToFit: true } : {}),
      ...(readingOrder === '1' ? { readingOrder: 'ltr' } : readingOrder === '2' ? { readingOrder: 'rtl' } : {}),
      ...(attr(alignment, 'justifyLastLine') === '1' || attr(alignment, 'justifyLastLine') === 'true'
        ? { justifyLastLine: true } : {}),
    };
    const result = Object.keys(style).length > 0 ? style : undefined;
    if (font?.shadow) report.add(exchangeCodes.fontEffects, 'Unsupported font effects are omitted.');
    cached.set(index, result);
    return result;
  }

  function isDateFormatted(index: number | undefined): boolean {
    const xf = index === undefined ? undefined : xfs[index];
    if (!xf) return false;
    if (xf.numFmtId >= 14 && xf.numFmtId <= 17) return true;
    const format = formatFor(index) ?? '';
    const stripped = format.replace(/"[^"]*"|\\./gu, '').toLowerCase();
    return stripped.includes('y') || stripped.includes('d');
  }

  return { cellStyle, isDateFormatted, numberFormatForId: (id: number | undefined) => numberFormatForId(id, false) };
}

function borderDetails(border: BorderRecord): NonNullable<CellStyle['borderDetails']> {
  const details: Partial<Record<'top' | 'right' | 'bottom' | 'left' | 'diagonalUp' | 'diagonalDown', CellBorder>> = {
    ...border.edges,
  };
  if (border.diagonal && border.diagonalUp) details.diagonalUp = border.diagonal;
  if (border.diagonal && border.diagonalDown) details.diagonalDown = border.diagonal;
  return details;
}

/** Desktop legacy `borders` letters: four thin edges inked like the font color, with no diagonals. */
function legacyBorderText(border: BorderRecord, fontColor: string | undefined): string | undefined {
  if (border.diagonalUp || border.diagonalDown) return undefined;
  const keys: ('top' | 'right' | 'bottom' | 'left')[] = ['top', 'right', 'bottom', 'left'];
  const text = keys.map((side, index) => {
    const edge = border.edges[side];
    if (!edge) return '';
    if (edge.style !== 'thin' || (edge.rgb ?? '000000') !== (fontColor ?? '000000')) return undefined;
    return 'trbl'[index]!;
  });
  if (text.some(value => value === undefined)) return undefined;
  return text.join('') || undefined;
}

/* ---------------------------------------------------------------- export */

const sideNames = ['left', 'right', 'top', 'bottom'] as const;

export interface StyleWriter {
  /** Register a cell style and return its cellXfs index. Deduplicates fonts, fills, borders and formats. */
  indexOf(style: CellStyle | undefined): number;
  /** Serialize the styles part after all cell styles were registered. */
  toXml(): string;
}

export function createStyleWriter(): StyleWriter {
  const numFmts: string[] = [];
  const numFmtIndex = new Map<string, number>();
  const fonts: string[] = ['<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>'];
  const fontIndex = new Map<string, number>([['|||||Calibri|11|', 0]]);
  const fills: string[] = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
  ];
  const fillIndex = new Map<string, number>();
  const borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const borderIndex = new Map<string, number>();
  const xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const xfIndex = new Map<string, number>([['', 0]]);

  const numberFormatId = (code: string): number => {
    const builtIn = builtInByCode.get(code);
    if (builtIn !== undefined) return builtIn;
    const existing = numFmtIndex.get(code);
    if (existing !== undefined) return existing;
    const id = 164 + numFmts.length;
    numFmts.push(`<numFmt numFmtId="${id}" formatCode="${escapeAttribute(code)}"/>`);
    numFmtIndex.set(code, id);
    return id;
  };

  const fontId = (style: CellStyle): number => {
    const underline = style.underlineKind === 'singleAccounting' ? '<u val="singleAccounting"/>' :
      style.underlineKind === 'doubleAccounting' ? '<u val="doubleAccounting"/>' :
      style.underline ? style.doubleUnderline ? '<u val="double"/>' : '<u/>' : '';
    const key = [
      style.bold ? 'b' : '', style.italic ? 'i' : '', style.strikethrough ? 's' : '',
      underline, style.script ?? '', style.fontFamily ?? 'Calibri', String(style.fontSize ?? 11),
      style.textRgb ?? '',
    ].join('|');
    const existing = fontIndex.get(key);
    if (existing !== undefined) return existing;
    const id = fonts.length;
    fonts.push('<font>' + (style.bold ? '<b/>' : '') + (style.italic ? '<i/>' : '') +
      (style.strikethrough ? '<strike/>' : '') + underline +
      (style.script === 'superscript' ? '<vertAlign val="superscript"/>' :
        style.script === 'subscript' ? '<vertAlign val="subscript"/>' : '') +
      `<sz val="${style.fontSize ?? 11}"/>` +
      (style.textRgb ? `<color rgb="FF${style.textRgb}"/>` : '<color theme="1"/>') +
      `<name val="${escapeAttribute(style.fontFamily ?? 'Calibri')}"/></font>`);
    fontIndex.set(key, id);
    return id;
  };

  const fillId = (style: CellStyle): number => {
    const details = style.fillDetails;
    if (details?.shading) {
      const key = 'gradient|' + [style.fillRgb ?? '', details.gradientRgb ?? '', details.shading,
        details.variant ?? 0].join('|');
      const existing = fillIndex.get(key);
      if (existing !== undefined) return existing;
      const id = fills.length;
      fills.push(gradientFill(style.fillRgb ?? 'FFFFFF', details.gradientRgb ?? 'FFFFFF',
        details.shading as FillShading, details.variant ?? 0));
      fillIndex.set(key, id);
      return id;
    }
    if (details && details.pattern !== 'none') {
      const key = 'pattern|' + [details.pattern, details.patternRgb ?? '', style.fillRgb ?? ''].join('|');
      const existing = fillIndex.get(key);
      if (existing !== undefined) return existing;
      const id = fills.length;
      fills.push(`<fill><patternFill patternType="${details.pattern}">` +
        `<fgColor rgb="FF${details.patternRgb ?? '000000'}"/>` +
        `<bgColor rgb="FF${style.fillRgb ?? 'FFFFFF'}"/></patternFill></fill>`);
      fillIndex.set(key, id);
      return id;
    }
    if (!style.fillRgb) return 0;
    const key = 'solid|' + style.fillRgb;
    const existing = fillIndex.get(key);
    if (existing !== undefined) return existing;
    const id = fills.length;
    fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="FF${style.fillRgb}"/>` +
      `<bgColor rgb="FF${style.fillRgb}"/></patternFill></fill>`);
    fillIndex.set(key, id);
    return id;
  };

  const borderId = (style: CellStyle): number => {
    const edges: Partial<Record<typeof sideNames[number], CellBorder>> = {};
    const details = style.borderDetails;
    if (details) {
      for (const side of sideNames) if (details[side]) edges[side] = details[side]!;
    } else if (style.borders) {
      for (const side of sideNames) if (style.borders.includes(side[0]!)) edges[side] = { style: 'thin' };
    }
    const diagonal = details?.diagonalUp ?? details?.diagonalDown;
    const key = sideNames.map(side => {
      const pen = edges[side];
      return pen ? pen.style + (pen.rgb ?? '') : '';
    }).join('|') + '|' + (diagonal ? diagonal.style + (diagonal.rgb ?? '') +
      (details?.diagonalUp ? 'u' : '') + (details?.diagonalDown ? 'd' : '') : '');
    if (!key.replaceAll('|', '')) return 0;
    const existing = borderIndex.get(key);
    if (existing !== undefined) return existing;
    const pen = (element: 'left' | 'right' | 'top' | 'bottom') => {
      const edge = edges[element];
      return edge ? `<${element} style="${edge.style}">` +
        (edge.rgb ? `<color rgb="FF${edge.rgb}"/>` : '<color indexed="64"/>') + `</${element}>` :
        `<${element}/>`;
    };
    const id = borders.length;
    borders.push('<border' + (details?.diagonalUp ? ' diagonalUp="1"' : '') +
      (details?.diagonalDown ? ' diagonalDown="1"' : '') + '>' +
      pen('left') + pen('right') + pen('top') + pen('bottom') +
      (diagonal ? `<diagonal style="${diagonal.style}">` +
        (diagonal.rgb ? `<color rgb="FF${diagonal.rgb}"/>` : '<color indexed="64"/>') + '</diagonal>' :
        '<diagonal/>') + '</border>');
    borderIndex.set(key, id);
    return id;
  };

  const alignmentXml = (style: CellStyle): string => {
    const parts: string[] = [];
    if (style.alignment) parts.push(`horizontal="${style.alignment}"`);
    if (style.verticalAlignment) parts.push(`vertical="${style.verticalAlignment}"`);
    if (style.wrap) parts.push('wrapText="1"');
    if (style.textRotation !== undefined && style.textRotation !== 0) {
      parts.push(`textRotation="${style.textRotation === 255 ? 255 :
        style.textRotation < 0 ? 90 - style.textRotation : style.textRotation}"`);
    }
    if (style.indent) parts.push(`indent="${style.indent}"`);
    if (style.shrinkToFit) parts.push('shrinkToFit="1"');
    if (style.readingOrder) parts.push(`readingOrder="${style.readingOrder === 'ltr' ? 1 : 2}"`);
    if (style.justifyLastLine) parts.push('justifyLastLine="1"');
    return parts.length ? `<alignment ${parts.join(' ')}/>` : '';
  };

  const indexOf = (style: CellStyle | undefined): number => {
    if (!style || Object.keys(style).length === 0) return 0;
    const key = JSON.stringify(style);
    const existing = xfIndex.get(key);
    if (existing !== undefined) return existing;
    const numFmtId = style.numberFormat ? numberFormatId(style.numberFormat) : 0;
    const font = fontId(style);
    const fill = fillId(style);
    const border = borderId(style);
    const alignment = alignmentXml(style);
    const id = xfs.length;
    xfs.push(`<xf numFmtId="${numFmtId}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"` +
      (numFmtId ? ' applyNumberFormat="1"' : '') + (font ? ' applyFont="1"' : '') +
      (fill ? ' applyFill="1"' : '') + (border ? ' applyBorder="1"' : '') +
      (alignment ? ' applyAlignment="1">' + alignment + '</xf>' : '/>'));
    xfIndex.set(key, id);
    return id;
  };

  const toXml = (): string => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<styleSheet xmlns="${'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}">` +
    (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>` : '') +
    `<fonts count="${fonts.length}">${fonts.join('')}</fonts>` +
    `<fills count="${fills.length}">${fills.join('')}</fills>` +
    `<borders count="${borders.length}">${borders.join('')}</borders>` +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>';
  return { indexOf, toXml };
}

/** Desktop XlsxGradients.GradientXml: standard fill XML for the supported shading styles. */
function gradientFill(firstRgb: string, secondRgb: string, shading: FillShading, variant: number): string {
  let attributes: string;
  if (shading === 'corner' || shading === 'center') {
    const x = shading === 'center' ? (variant >= 2 ? .25 : .5) : variant % 2;
    const y = shading === 'center' ? x : Math.floor(variant / 2);
    attributes = `type="path" left="${x}" right="${x}" top="${y}" bottom="${y}"`;
  } else {
    const base = shading === 'horizontal' ? 90 : shading === 'vertical' ? 0 : shading === 'diagonalDown' ? 45 : 135;
    attributes = `degree="${base + (variant % 2) * 180}"`;
  }
  let first = firstRgb;
  let second = secondRgb;
  if (shading === 'center' && variant % 2 === 1) [first, second] = [second, first];
  const stop = (position: string, rgb: string) =>
    `<stop position="${position}"><color rgb="FF${rgb}"/></stop>`;
  const stops = shading !== 'corner' && shading !== 'center' && variant >= 2
    ? stop('0', first) + stop('.5', second) + stop('1', first)
    : stop('0', first) + stop('1', second);
  return `<fill><gradientFill ${attributes}>${stops}</gradientFill></fill>`;
}
