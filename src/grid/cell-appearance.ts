import type { CellStyle, Coordinate, WorkbookName } from '../core/workbook';
import { extentContains, tableDesign } from '../core/workbook-names';

export function cellFont(style: CellStyle | undefined, fallback: string, scale = 1): { css: string; size: number; shift: number } {
  const normalSize = (style?.fontSize ?? 10.5) * 4 / 3 * scale;
  const size = style?.script ? normalSize * .75 : normalSize;
  const family = style?.fontFamily ? '"' + style.fontFamily.replace(/["\\]/gu, '\\$&') + '", ' : '';
  const weight = style?.bold ? '700 ' : '';
  const italic = style?.italic ? 'italic ' : '';
  const shift = style?.script ? normalSize * (style.script === 'superscript' ? -.25 : .25) : 0;
  return { css: italic + weight + size + 'px ' + family + fallback, size, shift };
}

export function drawCellText(context: CanvasRenderingContext2D, text: string, style: CellStyle | undefined,
  fallback: string, scale: number, x: number, y: number, alignment: 'left' | 'right' | 'center', bottom: number) {
  const { css, size, shift } = cellFont(style, fallback, scale);
  context.font = css;
  context.textAlign = alignment;
  context.textBaseline = 'middle';
  context.fillText(text, x, y + shift);
  if (!text || !style || !(style.underline || style.strikethrough)) return;
  const metrics = context.measureText(text);
  const width = metrics.width;
  const start = alignment === 'right' ? x - width : alignment === 'center' ? x - width / 2 : x;
  const thickness = Math.max(1, size / 15);
  if (style.strikethrough) context.fillRect(start, y + shift - size * .05, width, thickness);
  if (!style.underline) return;
  const double = style.doubleUnderline || style.underlineKind === 'doubleAccounting';
  const belowText = y + shift + Math.max(size * .3, metrics.actualBoundingBoxDescent + thickness);
  const underline = Math.min(bottom - thickness, belowText + (style.underlineKind ? size * .14 : 0));
  context.fillRect(start, underline, width, thickness);
  if (double) context.fillRect(start, underline - 2 * thickness, width, thickness);
}

interface TextBox { readonly x: number; readonly y: number; readonly width: number; readonly height: number }

export function drawAlignedText(ctx: CanvasRenderingContext2D, text: string, style: CellStyle | undefined,
  fallback: string, scale: number, box: TextBox, valueKind: string) {
  if (!text) return;
  const align = style?.alignment ?? (valueKind === 'number' || style?.readingOrder === 'rtl' ? 'right' :
    valueKind === 'boolean' || valueKind === 'error' ? 'center' : 'left');
  const vertical = style?.verticalAlignment ?? 'bottom';
  const normal = cellFont(style, fallback, scale);
  const padding = 3 * scale;
  const indent = (style?.indent ?? 0) * normal.size * 1.5;
  const left = box.x + padding + (style?.alignment === 'left' || align === 'distributed' ? indent : 0);
  const right = box.x + box.width - padding - (style?.alignment === 'right' || align === 'distributed' ? indent : 0);
  const top = box.y + padding;
  const bottom = box.y + box.height - padding;
  const width = right - left, height = bottom - top;
  if (width <= 0 || height <= 0) return;
  const stacked = style?.textRotation === 255;
  const angle = stacked ? 0 : -(style?.textRotation ?? 0) * Math.PI / 180;
  const wrap = !stacked && !angle && (style?.wrap || align === 'justify' || align === 'distributed' ||
    vertical === 'justify' || vertical === 'distributed');
  const horizontal = align === 'right' ? 'right' :
    ['center', 'centerContinuous', 'distributed'].includes(align) ? 'center' : 'left';
  ctx.direction = style?.readingOrder === 'rtl' ? 'rtl' : style?.readingOrder === 'ltr' ? 'ltr' :
    /^\s*[\u0590-\u08ff]/u.test(text) ? 'rtl' : 'ltr';
  ctx.font = normal.css;
  ctx.textBaseline = 'middle';
  const measure = (value: string) => ctx.measureText(value).width;
  const lines = stacked ? Array.from(text.replace(/\r\n?/gu, '\n'), glyph => glyph === '\n' ? '' : glyph) :
    text.split(/\r\n|\r|\n/u).flatMap(line => wrap ? wrapText(line, width, measure) : [line]);
  let fontScale = scale;
  const baseHeight = normal.size * 1.18;
  const textWidth = lines.reduce((maximum, line) => Math.max(maximum, measure(line)), 0);
  const cos = Math.abs(Math.cos(angle)), sin = Math.abs(Math.sin(angle));
  const projectedWidth = textWidth * cos + baseHeight * lines.length * sin;
  if (!stacked && !wrap && style?.shrinkToFit && align !== 'fill' && projectedWidth > width) fontScale *= width / projectedWidth;
  const shrink = fontScale / scale;
  const lineHeight = baseHeight * shrink;
  const textHeight = lineHeight * lines.length;
  const verticalOrigin = (size: number) => {
    const extra = Math.max(0, height - size);
    return top + (vertical === 'top' || vertical === 'justify' ? 0 :
      vertical === 'center' || vertical === 'distributed' ? extra / 2 : extra);
  };
  const x = horizontal === 'right' ? right : horizontal === 'center' ? left + width / 2 : left;
  if (stacked) {
    // Number formats can add invisible padding. Align the painted glyphs, including decimal separators.
    let first = -1, last = -1;
    let inkTop = Infinity, inkBottom = -Infinity;
    for (const [index, glyph] of lines.entries()) {
      const { actualBoundingBoxAscent: ascent, actualBoundingBoxDescent: descent } = ctx.measureText(glyph);
      if (ascent + descent <= 0) continue;
      if (first < 0) first = index;
      last = index;
      const baseline = index * lineHeight + normal.shift;
      inkTop = Math.min(inkTop, baseline - ascent);
      inkBottom = Math.max(inkBottom, baseline + descent);
    }
    if (first < 0) return;
    const inkHeight = inkBottom - inkTop;
    const slots = last - first + 1;
    const extra = Math.max(0, height - inkHeight);
    const gap = vertical === 'distributed' ? extra / slots :
      vertical === 'justify' && slots > 1 ? extra / (slots - 1) : 0;
    const inkStart = vertical === 'distributed' ? top + gap / 2 :
      vertical === 'top' ? top : vertical === 'center' ? top + (height - inkHeight) / 2 :
        vertical === 'justify' ? top : bottom - inkHeight;
    for (let index = first; index <= last; index++) {
      const baseline = inkStart - inkTop + index * lineHeight + (index - first) * gap;
      drawCellText(ctx, lines[index]!, style, fallback, fontScale, x, baseline, horizontal, bottom);
    }
    return;
  }
  if (angle) {
    const rotatedWidth = projectedWidth * shrink;
    const rotatedHeight = textWidth * shrink * sin + textHeight * cos;
    const centerX = horizontal === 'right' ? x - rotatedWidth / 2 : horizontal === 'left' ? x + rotatedWidth / 2 : x;
    const words = lines[0]!.trim().split(/\s+/u);
    const spread = Math.abs(style?.textRotation ?? 0) === 90 && vertical === 'distributed' && lines.length === 1 &&
      words.length > 1 && words.reduce((sum, word) => sum + measure(word) * shrink, 0) <= height;
    ctx.translate(centerX, spread ? top + height / 2 : verticalOrigin(rotatedHeight) + rotatedHeight / 2);
    ctx.rotate(angle);
    if (spread && drawSpaced(lines[0]!, height, -height / 2, 0, false)) return;
    lines.forEach((line, index) => drawCellText(ctx, line, style, fallback, fontScale,
      0, lineHeight * (index + .5) - textHeight / 2, 'center', textHeight / 2 + lineHeight));
    return;
  }
  const extra = Math.max(0, height - textHeight);
  const gap = vertical === 'distributed' ? extra / lines.length :
    vertical === 'justify' && lines.length > 1 ? extra / (lines.length - 1) : 0;
  const start = vertical === 'distributed' ? top + gap / 2 : verticalOrigin(textHeight);
  lines.forEach((line, index) => {
    const baseline = start + (lineHeight + gap) * index + lineHeight / 2;
    if (align === 'fill' && line) {
      const count = Math.max(1, Math.min(Math.floor(width / (measure(line) || 1)), Math.floor(32768 / line.length)));
      drawCellText(ctx, line.repeat(count), style, fallback, scale, left, baseline, 'left', bottom);
      return;
    }
    if ((align === 'distributed' || align === 'justify' && (style?.justifyLastLine || index < lines.length - 1)) &&
        drawSpaced(line, width, left, baseline, align === 'distributed')) return;
    drawCellText(ctx, line, style, fallback, fontScale, x, baseline, horizontal, bottom);
  });

  function drawSpaced(line: string, area: number, start: number, baseline: number, distribute: boolean) {
    const words = line.trim().split(/\s+/u);
    const parts = words.length === 1 && distribute ? Array.from(line.trim()) : words;
    if (parts.length < 2) return false;
    ctx.font = cellFont(style, fallback, fontScale).css;
    const rtl = ctx.direction === 'rtl';
    if (rtl) parts.reverse();
    const widths = parts.map(measure);
    const spare = area - widths.reduce((sum, value) => sum + value, 0);
    if (spare < 0) return false;
    const space = spare / (parts.length - (distribute ? 0 : 1));
    let offset = start + (distribute ? space / 2 : 0);
    parts.forEach((part, index) => {
      drawCellText(ctx, part, style, fallback, fontScale, offset, baseline, 'left', baseline + lineHeight);
      offset += widths[index]! + space;
    });
    return true;
  }
}

function wrapText(text: string, width: number, measure: (line: string) => number): string[] {
  const lines: string[] = [];
  let line = '';
  for (const word of text.split(/(\s+)/u)) {
    if (measure(line + word) <= width) { line += word; continue; }
    if (line.trimEnd()) lines.push(line.trimEnd());
    line = '';
    if (word.trim() === '') continue;
    if (measure(word) <= width) { line = word; continue; }
    for (const glyph of word) {
      if (line && measure(line + glyph) > width) { lines.push(line); line = ''; }
      line += glyph;
    }
  }
  lines.push(line.trimEnd());
  return lines;
}

function patternInk(pattern: string, x: number, y: number): boolean {
  const width = pattern.startsWith('dark') ? 2 : 1;
  const vertical = x % 4 < width;
  const horizontal = y % 4 < width;
  const down = (x - y + 8) % 4 < width;
  const up = (x + y) % 4 < width;
  switch (pattern) {
    case 'darkGray': return (x + y) % 4 !== 0;
    case 'mediumGray': return (x + y) % 2 === 0;
    case 'lightGray': return x % 2 === 0 && y % 2 === 0;
    case 'gray125': return x % 4 === 0 && y % 2 === 0;
    case 'gray0625': return x % 4 === 0 && y % 4 === 0;
    case 'darkHorizontal': case 'lightHorizontal': return horizontal;
    case 'darkVertical': case 'lightVertical': return vertical;
    case 'darkDown': case 'lightDown': return down;
    case 'darkUp': case 'lightUp': return up;
    case 'darkGrid': case 'lightGrid': return vertical || horizontal;
    case 'darkTrellis': case 'lightTrellis': return down || up;
    default: return false;
  }
}

const patterns = new WeakMap<CanvasRenderingContext2D, Map<string, CanvasPattern>>();
function canvasPattern(context: CanvasRenderingContext2D, name: string, ink: string, paper: string) {
  let cache = patterns.get(context);
  if (!cache) patterns.set(context, cache = new Map());
  const key = name + ink + paper;
  const saved = cache.get(key);
  if (saved) return saved;
  const tile = context.canvas.ownerDocument.createElement('canvas');
  tile.width = tile.height = 8;
  const pixels = tile.getContext('2d')!;
  pixels.fillStyle = '#' + paper;
  pixels.fillRect(0, 0, 8, 8);
  pixels.fillStyle = '#' + ink;
  for (let row = 0; row < 8; row++) {
    for (let column = 0; column < 8; column++) {
      if (patternInk(name, column, row)) pixels.fillRect(column, row, 1, 1);
    }
  }
  const pattern = context.createPattern(tile, 'repeat')!;
  if (cache.size >= 128) cache.delete(cache.keys().next().value!);
  cache.set(key, pattern);
  return pattern;
}

function pathGradient(context: CanvasRenderingContext2D, x: number, y: number, width: number, height: number,
  center: boolean, variant: number, first: string, second: string) {
  const channels = (rgb: string) => [0, 2, 4].map(index => Number.parseInt(rgb.slice(index, index + 2), 16));
  const reversed = center && variant % 2 === 1;
  const start = channels(reversed ? second : first), end = channels(reversed ? first : second);
  for (let step = 64; step > 0; step--) {
    const t = step / 64, w = width * t, h = height * t;
    const left = x + (center ? (width - w) / 2 : variant % 2 ? width - w : 0);
    const top = y + (center ? (height - h) / 2 : variant >= 2 ? height - h : 0);
    context.fillStyle = 'rgb(' + start.map((value, index) => Math.round(value + (end[index]! - value) * t)).join(',') + ')';
    context.fillRect(left, top, w, h);
  }
}

export function drawCellFill(context: CanvasRenderingContext2D, style: CellStyle | undefined, x: number, y: number,
  width: number, height: number, scale = 1, fallback?: string) {
  const fill = style?.fillDetails;
  const paper = style?.fillRgb ?? 'FFFFFF';
  const shading = fill?.shading;
  if (shading) {
    const variant = fill?.variant ?? 0;
    const end = fill?.gradientRgb ?? '5B9BD5';
    if (shading === 'center' || shading === 'corner') {
      pathGradient(context, x, y, width, height, shading === 'center', variant, paper, end);
      return;
    }
    const dx = shading === 'horizontal' ? 0 : 1;
    const dy = shading === 'diagonalUp' ? -1 : shading === 'vertical' ? 0 : 1;
    const radius = (width * dx + height * Math.abs(dy)) / (2 * (dx * dx + dy * dy));
    const cx = x + width / 2, cy = y + height / 2;
    const gradient = context.createLinearGradient(cx - dx * radius, cy - dy * radius, cx + dx * radius, cy + dy * radius);
    const [first, second] = variant % 2 ? [end, paper] : [paper, end];
    gradient.addColorStop(0, '#' + first);
    if (variant >= 2) gradient.addColorStop(.5, '#' + second);
    gradient.addColorStop(1, '#' + (variant >= 2 ? first : second));
    context.fillStyle = gradient;
  } else if (fill && fill.pattern !== 'none') {
    const pattern = canvasPattern(context, fill.pattern, fill.patternRgb ?? '000000', paper);
    pattern.setTransform(new DOMMatrix([scale, 0, 0, scale, x, y]));
    context.fillStyle = pattern;
  } else {
    if (!style?.fillRgb && !fallback) return;
    context.fillStyle = style?.fillRgb ? '#' + style.fillRgb : fallback!;
  }
  context.fillRect(x, y, width, height);
}

export function usesLightPaper(style?: CellStyle): boolean {
  const effect = style?.fillDetails;
  const paper = style?.fillRgb ?? (effect && (effect.shading || effect.pattern !== 'none') ? 'FFFFFF' : '');
  if (!paper) return false;
  const channels = [0, 2, 4].map(index => Number.parseInt(paper.slice(index, index + 2), 16));
  return channels[0]! * .2126 + channels[1]! * .7152 + channels[2]! * .0722 > 170;
}

export function effectiveCellStyle(
  style: CellStyle | undefined,
  point: Coordinate,
  tables: readonly WorkbookName[],
): CellStyle | undefined {
  if (!tables.length) return style;
  const table = tables.find(t => !t.area.invalid && extentContains(t.area, point));
  if (!table) return style;
  const design = tableDesign(table);
  const area = table.area;
  const header = point.row === area.row && design.headerRow;
  const total = point.row === area.row + area.rows - 1 && design.totalRow;
  const strong = header || total || (design.firstColumn && point.column === area.column) ||
    (design.lastColumn && point.column === area.column + area.columns - 1);
  const stripedRow = design.bandedRows && ((point.row - area.row - (design.headerRow ? 1 : 0)) % 2 !== 0);
  const striped = stripedRow || (design.bandedColumns && ((point.column - area.column) % 2 === 1));

  let accent: string, plain: string, band: string;
  switch (design.style) {
    case 'Blue':
      accent = '4F81BD'; plain = 'DCE6F1'; band = 'B8CCE4'; break;
    case 'Green':
      accent = '548235'; plain = 'E2EFDA'; band = 'C6DFB2'; break;
    case 'Orange':
      accent = 'C65911'; plain = 'FCE4D6'; band = 'F8CBAD'; break;
    case 'Purple':
      accent = '8064A2'; plain = 'E4DFEC'; band = 'CCC0DA'; break;
    default:
      accent = '303844'; plain = '23262B'; band = '2A3037'; break;
  }
  const fill = design.style === 'Neutral'
    ? (header || total ? accent : stripedRow ? band : plain)
    : (strong ? accent : striped ? band : plain);

  const fillRgb = style?.fillRgb ?? (style?.fillDetails ? undefined : fill);
  const textRgb = style?.textRgb ?? (strong || design.style === 'Neutral' ? 'F5F5F5' : '172B4D');
  const bold = Boolean(style?.bold || strong);

  return {
    ...style,
    ...(fillRgb && { fillRgb }),
    textRgb,
    bold,
  };
}
