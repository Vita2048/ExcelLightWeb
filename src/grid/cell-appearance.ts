import type { CellStyle } from '../core/workbook';

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
