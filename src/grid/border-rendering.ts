import type { CellBorder } from '../core/workbook';

const longDash = [6, 2, 1, 2];
const longDoubleDot = [6, 2, 1, 2, 1, 2];
const dashes: Readonly<Record<string, readonly number[]>> = {
  hair: [1, 2], dotted: [1, 2], dashed: [4, 2], mediumDashed: [4, 2],
  dashDot: longDash, mediumDashDot: longDash, slantDashDot: longDash,
  dashDotDot: longDoubleDot, mediumDashDotDot: longDoubleDot,
};

export function drawBorder(ctx: CanvasRenderingContext2D, pen: CellBorder, x: number, y: number, endX: number,
  endY: number, scale: number, fallback: string, forced: boolean) {
  const length = Math.hypot(endX - x, endY - y);
  if (!length) return;
  const weight = pen.style === 'hair' ? .5 : pen.style === 'thick' ? 3 : pen.style.startsWith('medium') ? 2 : 1;
  ctx.save();
  ctx.strokeStyle = !forced && pen.rgb ? '#' + pen.rgb : fallback;
  ctx.lineWidth = scale * weight;
  ctx.setLineDash((dashes[pen.style] ?? []).map(length => length * scale));
  const line = (dx: number, dy: number) => {
    ctx.moveTo(x + dx, y + dy);
    ctx.lineTo(endX + dx, endY + dy);
  };
  ctx.beginPath();
  if (pen.style === 'double') {
    const dx = (y - endY) / length * scale, dy = (endX - x) / length * scale;
    line(dx, dy);
    line(-dx, -dy);
  } else line(0, 0);
  ctx.stroke();
  ctx.restore();
}
