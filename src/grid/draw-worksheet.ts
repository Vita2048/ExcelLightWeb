import { findCell, type SheetDocument } from '../core/workbook';
import { drawCellFill, drawCellText, usesLightPaper } from './cell-appearance';
import { columnName, gridMetrics, type Track } from './grid-geometry';
import { presentCell, includes, type Area, type GridSelection } from './grid-selection';

export interface GridViewport {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

interface GridFrame {
  readonly sheet: SheetDocument;
  readonly view: GridViewport;
  readonly rows: readonly Track[];
  readonly columns: readonly Track[];
  readonly selection: GridSelection;
  readonly areas: readonly Area[];
  readonly scale: number;
  readonly headings: boolean;
  readonly gridlines: boolean;
  readonly showFormulas: boolean;
  readonly staleCells: ReadonlySet<string>;
  readonly separator: ',' | ';';
  readonly referenceArea?: Area | undefined;
}

export function drawWorksheet(canvas: HTMLCanvasElement, frame: GridFrame): void {
  const context = canvas.getContext('2d');
  const { width, height } = frame.view;
  if (!context || width === 0 || height === 0) return;
  const ratio = window.devicePixelRatio || 1;
  const pixelsWide = Math.round(width * ratio);
  const pixelsHigh = Math.round(height * ratio);
  if (canvas.width !== pixelsWide) canvas.width = pixelsWide;
  if (canvas.height !== pixelsHigh) canvas.height = pixelsHigh;
  context.setTransform(ratio, 0, 0, ratio, 0, 0);

  const styles = getComputedStyle(canvas);
  const token = (key: string) => styles.getPropertyValue(key).trim();
  const forced = matchMedia('(forced-colors: active)').matches;
  const sheetColor = forced ? 'Canvas' : token('--sheet');
  const headerColor = forced ? 'Canvas' : token('--grid-header');
  const gridColor = forced ? 'GrayText' : token('--gridline');
  const textColor = forced ? 'CanvasText' : token('--text');
  const accent = forced ? 'Highlight' : token('--accent');
  const tint = forced ? 'Highlight' : token('--grid-selection');
  const { rows, columns, selection, areas, scale, headings } = frame;
  const left = headings ? gridMetrics.rowHeader * scale : 0;
  const top = headings ? gridMetrics.columnHeader * scale : 0;
  const x = (track: Track) => left + track.start - frame.view.x;
  const y = (track: Track) => top + track.start - frame.view.y;
  const line = (position: number) => Math.round(position * ratio) / ratio - 1 / ratio;
  const font = (size: number, weight = '') => `${weight ? weight + ' ' : ''}${size * scale}px ${styles.fontFamily}`;

  context.fillStyle = sheetColor;
  context.fillRect(0, 0, width, height);
  context.save();
  context.beginPath();
  context.rect(left, top, Math.max(0, width - left), Math.max(0, height - top));
  context.clip();
  context.textBaseline = 'middle';

  for (const row of rows) {
    for (const column of columns) {
      const point = { row: row.index, column: column.index };
      const cell = findCell(frame.sheet, point);
      const selected = areas.some(area => includes(area, point));
      const active = point.row === selection.anchor.row && point.column === selection.anchor.column;
      const cellX = x(column);
      const cellY = y(row);
      if (!forced) drawCellFill(context, cell?.style, cellX, cellY, column.size, row.size, scale);
      if (selected && !active) {
        context.fillStyle = tint;
        context.fillRect(cellX, cellY, column.size, row.size);
      }
      if (!cell) continue;
      const { text: label, color } = presentCell(cell, frame.showFormulas, frame.separator);
      const value = frame.showFormulas ? cell.content : cell.cache ?? cell.content;
      context.save();
      context.beginPath();
      context.rect(cellX, cellY, column.size, row.size);
      context.clip();
      const right = value.kind === 'number';
      const authoredColor = color ?? cell.style?.textRgb;
      context.fillStyle = forced ? selected && !active ? 'HighlightText' : textColor :
        authoredColor ? '#' + authoredColor : usesLightPaper(cell.style) ? '#1b1b1b' : textColor;
      drawCellText(context, label, cell.style, styles.fontFamily, scale,
        right ? cellX + column.size - 5 * scale : cellX + 5 * scale, cellY + row.size / 2, right ? 'right' : 'left',
        cellY + row.size - 2 * scale);
      if (frame.staleCells.has(cell.id)) {
        context.strokeStyle = forced ? context.fillStyle : token('--stale-calculation');
        context.lineWidth = 1;
        context.setLineDash([3, 2]);
        context.beginPath();
        context.moveTo(cellX + 2, cellY + row.size - 3);
        context.lineTo(cellX + column.size - 2, cellY + row.size - 3);
        context.stroke();
      }
      context.restore();
    }
  }

  if (frame.gridlines) {
    context.fillStyle = gridColor;
    for (const row of rows) context.fillRect(left, line(y(row) + row.size), width - left, 1 / ratio);
    for (const column of columns) context.fillRect(line(x(column) + column.size), top, 1 / ratio, height - top);
  }
  const outline = (area: Area) => {
    const firstRow = rows.find(row => row.index >= area.top);
    const lastRow = lastTrack(rows, area.bottom);
    const firstColumn = columns.find(column => column.index >= area.left);
    const lastColumn = lastTrack(columns, area.right);
    if (!(firstRow && lastRow && firstColumn && lastColumn) ||
        firstRow.index > lastRow.index || firstColumn.index > lastColumn.index) return;
    const leftEdge = x(firstColumn), rightEdge = x(lastColumn) + lastColumn.size;
    const topEdge = y(firstRow), bottomEdge = y(lastRow) + lastRow.size;
    context.lineWidth = 1 / ratio;
    context.beginPath();
    if (firstRow.index === area.top) { context.moveTo(leftEdge, topEdge); context.lineTo(rightEdge, topEdge); }
    if (lastRow.index === area.bottom) { context.moveTo(leftEdge, bottomEdge); context.lineTo(rightEdge, bottomEdge); }
    if (firstColumn.index === area.left) { context.moveTo(leftEdge, topEdge); context.lineTo(leftEdge, bottomEdge); }
    if (lastColumn.index === area.right) { context.moveTo(rightEdge, topEdge); context.lineTo(rightEdge, bottomEdge); }
    context.stroke();
  };
  context.strokeStyle = accent;
  for (const area of areas) outline(area);
  const activeRow = rows.find(row => row.index === selection.anchor.row);
  const activeColumn = columns.find(column => column.index === selection.anchor.column);
  if (activeRow && activeColumn) {
    context.lineWidth = 2;
    context.strokeRect(x(activeColumn) + 1, y(activeRow) + 1, Math.max(0, activeColumn.size - 2), Math.max(0, activeRow.size - 2));
  }
  if (frame.referenceArea) {
    context.strokeStyle = forced ? 'CanvasText' : token('--formula-reference');
    context.setLineDash([4 * scale, 2 * scale]);
    outline(frame.referenceArea);
  }
  context.restore();

  if (!headings) return;
  context.fillStyle = headerColor;
  context.fillRect(0, 0, width, top);
  context.fillRect(0, top, left, height - top);
  context.font = font(12);
  context.textAlign = 'center';
  context.textBaseline = 'middle';
  context.save();
  context.beginPath();
  context.rect(left, 0, width - left, top);
  context.clip();
  for (const column of columns) {
    const start = x(column);
    const selected = areas.some(area => column.index >= area.left && column.index <= area.right);
    context.fillStyle = selected ? tint : headerColor;
    context.fillRect(start, 0, column.size, top);
    context.fillStyle = selected && forced ? 'HighlightText' : textColor;
    context.fillText(columnName(column.index), start + column.size / 2, top / 2);
    context.fillStyle = gridColor;
    context.fillRect(line(start + column.size), 0, 1 / ratio, top);
  }
  context.restore();
  context.save();
  context.beginPath();
  context.rect(0, top, left, height - top);
  context.clip();
  for (const row of rows) {
    const start = y(row);
    const selected = areas.some(area => row.index >= area.top && row.index <= area.bottom);
    context.fillStyle = selected ? tint : headerColor;
    context.fillRect(0, start, left, row.size);
    context.fillStyle = selected && forced ? 'HighlightText' : textColor;
    context.fillText(String(row.index), left / 2, start + row.size / 2);
    context.fillStyle = gridColor;
    context.fillRect(0, line(start + row.size), left, 1 / ratio);
  }
  context.restore();
  context.fillStyle = gridColor;
  context.fillRect(line(left), 0, 1 / ratio, height);
  context.fillRect(0, line(top), width, 1 / ratio);
  context.fillStyle = textColor;
  context.beginPath();
  context.moveTo(left * 0.42, top * 0.75);
  context.lineTo(left * 0.8, top * 0.75);
  context.lineTo(left * 0.8, top * 0.3);
  context.fill();
}

function lastTrack(tracks: readonly Track[], index: number): Track | undefined {
  for (let offset = tracks.length - 1; offset >= 0; offset--) {
    if (tracks[offset]!.index <= index) return tracks[offset];
  }
}
