import { findCell, type CellBorder, type CellDocument, type CellExtent, type SheetDocument, type WorkbookName } from '../core/workbook';
import { mergeAt } from '../application/border-commands';
import { cellBorders, type BorderSegment } from '../core/formats/borders';
import { drawBorder } from './border-rendering';
import { drawCellFill, drawAlignedText, effectiveCellStyle, usesLightPaper } from './cell-appearance';
import { columnName, gridMetrics, type GridAxis, type Track } from './grid-geometry';
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
  readonly rowAxis: GridAxis;
  readonly columnAxis: GridAxis;
  readonly selection: GridSelection;
  readonly areas: readonly Area[];
  readonly scale: number;
  readonly headings: boolean;
  readonly gridlines: boolean;
  readonly showFormulas: boolean;
  readonly staleCells: ReadonlySet<string>;
  readonly separator: ',' | ';';
  readonly referenceArea?: Area | undefined;
  readonly borderPreview?: { readonly area: Area } | { readonly segments: readonly BorderSegment[]; readonly pen: CellBorder } | undefined;
  readonly resizePreview?: Area | undefined;
  readonly tables?: readonly WorkbookName[] | undefined;
  readonly frozenRows: number;
  readonly frozenColumns: number;
  readonly frozenHeight: number;
  readonly frozenWidth: number;
}

export function drawWorksheet(canvas: HTMLCanvasElement, frame: GridFrame): void {
  const painting = canvas.getContext('2d');
  const { width, height } = frame.view;
  if (!painting || width === 0 || height === 0) return;
  const context = painting;
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
  const { selection, areas, scale, headings } = frame;
  const left = headings ? gridMetrics.rowHeader * scale : 0;
  const top = headings ? gridMetrics.columnHeader * scale : 0;
  const x = (track: Track) => left + track.start;
  const y = (track: Track) => top + track.start;
  const line = (position: number) => Math.round(position * ratio) / ratio - 1 / ratio;
  const font = (size: number, weight = '') => `${weight ? weight + ' ' : ''}${size * scale}px ${styles.fontFamily}`;

  context.fillStyle = sheetColor;
  context.fillRect(0, 0, width, height);
  const splitX = Math.min(width, left + frame.frozenWidth), splitY = Math.min(height, top + frame.frozenHeight);
  const groups = (tracks: readonly Track[], frozen: number, start: number, split: number, end: number) =>
    [{ tracks: tracks.filter(track => track.index <= frozen), start, end: split },
      { tracks: tracks.filter(track => track.index > frozen), start: split, end }].filter(group => group.tracks.length && group.start < group.end);
  const rowGroups = groups(frame.rows, frame.frozenRows, top, splitY, height);
  const columnGroups = groups(frame.columns, frame.frozenColumns, left, splitX, width);

  function clip(x: number, y: number, w: number, h: number) {
    context.save();
    context.beginPath();
    context.rect(x, y, w, h);
    context.clip();
  }

  function pane(rows: readonly Track[], columns: readonly Track[], bounds: { left: number; top: number; right: number; bottom: number }) {
    clip(bounds.left, bounds.top, bounds.right - bounds.left, bounds.bottom - bounds.top);
    drawPane(rows, columns);
    context.restore();
  }

  function extentRect(merge: CellExtent, row: Track, column: Track) {
    const firstRow = frame.rowAxis.track(merge.row) ?? frame.rowAxis.tracks.find(track => track.index >= merge.row)!;
    const firstColumn = frame.columnAxis.track(merge.column) ?? frame.columnAxis.tracks.find(track => track.index >= merge.column)!;
    return { x: x(column) + firstColumn.start - frame.columnAxis.track(column.index)!.start,
      y: y(row) + firstRow.start - frame.rowAxis.track(row.index)!.start,
      width: frame.columnAxis.span(merge.column, merge.column + merge.columns - 1),
      height: frame.rowAxis.span(merge.row, merge.row + merge.rows - 1) };
  }

  function drawPane(rows: readonly Track[], columns: readonly Track[]) {
    context.textBaseline = 'middle';

    const drawn = new Set<CellExtent>();
    const centered = new Set<CellDocument>();
    const entries: { cell: CellDocument; x: number; y: number; width: number; height: number; selected: boolean; active: boolean }[] = [];
    for (const row of rows) {
      for (const column of columns) {
        const point = { row: row.index, column: column.index };
        const merge = mergeAt(frame.sheet, point.row, point.column);
        if (merge && drawn.has(merge)) continue;
        if (merge) drawn.add(merge);
        const target = merge ?? point;
        const cell = findCell(frame.sheet, target);
        const cellStyle = effectiveCellStyle(cell?.style, target, frame.tables ?? []);
        const effectiveCell: CellDocument | undefined = cell ? (cellStyle !== cell.style ? { ...cell, ...(cellStyle ? { style: cellStyle } : {}) } as CellDocument : cell) : undefined;
        const selected = areas.some(area => includes(area, target));
        const active = target.row === selection.anchor.row && target.column === selection.anchor.column;
        const rect = merge ? extentRect(merge, row, column) :
          { x: x(column), y: y(row), width: column.size, height: row.size };
        if (!forced) drawCellFill(context, cellStyle, rect.x, rect.y, rect.width, rect.height, scale);
        if (selected && !active) {
          context.fillStyle = tint;
          context.fillRect(rect.x, rect.y, rect.width, rect.height);
        }
        if (!effectiveCell) continue;
        if (!merge && effectiveCell.style?.alignment === 'centerContinuous') {
          const hasText = (item: CellDocument) => !!presentCell(item, frame.showFormulas, frame.separator).text;
          const across = (index: number): CellDocument | undefined => {
            const candidate = findCell(frame.sheet, { row: row.index, column: index });
            const candidateStyle = effectiveCellStyle(candidate?.style, { row: row.index, column: index }, frame.tables ?? []);
            return candidateStyle?.alignment === 'centerContinuous' && !mergeAt(frame.sheet, row.index, index) ?
              (candidate && candidateStyle !== candidate.style ? { ...candidate, ...(candidateStyle ? { style: candidateStyle } : {}) } as CellDocument : candidate) : undefined;
          };
          let anchor = effectiveCell;
          if (!hasText(anchor)) {
            for (let prior = column.index - 1; prior >= 1; prior--) {
              const candidate = across(prior);
              if (!candidate) break;
              if (hasText(candidate)) { anchor = candidate; break; }
            }
          }
          if (!hasText(anchor) || centered.has(anchor)) continue;
          centered.add(anchor);
          let end = anchor.column;
          for (let next = end + 1; next <= frame.columnAxis.tracks.at(-1)!.index; next++) {
            const candidate = across(next);
            if (!candidate || hasText(candidate)) break;
            end = next;
          }
          const bounds = extentRect({ row: row.index, column: anchor.column, rows: 1, columns: end - anchor.column + 1 }, row, column);
          entries.push({ cell: anchor, ...bounds, selected, active });
        } else entries.push({ cell: effectiveCell, ...rect, selected, active });
      }
    }
    for (const { cell, x: cellX, y: cellY, width, height, selected, active } of entries) {
      const { text: label, color } = presentCell(cell, frame.showFormulas, frame.separator);
      const value = frame.showFormulas ? cell.content : cell.cache ?? cell.content;
      clip(cellX, cellY, width, height);
      const authoredColor = color ?? cell.style?.textRgb;
      context.fillStyle = forced ? selected && !active ? 'HighlightText' : textColor :
        authoredColor ? '#' + authoredColor : usesLightPaper(cell.style) ? '#1b1b1b' : textColor;
      context.save();
      drawAlignedText(context, label, cell.style, styles.fontFamily, scale,
        { x: cellX, y: cellY, width, height }, value.kind);
      context.restore();
      if (frame.staleCells.has(cell.id)) {
        context.strokeStyle = forced ? context.fillStyle : token('--stale-calculation');
        context.lineWidth = 1;
        context.setLineDash([3, 2]);
        context.beginPath();
        context.moveTo(cellX + 2, cellY + height - 3);
        context.lineTo(cellX + width - 2, cellY + height - 3);
        context.stroke();
      }
      context.restore();
    }

    if (frame.gridlines) {
      context.fillStyle = gridColor;
      if (!frame.sheet.merges.length) {
        for (const row of rows) context.fillRect(left, line(y(row) + row.size), width - left, 1 / ratio);
        for (const column of columns) context.fillRect(line(x(column) + column.size), top, 1 / ratio, height - top);
      } else {
        for (const row of rows) for (const column of columns) {
          const merge = mergeAt(frame.sheet, row.index, column.index);
          const nextRow = frame.rowAxis.step(row.index, 1);
          const nextColumn = frame.columnAxis.step(column.index, 1);
          if (!merge || nextRow === row.index || nextRow >= merge.row + merge.rows) {
            context.fillRect(x(column), line(y(row) + row.size), column.size, 1 / ratio);
          }
          if (!merge || nextColumn === column.index || nextColumn >= merge.column + merge.columns) {
            context.fillRect(line(x(column) + column.size), y(row), 1 / ratio, row.size);
          }
        }
      }
    }
    const painted = new Map<string, { pen: CellBorder; x: number; y: number; x2: number; y2: number }>();
    const add = (pen: CellBorder | undefined, x: number, y: number, x2: number, y2: number) => {
      if (pen) painted.set([x, y, x2, y2].join(','), { pen, x, y, x2, y2 });
    };
    drawn.clear();
    for (const row of rows) for (const column of columns) {
      const merge = mergeAt(frame.sheet, row.index, column.index);
      if (merge && drawn.has(merge)) continue;
      if (merge) drawn.add(merge);
      const style = findCell(frame.sheet, merge ?? { row: row.index, column: column.index })?.style;
      if (!style?.borders && !style?.borderDetails) continue;
      const borders = cellBorders(style);
      const rect = merge ? extentRect(merge, row, column) : { x: x(column), y: y(row), width: column.size, height: row.size };
      const a = rect.x, b = rect.y, c = a + rect.width, d = b + rect.height;
      const cellRows = merge ? rows.filter(track => track.index >= merge.row && track.index < merge.row + merge.rows) : [row];
      const cellColumns = merge ? columns.filter(track => track.index >= merge.column && track.index < merge.column + merge.columns) : [column];
      // Shared worksheet edges use the same segment keys as the adjoining cells.
      for (const track of cellColumns) {
        add(borders.top, x(track), b, x(track) + track.size, b);
        add(borders.bottom, x(track), d, x(track) + track.size, d);
      }
      for (const track of cellRows) {
        add(borders.left, a, y(track), a, y(track) + track.size);
        add(borders.right, c, y(track), c, y(track) + track.size);
      }
      add(borders.diagonalUp, a, d, c, b);
      add(borders.diagonalDown, a, b, c, d);
    }
    for (const segment of painted.values()) drawBorder(context, segment.pen, segment.x, segment.y, segment.x2, segment.y2, scale, textColor, forced);

    const outline = (area: Area, width = 1 / ratio) => {
      const firstRow = rows.find(row => row.index >= area.top);
      const lastRow = lastTrack(rows, area.bottom);
      const firstColumn = columns.find(column => column.index >= area.left);
      const lastColumn = lastTrack(columns, area.right);
      if (!(firstRow && lastRow && firstColumn && lastColumn) ||
          firstRow.index > lastRow.index || firstColumn.index > lastColumn.index) return;
      const leftEdge = x(firstColumn), rightEdge = x(lastColumn) + lastColumn.size;
      const topEdge = y(firstRow), bottomEdge = y(lastRow) + lastRow.size;
      context.lineWidth = width;
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
    const activeMerge = mergeAt(frame.sheet, selection.anchor.row, selection.anchor.column);
    const mergeRow = activeMerge && rows.find(track => track.index >= activeMerge.row && track.index < activeMerge.row + activeMerge.rows);
    const mergeColumn = activeMerge && columns.find(track => track.index >= activeMerge.column && track.index < activeMerge.column + activeMerge.columns);
    const activeRect = activeMerge ? mergeRow && mergeColumn && extentRect(activeMerge, mergeRow, mergeColumn) :
      activeRow && activeColumn && { x: x(activeColumn), y: y(activeRow), width: activeColumn.size, height: activeRow.size };
    if (activeRect) {
      context.lineWidth = 2;
      context.strokeRect(activeRect.x + 1, activeRect.y + 1, Math.max(0, activeRect.width - 2), Math.max(0, activeRect.height - 2));
    }
    if (frame.referenceArea) {
      context.strokeStyle = forced ? 'CanvasText' : token('--formula-reference');
      context.setLineDash([4 * scale, 2 * scale]);
      outline(frame.referenceArea);
      context.setLineDash([]);
    }
    if (frame.resizePreview) {
      context.save();
      context.strokeStyle = accent;
      context.setLineDash([6 * scale, 4 * scale]);
      outline(frame.resizePreview, 2);
      context.restore();
    }
    const preview = frame.borderPreview;
    if (preview) {
      context.strokeStyle = accent;
      context.setLineDash([4 * scale, 2 * scale]);
      if ('area' in preview) outline(preview.area);
      else for (const segment of preview.segments) {
        const row = rows.find(track => track.index === segment.row);
        const column = columns.find(track => track.index === segment.column);
        if (segment.edge === 'top' && column) {
          const previous = rows.find(track => track.index === segment.row - 1);
          const lineY = row ? y(row) : previous && y(previous) + previous.size;
          if (lineY !== undefined) drawBorder(context, preview.pen, x(column), lineY, x(column) + column.size, lineY, scale, accent, forced);
        }
        if (segment.edge === 'left' && row) {
          const previous = columns.find(track => track.index === segment.column - 1);
          const lineX = column ? x(column) : previous && x(previous) + previous.size;
          if (lineX !== undefined) drawBorder(context, preview.pen, lineX, y(row), lineX, y(row) + row.size, scale, accent, forced);
        }
      }
    }
  }

  for (const rowGroup of rowGroups) for (const columnGroup of columnGroups) pane(rowGroup.tracks, columnGroup.tracks,
    { left: columnGroup.start, top: rowGroup.start, right: columnGroup.end, bottom: rowGroup.end });

  if (!headings) return;
  context.fillStyle = headerColor;
  context.fillRect(0, 0, width, top);
  context.fillRect(0, top, left, height - top);
  context.font = font(12);
  context.textAlign = 'center';
  context.textBaseline = 'middle';
  for (const group of columnGroups) {
    clip(group.start, 0, group.end - group.start, top);
    for (const column of group.tracks) {
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
  }
  for (const group of rowGroups) {
    clip(0, group.start, left, group.end - group.start);
    for (const row of group.tracks) {
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
  }
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
