import { findCell, findColumnStyle, findRowStyle, type CellBorder, type CellDocument, type CellExtent, type CellStyle, type SheetDocument, type WorkbookName } from '../core/workbook';
import { mergeAt } from '../application/border-commands';
import { cellBorders, type BorderSegment } from '../core/formats/borders';
import { displayCellValue } from '../core/formats/cell-format';
import { drawBorder } from './border-rendering';
import { cellFont, drawCellFill, drawAlignedText, effectiveCellStyle, usesLightPaper } from './cell-appearance';
import { columnName, gridMetrics, type GridAxis, type Track } from './grid-geometry';
import { formulaSpillArea, presentCell, includes, type Area, type GridSelection } from './grid-selection';

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
  readonly calculatingCells: ReadonlySet<string>;
  readonly separator: ',' | ';';
  readonly referenceArea?: Area | undefined;
  readonly formulaHighlights?: readonly { readonly area: Area; readonly color: string }[] | undefined;
  readonly borderPreview?: { readonly area: Area } | { readonly segments: readonly BorderSegment[]; readonly pen: CellBorder } | undefined;
  readonly resizePreview?: Area | undefined;
  readonly tables?: readonly WorkbookName[] | undefined;
  readonly frozenRows: number;
  readonly frozenColumns: number;
  readonly frozenHeight: number;
  readonly frozenWidth: number;
  /** CSS pixel width of the row-number column, already scaled. */
  readonly rowHeaderWidth: number;
  /** Incremented only when theme/forced-colors palette changes. */
  readonly themeVersion?: number;
}

interface CachedTheme {
  readonly version: number;
  readonly forced: boolean;
  readonly sheetColor: string;
  readonly headerColor: string;
  readonly gridColor: string;
  readonly textColor: string;
  readonly accent: string;
  readonly tint: string;
  readonly staleColor: string;
  readonly freezeColor: string;
  readonly referenceColor: string;
  readonly fontFamily: string;
}

const themeCache = new WeakMap<HTMLCanvasElement, CachedTheme>();

function resolveTheme(canvas: HTMLCanvasElement, version: number | undefined): CachedTheme {
  if (version !== undefined) {
    const cached = themeCache.get(canvas);
    if (cached && cached.version === version) return cached;
  }
  const styles = getComputedStyle(canvas);
  const token = (key: string) => styles.getPropertyValue(key).trim();
  const forced = matchMedia('(forced-colors: active)').matches;
  const theme: CachedTheme = {
    version: version ?? -1,
    forced,
    sheetColor: forced ? 'Canvas' : token('--sheet'),
    headerColor: forced ? 'Canvas' : token('--grid-header'),
    gridColor: forced ? 'GrayText' : token('--gridline'),
    textColor: forced ? 'CanvasText' : token('--text'),
    accent: forced ? 'Highlight' : token('--accent'),
    tint: forced ? 'Highlight' : token('--grid-selection'),
    staleColor: forced ? 'CanvasText' : token('--stale-calculation'),
    freezeColor: forced ? 'CanvasText' : token('--freeze-line'),
    referenceColor: forced ? 'CanvasText' : token('--formula-reference'),
    fontFamily: styles.fontFamily,
  };
  if (version !== undefined) themeCache.set(canvas, theme);
  return theme;
}

export interface HeaderCornerPoint { readonly x: number; readonly y: number }

/** Returns the centered 45° select-all marker as an isosceles right triangle. */
export function selectAllCornerTriangle(width: number, height: number): readonly [HeaderCornerPoint, HeaderCornerPoint, HeaderCornerPoint] {
  const size = Math.min(width, height) * 0.64;
  const right = width * 0.8;
  const bottom = height * 0.75;
  return [
    { x: right - size, y: bottom },
    { x: right, y: bottom },
    { x: right, y: bottom - size },
  ];
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

  const theme = resolveTheme(canvas, frame.themeVersion);
  const { forced, sheetColor, headerColor, gridColor, textColor, accent, tint, staleColor, freezeColor, referenceColor, fontFamily } = theme;
  const { selection, areas, scale, headings } = frame;
  const left = headings ? frame.rowHeaderWidth : 0;
  const top = headings ? gridMetrics.columnHeader * scale : 0;
  const x = (track: Track) => left + track.start;
  const y = (track: Track) => top + track.start;
  const line = (position: number) => Math.round(position * ratio) / ratio - 1 / ratio;
  const font = (size: number, weight = '') => `${weight ? weight + ' ' : ''}${size * scale}px ${fontFamily}`;

  context.fillStyle = sheetColor;
  context.fillRect(0, 0, width, height);
  const splitX = Math.min(width, left + frame.frozenWidth), splitY = Math.min(height, top + frame.frozenHeight);
  const groups = (tracks: readonly Track[], frozen: number, start: number, split: number, end: number) =>
    [{ tracks: tracks.filter(track => track.index <= frozen), start, end: split },
      { tracks: tracks.filter(track => track.index > frozen), start: split, end }].filter(group => group.tracks.length && group.start < group.end);
  const rowGroups = groups(frame.rows, frame.frozenRows, top, splitY, height);
  const columnGroups = groups(frame.columns, frame.frozenColumns, left, splitX, width);
  const firstVisibleRow = frame.rows[0]?.index ?? 0;
  const lastVisibleRow = frame.rows.at(-1)?.index ?? 0;
  const firstVisibleCol = frame.columns[0]?.index ?? 0;
  const lastVisibleCol = frame.columns.at(-1)?.index ?? 0;
  const visibleTables = frame.tables?.length
    ? frame.tables.filter(t => !t.area.invalid &&
      t.area.row <= lastVisibleRow && t.area.row + t.area.rows - 1 >= firstVisibleRow &&
      t.area.column <= lastVisibleCol && t.area.column + t.area.columns - 1 >= firstVisibleCol)
    : [];

  function clip(target: CanvasRenderingContext2D, x: number, y: number, w: number, h: number) {
    target.save();
    target.beginPath();
    target.rect(x, y, w, h);
    target.clip();
  }

  function extentRect(merge: CellExtent, row: Track, column: Track) {
    const firstRow = frame.rowAxis.trackAtOrAfter(merge.row)!;
    const firstColumn = frame.columnAxis.trackAtOrAfter(merge.column)!;
    return { x: x(column) + firstColumn.start - frame.columnAxis.track(column.index)!.start,
      y: y(row) + firstRow.start - frame.rowAxis.track(row.index)!.start,
      width: frame.columnAxis.span(merge.column, merge.column + merge.columns - 1),
      height: frame.rowAxis.span(merge.row, merge.row + merge.rows - 1) };
  }

  function drawPaneBody(target: CanvasRenderingContext2D, rows: readonly Track[], columns: readonly Track[], includeSelectionTint: boolean) {
    target.textBaseline = 'middle';

    const drawn = new Set<CellExtent>();
    const centered = new Set<CellDocument>();
    const entries: { cell: CellDocument; style: CellStyle | undefined; numberFormat: string | undefined; x: number; y: number; width: number; height: number; selected: boolean; active: boolean }[] = [];
    const bordered: { style: CellStyle; merge: CellExtent | undefined; row: Track; column: Track; rect: { x: number; y: number; width: number; height: number } }[] = [];
    const singleArea = areas.length === 1 ? areas[0] : undefined;
    for (const row of rows) {
      const rowStyle = findRowStyle(frame.sheet, row.index);
      for (const column of columns) {
        const point = { row: row.index, column: column.index };
        const merge = mergeAt(frame.sheet, point.row, point.column);
        if (merge && drawn.has(merge)) continue;
        if (merge) drawn.add(merge);
        const cellPoint = merge ?? point;
        const cell = findCell(frame.sheet, cellPoint);
        const authoredStyle = cell?.style !== undefined ? cell.style :
          (merge ? findRowStyle(frame.sheet, cellPoint.row) : rowStyle) ?? findColumnStyle(frame.sheet, cellPoint.column);
        const cellStyle = effectiveCellStyle(authoredStyle, cellPoint, visibleTables);
        const selected = includeSelectionTint ? (singleArea ? includes(singleArea, cellPoint) : areas.some(area => includes(area, cellPoint))) : false;
        const active = includeSelectionTint ? (cellPoint.row === selection.anchor.row && cellPoint.column === selection.anchor.column) : false;
        const rect = merge ? extentRect(merge, row, column) :
          { x: x(column), y: y(row), width: column.size, height: row.size };
        if (authoredStyle && (authoredStyle.borders || authoredStyle.borderDetails)) {
          bordered.push({ style: authoredStyle, merge, row, column, rect });
        }
        if (!forced) drawCellFill(target, cellStyle, rect.x, rect.y, rect.width, rect.height, scale);
        if (selected && !active) {
          target.fillStyle = tint;
          target.fillRect(rect.x, rect.y, rect.width, rect.height);
        }
        if (!cell) continue;
        if (!merge && cellStyle?.alignment === 'centerContinuous') {
          const hasText = (item: CellDocument, itemStyle: CellStyle | undefined) =>
            !!presentCell(item, frame.showFormulas, frame.separator, itemStyle?.numberFormat).text;
          const across = (index: number): { cell: CellDocument; style: CellStyle | undefined; rawStyle: CellStyle | undefined } | undefined => {
            const candidate = findCell(frame.sheet, { row: row.index, column: index });
            const rawStyle = candidate?.style !== undefined ? candidate.style : rowStyle ?? findColumnStyle(frame.sheet, index);
            const candidateStyle = effectiveCellStyle(rawStyle, { row: row.index, column: index }, visibleTables);
            return candidate && candidateStyle?.alignment === 'centerContinuous' && !mergeAt(frame.sheet, row.index, index) ?
              { cell: candidate, style: candidateStyle, rawStyle } : undefined;
          };
          let anchor = cell;
          let anchorStyle: CellStyle | undefined = cellStyle;
          let anchorRawStyle: CellStyle | undefined = authoredStyle;
          if (!hasText(anchor, anchorRawStyle)) {
            for (let prior = column.index - 1; prior >= 1; prior--) {
              const candidate = across(prior);
              if (!candidate) break;
              if (hasText(candidate.cell, candidate.rawStyle)) {
                anchor = candidate.cell; anchorStyle = candidate.style; anchorRawStyle = candidate.rawStyle; break;
              }
            }
          }
          if (!hasText(anchor, anchorRawStyle) || centered.has(anchor)) continue;
          centered.add(anchor);
          let end = anchor.column;
          const lastIndex = frame.columnAxis.last?.index ?? end;
          for (let next = end + 1; next <= lastIndex; next++) {
            const candidate = across(next);
            if (!candidate || hasText(candidate.cell, candidate.rawStyle)) break;
            end = next;
          }
          const bounds = extentRect({ row: row.index, column: anchor.column, rows: 1, columns: end - anchor.column + 1 }, row, column);
          entries.push({ cell: anchor, style: anchorStyle, numberFormat: anchorRawStyle?.numberFormat, ...bounds, selected, active });
        } else entries.push({ cell, style: cellStyle, numberFormat: authoredStyle?.numberFormat, ...rect, selected, active });
      }
    }
    const defaultFont = cellFont(undefined, fontFamily, scale).css;
    target.font = defaultFont;
    target.direction = 'ltr';
    target.textBaseline = 'middle';
    for (const { cell, style, numberFormat, x: cellX, y: cellY, width, height, selected, active } of entries) {
      const { text: label, color } = presentCell(cell, frame.showFormulas, frame.separator, numberFormat);
      const value = displayCellValue(cell, frame.showFormulas);
      clip(target, cellX, cellY, width, height);
      const authoredColor = color ?? style?.textRgb;
      target.fillStyle = forced ? selected && !active ? 'HighlightText' : textColor :
        authoredColor ? '#' + authoredColor : usesLightPaper(style) ? '#1b1b1b' : textColor;
      if (style?.textRotation) {
        target.save();
        drawAlignedText(target, label, style, fontFamily, scale,
          { x: cellX, y: cellY, width, height }, value?.kind ?? 'blank', defaultFont, 'ltr');
        target.restore();
      } else {
        drawAlignedText(target, label, style, fontFamily, scale,
          { x: cellX, y: cellY, width, height }, value?.kind ?? 'blank', defaultFont, 'ltr');
      }
      if (frame.calculatingCells.has(cell.id)) {
        target.strokeStyle = accent;
        target.lineWidth = 1.6;
        target.setLineDash([]);
        target.beginPath();
        target.arc(cellX + width - 7, cellY + 7, 3, -Math.PI * 0.42, Math.PI * 1.28);
        target.stroke();
      }
      if (frame.staleCells.has(cell.id)) {
        target.strokeStyle = forced ? target.fillStyle : staleColor;
        target.lineWidth = 1;
        target.setLineDash([3, 2]);
        target.beginPath();
        target.moveTo(cellX + 2, cellY + height - 3);
        target.lineTo(cellX + width - 2, cellY + height - 3);
        target.stroke();
      }
      target.restore();
    }

    if (frame.gridlines) {
      target.fillStyle = gridColor;
      if (!frame.sheet.merges.length) {
        for (const row of rows) target.fillRect(left, line(y(row) + row.size), width - left, 1 / ratio);
        for (const column of columns) target.fillRect(line(x(column) + column.size), top, 1 / ratio, height - top);
      } else {
        for (const row of rows) for (const column of columns) {
          const merge = mergeAt(frame.sheet, row.index, column.index);
          const nextRow = frame.rowAxis.step(row.index, 1);
          const nextColumn = frame.columnAxis.step(column.index, 1);
          if (!merge || nextRow === row.index || nextRow >= merge.row + merge.rows) {
            target.fillRect(x(column), line(y(row) + row.size), column.size, 1 / ratio);
          }
          if (!merge || nextColumn === column.index || nextColumn >= merge.column + merge.columns) {
            target.fillRect(line(x(column) + column.size), y(row), 1 / ratio, row.size);
          }
        }
      }
    }
    if (bordered.length) {
      const painted = new Map<string, { pen: CellBorder; x: number; y: number; x2: number; y2: number }>();
      const add = (pen: CellBorder | undefined, x: number, y: number, x2: number, y2: number) => {
        if (pen) painted.set([x, y, x2, y2].join(','), { pen, x, y, x2, y2 });
      };
      for (const { style, merge, row, column, rect } of bordered) {
        const borders = cellBorders(style);
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
      for (const segment of painted.values()) drawBorder(target, segment.pen, segment.x, segment.y, segment.x2, segment.y2, scale, textColor, forced);
    }
  }

  function drawPaneOverlay(rows: readonly Track[], columns: readonly Track[]) {

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
    const spillArea = formulaSpillArea(frame.sheet, selection.anchor);
    if (spillArea) {
      context.save();
      context.strokeStyle = forced ? 'Highlight' : accent;
      context.globalAlpha = forced ? 1 : 0.65;
      outline(spillArea, 1.5);
      context.restore();
    }
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
    if (frame.formulaHighlights?.length) {
      for (const highlight of frame.formulaHighlights) {
        const { area, color } = highlight;
        const firstRow = rows.find(row => row.index >= area.top);
        const lastRow = lastTrack(rows, area.bottom);
        const firstColumn = columns.find(column => column.index >= area.left);
        const lastColumn = lastTrack(columns, area.right);
        if (!(firstRow && lastRow && firstColumn && lastColumn) ||
            firstRow.index > lastRow.index || firstColumn.index > lastColumn.index) continue;
        const leftEdge = x(firstColumn), rightEdge = x(lastColumn) + lastColumn.size;
        const topEdge = y(firstRow), bottomEdge = y(lastRow) + lastRow.size;
        const paint = forced ? 'Highlight' : color;
        context.save();
        context.fillStyle = paint;
        context.globalAlpha = forced ? 0.15 : 0.14;
        context.fillRect(leftEdge, topEdge, rightEdge - leftEdge, bottomEdge - topEdge);
        context.globalAlpha = 1;
        context.strokeStyle = paint;
        outline(area, 2);
        const handleSize = 4;
        const half = handleSize / 2;
        if (firstRow.index === area.top && firstColumn.index === area.left) {
          context.fillRect(leftEdge - half, topEdge - half, handleSize, handleSize);
        }
        if (firstRow.index === area.top && lastColumn.index === area.right) {
          context.fillRect(rightEdge - half, topEdge - half, handleSize, handleSize);
        }
        if (lastRow.index === area.bottom && firstColumn.index === area.left) {
          context.fillRect(leftEdge - half, bottomEdge - half, handleSize, handleSize);
        }
        if (lastRow.index === area.bottom && lastColumn.index === area.right) {
          context.fillRect(rightEdge - half, bottomEdge - half, handleSize, handleSize);
        }
        context.restore();
      }
    }
    if (frame.referenceArea) {
      context.strokeStyle = referenceColor;
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

  for (const rowGroup of rowGroups) for (const columnGroup of columnGroups) {
    clip(context, columnGroup.start, rowGroup.start, columnGroup.end - columnGroup.start, rowGroup.end - rowGroup.start);
    drawPaneBody(context, rowGroup.tracks, columnGroup.tracks, true);
    drawPaneOverlay(rowGroup.tracks, columnGroup.tracks);
    context.restore();
  }

  // Desktop DrawFrozenDividers paints a 2px divider over content and headings at each frozen edge.
  function drawFrozenDividers() {
    const thickness = Math.max(1, 2 * scale);
    context.fillStyle = freezeColor;
    if (frame.frozenRows > 0 && splitY > top) context.fillRect(0, splitY - thickness, width, thickness);
    if (frame.frozenColumns > 0 && splitX > left) context.fillRect(splitX - thickness, 0, thickness, height);
  }

  if (!headings) {
    drawFrozenDividers();
    return;
  }
  context.fillStyle = headerColor;
  context.fillRect(0, 0, width, top);
  context.fillRect(0, top, left, height - top);
  context.font = font(12);
  context.textAlign = 'center';
  context.textBaseline = 'middle';
  for (const group of columnGroups) {
    clip(context, group.start, 0, group.end - group.start, top);
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
    clip(context, 0, group.start, left, group.end - group.start);
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
  const [bottomLeft, bottomRight, topRight] = selectAllCornerTriangle(left, top);
  context.beginPath();
  context.moveTo(bottomLeft.x, bottomLeft.y);
  context.lineTo(bottomRight.x, bottomRight.y);
  context.lineTo(topRight.x, topRight.y);
  context.fill();
  drawFrozenDividers();
}

function lastTrack(tracks: readonly Track[], index: number): Track | undefined {
  for (let offset = tracks.length - 1; offset >= 0; offset--) {
    if (tracks[offset]!.index <= index) return tracks[offset];
  }
}
