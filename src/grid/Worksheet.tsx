import {
  useId, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState,
  type KeyboardEvent, type PointerEvent, type Ref,
} from 'react';
import { findCell, worksheetViewport, type Coordinate, type SheetDocument } from '../core/workbook';
import { drawWorksheet, type GridViewport } from './draw-worksheet';
import { address, clamp, GridAxis, gridMetrics } from './grid-geometry';
import {
  cellDisplay, includes, navigate, select, selectionAreas, singleSelection,
  type AreaKind, type GridSelection,
} from './grid-selection';

export interface WorksheetHandle {
  focus(): void;
  goTo(point: Coordinate): boolean;
  zoomToSelection(): number;
}

interface WorksheetProps {
  readonly sheet: SheetDocument;
  readonly selection: GridSelection;
  readonly onSelection: (selection: GridSelection) => void;
  readonly onAdjacentSheet: (direction: number) => void;
  readonly zoom: number;
  readonly headings: boolean;
  readonly gridlines: boolean;
  readonly ref: Ref<WorksheetHandle>;
}

interface Drag {
  readonly id: number;
  readonly kind: AreaKind;
  selection: GridSelection;
  x: number;
  y: number;
}

const initialViewport: GridViewport = { x: 0, y: 0, width: 0, height: 0 };

export function Worksheet(props: WorksheetProps) {
  const { sheet, selection, onSelection, onAdjacentSheet, zoom, headings, gridlines } = props;
  const scroll = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const drag = useRef<Drag | null>(null);
  const exitOnTab = useRef(false);
  const fitStart = useRef<Coordinate | null>(null);
  const frame = useRef(0);
  const id = useId();
  const [view, setView] = useState(initialViewport);
  const [palette, setPalette] = useState(0);
  const scale = zoom / 100;
  const left = headings ? gridMetrics.rowHeader * scale : 0;
  const top = headings ? gridMetrics.columnHeader * scale : 0;
  const rows = useMemo(() => new GridAxis(worksheetViewport.rows, gridMetrics.row, sheet.rows, scale), [sheet.rows, scale]);
  const columns = useMemo(() => new GridAxis(worksheetViewport.columns, gridMetrics.column, sheet.columns, scale), [sheet.columns, scale]);
  const rowTracks = rows.visible(view.y, view.height - top);
  const columnTracks = columns.visible(view.x, view.width - left);
  const areas = selectionAreas(selection);
  const cellId = (point: Coordinate) => id + '-' + address(point);

  function measure() {
    const node = scroll.current;
    if (!node) return;
    const next = { x: node.scrollLeft, y: node.scrollTop, width: node.clientWidth, height: node.clientHeight };
    setView(current => current.x === next.x && current.y === next.y &&
      current.width === next.width && current.height === next.height ? current : next);
  }

  function reveal(point: Coordinate) {
    const node = scroll.current;
    if (!node) return;
    node.scrollLeft = columns.reveal(point.column, node.scrollLeft, node.clientWidth - left);
    node.scrollTop = rows.reveal(point.row, node.scrollTop, node.clientHeight - top);
    measure();
  }

  function focus() {
    scroll.current?.focus({ preventScroll: true });
  }

  useImperativeHandle(props.ref, () => ({
    focus,
    goTo(point) {
      if (!rows.track(point.row) || !columns.track(point.column)) return false;
      onSelection(singleSelection(point));
      reveal(point);
      focus();
      return true;
    },
    zoomToSelection() {
      const minRow = Math.min(...areas.map(area => area.top));
      const maxRow = Math.max(...areas.map(area => area.bottom));
      const minColumn = Math.min(...areas.map(area => area.left));
      const maxColumn = Math.max(...areas.map(area => area.right));
      const widthAt100 = (columns.span(minColumn, maxColumn) + left) / scale;
      const heightAt100 = (rows.span(minRow, maxRow) + top) / scale;
      const next = clamp(Math.floor(100 * Math.min(view.width / widthAt100, view.height / heightAt100)), 10, 400);
      if (next === zoom) alignTo({ row: minRow, column: minColumn });
      else fitStart.current = { row: minRow, column: minColumn };
      return next;
    },
  }));

  useLayoutEffect(() => {
    const node = scroll.current!;
    const observer = new ResizeObserver(measure);
    observer.observe(node);
    const media = matchMedia('(forced-colors: active)');
    const invalidateColors = () => setPalette(version => version + 1);
    const resized = () => { measure(); invalidateColors(); };
    window.addEventListener('resize', resized);
    media.addEventListener('change', invalidateColors);
    measure();
    return () => {
      observer.disconnect();
      window.removeEventListener('resize', resized);
      media.removeEventListener('change', invalidateColors);
      cancelAnimationFrame(frame.current);
    };
  }, []);

  useLayoutEffect(() => {
    if (fitStart.current) {
      alignTo(fitStart.current);
      fitStart.current = null;
    } else reveal(selection.anchor);
  }, [rows, columns, headings]); // Remap the active coordinate when the worksheet geometry changes.

  useLayoutEffect(() => {
    drawWorksheet(canvas.current!, {
      sheet, view, rows: rowTracks, columns: columnTracks, selection, areas, scale, headings, gridlines,
    });
  }, [sheet, view, rowTracks, columnTracks, selection, areas, scale, headings, gridlines, palette]);

  function alignTo(point: Coordinate) {
    const node = scroll.current!;
    const nextRow = rows.tracks.find(track => track.index >= point.row);
    const nextColumn = columns.tracks.find(track => track.index >= point.column);
    if (nextRow) node.scrollTop = nextRow.start;
    if (nextColumn) node.scrollLeft = nextColumn.start;
    measure();
  }

  function pointAt(x: number, y: number, kind: AreaKind, bounded = false): Coordinate | undefined {
    const node = scroll.current!;
    const locate = (axis: GridAxis, position: number, header: number, offset: number, size: number) => {
      const relative = (bounded ? clamp(position, header, size - 0.01) : position) - header + offset;
      return axis.at(bounded ? clamp(relative, 0, axis.total - 0.01) : relative)?.index;
    };
    const row = kind === 'columns' || kind === 'sheet' ? rows.tracks[0]?.index :
      locate(rows, y, top, node.scrollTop, node.clientHeight);
    const column = kind === 'rows' || kind === 'sheet' ? columns.tracks[0]?.index :
      locate(columns, x, left, node.scrollLeft, node.clientWidth);
    return row !== undefined && column !== undefined ? { row, column } : undefined;
  }

  function moveDrag() {
    const active = drag.current;
    if (!active) return;
    const node = scroll.current!;
    const { left: originX, top: originY } = node.getBoundingClientRect();
    const x = active.x - originX;
    const y = active.y - originY;
    const xOverflow = active.kind === 'rows' ? 0 : x - clamp(x, left, node.clientWidth);
    const yOverflow = active.kind === 'columns' ? 0 : y - clamp(y, top, node.clientHeight);
    node.scrollLeft += clamp(xOverflow, -35, 35);
    node.scrollTop += clamp(yOverflow, -30, 30);
    const point = pointAt(x, y, active.kind, true);
    if (point && (point.row !== active.selection.end.row || point.column !== active.selection.end.column)) {
      active.selection = select(active.selection, point, active.kind, true);
      onSelection(active.selection);
    }
    if ((xOverflow !== 0 || yOverflow !== 0) && drag.current) frame.current = requestAnimationFrame(moveDrag);
  }

  function pointerDown(event: PointerEvent<HTMLDivElement>) {
    if (event.button !== 0 || !event.isPrimary) return;
    const node = event.currentTarget;
    const bounds = node.getBoundingClientRect();
    const x = event.clientX - bounds.left;
    const y = event.clientY - bounds.top;
    if (x >= node.clientWidth || y >= node.clientHeight) return;
    const kind: AreaKind = y < top ? x < left ? 'sheet' : 'columns' : x < left ? 'rows' : 'cells';
    const point = pointAt(x, y, kind);
    if (!point) return;
    focus();
    const additive = event.ctrlKey || event.metaKey;
    const hitSelected = additive && areas.some(area => includes(area, point));
    const next = select(selection, point, kind, event.shiftKey, additive);
    onSelection(next);
    if (event.pointerType === 'touch' || kind === 'sheet' || hitSelected) return;
    event.preventDefault();
    node.setPointerCapture(event.pointerId);
    drag.current = { id: event.pointerId, kind, selection: next, x: event.clientX, y: event.clientY };
  }

  function pointerMove(event: PointerEvent<HTMLDivElement>) {
    const active = drag.current;
    if (!active || active.id !== event.pointerId) return;
    active.x = event.clientX;
    active.y = event.clientY;
    cancelAnimationFrame(frame.current);
    moveDrag();
  }

  function endDrag(event: PointerEvent<HTMLDivElement>) {
    if (drag.current?.id !== event.pointerId) return;
    drag.current = null;
    cancelAnimationFrame(frame.current);
  }

  function keyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (event.nativeEvent.isComposing || event.altKey) return;
    const control = event.ctrlKey || event.metaKey;
    const { key, shiftKey } = event;
    const anchor = selection.anchor;
    if (key === 'Tab' && (exitOnTab.current || control)) {
      exitOnTab.current = false;
      return;
    }
    if (key !== 'Escape') exitOnTab.current = false;
    if (control && (key === 'PageUp' || key === 'PageDown')) {
      event.preventDefault();
      onAdjacentSheet(key === 'PageUp' ? -1 : 1);
      return;
    }
    if (control && key.toLowerCase() === 'a' || key === ' ' && (control || shiftKey)) {
      event.preventDefault();
      const kind = key !== ' ' || control && shiftKey ? 'sheet' : control ? 'columns' : 'rows';
      const point = key.toLowerCase() === 'a' ? {
        row: rows.tracks[0]?.index ?? anchor.row, column: columns.tracks[0]?.index ?? anchor.column,
      } : anchor;
      onSelection(select(selection, point, kind));
      if (point !== anchor) reveal(point);
      return;
    }
    if (key === 'Escape') {
      event.preventDefault();
      exitOnTab.current = true;
      onSelection(singleSelection(anchor));
      reveal(anchor);
      return;
    }
    if (control && key === 'Enter') return;
    const extend = shiftKey && key !== 'Enter' && key !== 'Tab';
    const point = navigate(sheet, rows, columns, extend ? selection.end : anchor, key, control, shiftKey,
      Math.max(1, rowTracks.length - 1));
    if (!point) return;
    event.preventDefault();
    onSelection(select(selection, point, extend ? selection.kind : 'cells', extend));
    reveal(point);
  }

  // Keep the active descendant mounted even when native scrolling moves it outside the viewport.
  const projected = new Map(rowTracks.slice(0, 64).map(row => [row.index, columnTracks.map(column => column.index)]));
  const activeColumns = projected.get(selection.anchor.row) ?? [];
  if (!activeColumns.includes(selection.anchor.column)) {
    projected.set(selection.anchor.row, [...activeColumns, selection.anchor.column].sort((a, b) => a - b));
  }

  return (
    <div className="worksheet">
      <div ref={scroll} className="worksheet-scroll" role="grid" aria-label={sheet.name}
        aria-rowcount={worksheetViewport.rows} aria-colcount={worksheetViewport.columns}
        aria-readonly="true" aria-multiselectable="true" aria-activedescendant={cellId(selection.anchor)} tabIndex={0}
        aria-describedby={id + '-help'} onBlur={() => { exitOnTab.current = false; }}
        onScroll={measure} onKeyDown={keyDown} onPointerDown={pointerDown} onPointerMove={pointerMove}
        onPointerUp={endDrag} onPointerCancel={endDrag} onLostPointerCapture={endDrag}>
        <div aria-hidden="true" className="worksheet-spacer" style={{ width: left + columns.total, height: top + rows.total }} />
        <div className="worksheet-projection">
          {[...projected].sort(([a], [b]) => a - b).filter(([, indexes]) => indexes.length > 0).map(([row, indexes]) => (
            <div key={row} role="row" aria-rowindex={row}>
              {indexes.map(column => {
                const point = { row, column };
                const value = cellDisplay(findCell(sheet, point));
                return <div key={column} id={cellId(point)} role="gridcell" aria-colindex={column}
                  aria-selected={areas.some(area => includes(area, point))}
                  aria-label={address(point) + (value ? ', ' + value : ', blank')} />;
              })}
            </div>
          ))}
        </div>
      </div>
      <p className="worksheet-projection" id={id + '-help'}>
        Use arrow keys or Tab to move between cells. Press Escape then Tab to leave the worksheet.
      </p>
      <canvas ref={canvas} aria-hidden="true" className="worksheet-canvas" style={{ width: view.width, height: view.height }} />
    </div>
  );
}
