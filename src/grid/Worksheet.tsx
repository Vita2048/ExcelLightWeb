import {
  useId, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState,
  type ClipboardEvent, type KeyboardEvent, type PointerEvent, type ReactNode, type Ref,
} from 'react';
import type { BorderMode, BorderSegment } from '../core/formats/borders';
import { mergeAt } from '../application/border-commands';
import { drawGesture, hitBorder, type BorderHit, type DrawGesture } from './border-geometry';
import { findCell, worksheetViewport, type CellBorder, type Coordinate, type SheetDocument } from '../core/workbook';
import type { Axis } from '../core/axis-edit';
import type { AfterEditKey } from '../features/editing/useWorksheetEditing';
import { drawWorksheet, type GridViewport } from './draw-worksheet';
import drawBorderCursor from './cursors/draw-border.svg';
import eraseBorderCursor from './cursors/erase-border.svg';
import selectColumnCursor from './cursors/select-column.svg';
import selectRowCursor from './cursors/select-row.svg';
import { address, clamp, GridAxis, gridMetrics } from './grid-geometry';
import {
  cellDisplay, includes, mergeSelection, navigate, primaryArea, select, selectionAreas, singleSelection,
  type Area, type AreaKind, type GridSelection,
} from './grid-selection';

export interface WorksheetHandle {
  focus(): void;
  goTo(point: Coordinate): boolean;
  zoomToSelection(): number;
  moveAfterEdit(point: Coordinate, key: AfterEditKey, reverse: boolean): void;
}

interface WorksheetProps {
  readonly sheet: SheetDocument;
  readonly selection: GridSelection;
  readonly onSelection: (selection: GridSelection) => boolean;
  readonly borderTool: BorderMode | undefined;
  readonly borderPen: CellBorder | null;
  readonly onBeginBorder: () => ((gesture: DrawGesture) => void) | undefined;
  readonly onExitBorder: () => void;
  readonly paintActive: boolean;
  readonly onPaint: (area: Area) => void;
  readonly onReferencePick: () => ReferenceDrag | undefined;
  readonly referenceArea?: Area | undefined;
  readonly onAdjacentSheet: (direction: number) => void;
  readonly onEdit: (replacement?: string) => void;
  readonly onShortcut: (event: KeyboardEvent<HTMLDivElement>) => boolean;
  readonly onClipboard: (event: ClipboardEvent<HTMLDivElement>, kind: 'copy' | 'cut' | 'paste') => void;
  readonly editor?: { readonly at: Coordinate; readonly input: ReactNode } | undefined;
  readonly onHeaderResize: (axis: Axis, index: number) => ((size: number) => void) | undefined;
  readonly onHeaderFit: (axis: Axis, index: number) => void;
  readonly onStructureMenu: (x: number, y: number, axis?: Axis, index?: number, point?: Coordinate) => void;
  readonly zoom: number;
  readonly headings: boolean;
  readonly gridlines: boolean;
  readonly showFormulas: boolean;
  readonly staleCells: ReadonlySet<string>;
  readonly separator: ',' | ';';
  readonly ref: Ref<WorksheetHandle>;
}

interface Drag {
  readonly id: number;
  readonly kind: AreaKind;
  selection: GridSelection;
  x: number;
  y: number;
  readonly reference?: ReferenceDrag;
  readonly paint?: boolean;
}

interface BorderDrag {
  readonly id: number;
  readonly mode: BorderMode;
  readonly start: BorderHit;
  readonly apply: (gesture: DrawGesture) => void;
  readonly erased: Map<string, BorderSegment>;
  x: number;
  y: number;
  preview?: DrawGesture;
}

interface ReferenceDrag {
  readonly update: (area: Area) => void;
  readonly cancel: () => void;
}

const initialViewport: GridViewport = { x: 0, y: 0, width: 0, height: 0 };

export function Worksheet(props: WorksheetProps) {
  const { sheet, selection, onSelection, onAdjacentSheet, zoom, headings, gridlines, showFormulas, staleCells, separator } = props;
  const scroll = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const drag = useRef<Drag | null>(null);
  const borderDrag = useRef<BorderDrag | null>(null);
  const [borderPreview, setBorderPreview] = useState<DrawGesture>();
  const resize = useRef<{ id: number; axis: Axis; origin: number; size: number; apply: (size: number) => void } | null>(null);
  const [guide, setGuide] = useState<{ axis: Axis; offset: number }>();
  const exitOnTab = useRef(false);
  const fitStart = useRef<Coordinate | null>(null);
  const frame = useRef(0);
  const id = useId();
  const [view, setView] = useState(initialViewport);
  const [palette, setPalette] = useState(0);
  const scale = zoom / 100;
  const left = headings ? gridMetrics.rowHeader * scale : 0;
  const top = headings ? gridMetrics.columnHeader * scale : 0;
  const rows = useMemo(() => new GridAxis(worksheetViewport.rows, gridMetrics.row, sheet.rows, scale, sheet.frozenPanes.rows), [sheet.rows, scale, sheet.frozenPanes.rows]);
  const columns = useMemo(() => new GridAxis(worksheetViewport.columns, gridMetrics.column, sheet.columns, scale, sheet.frozenPanes.columns), [sheet.columns, scale, sheet.frozenPanes.columns]);
  const rowTracks = rows.projected(view.y, view.height - top);
  const columnTracks = columns.projected(view.x, view.width - left);
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
      if (!onSelection(singleSelection(point))) return false;
      reveal(point);
      focus();
      return true;
    },
    moveAfterEdit(point, key, reverse) {
      const next = navigate(sheet, rows, columns, point, key, false, reverse, 1) ?? point;
      if (onSelection(singleSelection(next))) reveal(next);
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

  const editAt = props.editor?.at;
  useLayoutEffect(() => { if (editAt) reveal(editAt); }, [editAt]);

  useLayoutEffect(() => {
    drawWorksheet(canvas.current!, {
      sheet, view, rows: rowTracks, columns: columnTracks, rowAxis: rows, columnAxis: columns, selection, areas, scale, headings, gridlines,
      showFormulas, staleCells, separator, referenceArea: props.referenceArea,
      borderPreview: props.borderTool && borderPreview && ('area' in borderPreview ? {
        area: { top: borderPreview.area.row, left: borderPreview.area.column,
          bottom: borderPreview.area.row + borderPreview.area.rows - 1, right: borderPreview.area.column + borderPreview.area.columns - 1 },
      } : { segments: borderPreview.segments, pen: props.borderPen ?? { style: 'dashed' } }),
      frozenRows: rows.frozen, frozenColumns: columns.frozen, frozenHeight: rows.frozenSize, frozenWidth: columns.frozenSize,
    });
  }, [sheet, view, rowTracks, columnTracks, selection, areas, scale, headings, gridlines, showFormulas, staleCells,
    separator, props.referenceArea, props.borderTool, props.borderPen, borderPreview, rows, columns, palette]);

  function alignTo(point: Coordinate) {
    const node = scroll.current!;
    const nextRow = rows.tracks.find(track => track.index >= point.row);
    const nextColumn = columns.tracks.find(track => track.index >= point.column);
    if (nextRow) node.scrollTop = Math.max(0, nextRow.start - rows.frozenSize);
    if (nextColumn) node.scrollLeft = Math.max(0, nextColumn.start - columns.frozenSize);
    measure();
  }

  function pointAt(x: number, y: number, kind: AreaKind, bounded = false): Coordinate | undefined {
    const node = scroll.current!;
    const locate = (axis: GridAxis, position: number, header: number, offset: number, size: number) => {
      const relative = (bounded ? clamp(position, header, size - 0.01) : position) - header;
      return (axis.atScreen(relative, offset) ?? (bounded && relative >= 0 ? axis.tracks.at(-1) : undefined))?.index;
    };
    const row = kind === 'columns' || kind === 'sheet' ? rows.tracks[0]?.index :
      locate(rows, y, top, node.scrollTop, node.clientHeight);
    const column = kind === 'rows' || kind === 'sheet' ? columns.tracks[0]?.index :
      locate(columns, x, left, node.scrollLeft, node.clientWidth);
    return row !== undefined && column !== undefined ? { row, column } : undefined;
  }

  function headerAt(x: number, y: number, edge: boolean) {
    const node = scroll.current!;
    if (x < 0 || y < 0 || x >= node.clientWidth || y >= node.clientHeight) return;
    const axis: Axis | undefined = x < left && y >= top ? 'row' : y < top && x >= left ? 'column' : undefined;
    if (!axis) return;
    const coord = axis === 'row' ? y - top : x - left;
    const offset = axis === 'row' ? node.scrollTop : node.scrollLeft;
    const tracks = axis === 'row' ? rows : columns;
    const track = tracks.atScreen(edge ? Math.max(0, coord - 4) : coord, offset);
    if (track && (!edge || Math.abs(tracks.screen(track, offset) + track.size - coord) <= 4)) return { axis, track };
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
      if (active.reference) active.reference.update(primaryArea(mergeSelection(sheet, active.selection)));
      else onSelection(active.selection);
    }
    if ((xOverflow !== 0 || yOverflow !== 0) && drag.current) frame.current = requestAnimationFrame(moveDrag);
  }

  useLayoutEffect(() => {
    const active = borderDrag.current;
    if (active && active.mode !== props.borderTool) {
      borderDrag.current = null;
      setBorderPreview(undefined);
    }
  }, [props.borderTool]);

  function borderHit(x: number, y: number, erase: boolean) {
    const node = scroll.current!;
    return hitBorder(sheet, { rows, columns, scrollX: node.scrollLeft, scrollY: node.scrollTop }, x - left, y - top, erase);
  }

  function addErase(active: BorderDrag, hit: BorderHit | undefined) {
    const edge = hit?.segment;
    if (edge && active.erased.size < 26_000) active.erased.set([edge.edge, edge.row, edge.column].join(','), edge);
  }

  function updateBorder(active: BorderDrag, event: PointerEvent<HTMLDivElement>) {
    const node = event.currentTarget;
    const box = node.getBoundingClientRect();
    const rawX = event.clientX - box.left, rawY = event.clientY - box.top;
    const x = clamp(rawX, left, node.clientWidth - .01), y = clamp(rawY, top, node.clientHeight - .01);
    node.scrollLeft += clamp(rawX - x, -35, 35);
    node.scrollTop += clamp(rawY - y, -30, 30);
    if (active.mode === 'erase') {
      const dx = x - active.x, dy = y - active.y;
      const steps = Math.max(1, Math.ceil(Math.hypot(dx, dy) / 4));
      for (let step = 1; step <= steps; step++) addErase(active,
        borderHit(active.x + dx * step / steps, active.y + dy * step / steps, true));
    } else {
      const hit = borderHit(x, y, false);
      if (hit) {
        active.preview = drawGesture(active.mode, active.start, hit);
        setBorderPreview(active.preview);
      }
    }
    active.x = x;
    active.y = y;
  }

  function pointerDown(event: PointerEvent<HTMLDivElement>) {
    if (event.button !== 0 || !event.isPrimary) return;
    const node = event.currentTarget;
    const bounds = node.getBoundingClientRect();
    const x = event.clientX - bounds.left;
    const y = event.clientY - bounds.top;
    if (x >= node.clientWidth || y >= node.clientHeight) return;
    if (props.borderTool) {
      if (x >= left && y >= top) {
        const start = borderHit(x, y, props.borderTool === 'erase');
        const apply = start && props.onBeginBorder();
        if (!start || !apply) return;
        const active: BorderDrag = { id: event.pointerId, mode: props.borderTool, start, apply, x, y, erased: new Map() };
        if (active.mode === 'erase') addErase(active, start);
        else {
          active.preview = drawGesture(active.mode, start, start);
          setBorderPreview(active.preview);
        }
        borderDrag.current = active;
        event.preventDefault();
        node.setPointerCapture(event.pointerId);
        focus();
        return;
      }
      props.onExitBorder();
    }
    const header = headerAt(x, y, true);
    if (header && event.pointerType !== 'touch') {
      const apply = props.onHeaderResize(header.axis, header.track.index);
      if (!apply) return;
      const origin = header.axis === 'row' ? event.clientY : event.clientX;
      resize.current = { id: event.pointerId, axis: header.axis, origin, size: header.track.size / scale, apply };
      event.preventDefault();
      node.setPointerCapture(event.pointerId);
      return;
    }
    const kind: AreaKind = y < top ? x < left ? 'sheet' : 'columns' : x < left ? 'rows' : 'cells';
    const point = pointAt(x, y, kind);
    if (!point) return;
    const reference = kind === 'cells' && event.pointerType !== 'touch' ? props.onReferencePick() : undefined;
    if (reference) {
      const selection = singleSelection(point);
      event.preventDefault();
      node.setPointerCapture(event.pointerId);
      drag.current = { id: event.pointerId, kind, selection, x: event.clientX, y: event.clientY, reference };
      reference.update(primaryArea(mergeSelection(sheet, selection)));
      return;
    }
    const additive = event.ctrlKey || event.metaKey;
    const hitSelected = additive && areas.some(area => includes(area, point));
    const next = select(selection, point, kind, event.shiftKey, additive);
    if (!onSelection(next)) return;
    focus();
    const paint = props.paintActive && kind === 'cells' && !next.additional.length && !additive;
    if (event.pointerType === 'touch' && !paint || kind === 'sheet' || hitSelected) return;
    event.preventDefault();
    node.setPointerCapture(event.pointerId);
    drag.current = { id: event.pointerId, kind, selection: next, x: event.clientX, y: event.clientY, paint };
  }

  function pointerMove(event: PointerEvent<HTMLDivElement>) {
    const node = event.currentTarget;
    const bounds = node.getBoundingClientRect();
    const header = headerAt(event.clientX - bounds.left, event.clientY - bounds.top, true);
    const axis = resize.current?.axis ?? header?.axis;
    const heading = header ?? headerAt(event.clientX - bounds.left, event.clientY - bounds.top, false);
    node.style.cursor = axis === 'row' ? 'row-resize' : axis === 'column' ? 'col-resize' :
      heading?.axis === 'row' ? `url("${selectRowCursor}") 14 8, default` :
      heading?.axis === 'column' ? `url("${selectColumnCursor}") 8 14, default` : props.borderTool === 'erase' ? `url("${eraseBorderCursor}") 7 17, cell` :
      props.borderTool ? `url("${drawBorderCursor}") 2 18, crosshair` : props.paintActive ? 'crosshair' : '';
    const stroke = borderDrag.current;
    if (stroke?.id === event.pointerId) {
      if (stroke.mode === props.borderTool) updateBorder(stroke, event);
      return;
    }
    if (resize.current?.id === event.pointerId) {
      const offset = axis === 'row' ? event.clientY - bounds.top : event.clientX - bounds.left;
      setGuide({ axis: axis!, offset });
      return;
    }
    const active = drag.current;
    if (!active || active.id !== event.pointerId) return;
    active.x = event.clientX;
    active.y = event.clientY;
    cancelAnimationFrame(frame.current);
    moveDrag();
  }

  function endDrag(event: PointerEvent<HTMLDivElement>) {
    const stroke = borderDrag.current;
    if (stroke?.id === event.pointerId) {
      borderDrag.current = null;
      if (event.type === 'pointerup' && stroke.mode === props.borderTool) {
        updateBorder(stroke, event);
        if (stroke.mode === 'erase') stroke.apply({ segments: [...stroke.erased.values()] });
        else if (stroke.preview) stroke.apply(stroke.preview);
      }
      setBorderPreview(undefined);
      return;
    }
    const sizing = resize.current;
    if (sizing?.id === event.pointerId) {
      resize.current = null;
      setGuide(undefined);
      const end = sizing.axis === 'row' ? event.clientY : event.clientX;
      if (event.type === 'pointerup' && Math.abs(end - sizing.origin) >= 2) {
        sizing.apply(clamp(Math.round(sizing.size + (end - sizing.origin) / scale), 4, 4096));
      }
      return;
    }
    const active = drag.current;
    if (active?.id !== event.pointerId) return;
    if ((active.reference || active.paint) && event.type === 'pointerup') {
      active.x = event.clientX;
      active.y = event.clientY;
      cancelAnimationFrame(frame.current);
      moveDrag();
    } else active.reference?.cancel();
    drag.current = null;
    cancelAnimationFrame(frame.current);
    if (active.paint && event.type === 'pointerup') props.onPaint(primaryArea(mergeSelection(sheet, active.selection)));
  }

  function keyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (props.onShortcut(event)) return;
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
      if (onSelection(select(selection, point, kind)) && point !== anchor) reveal(point);
      return;
    }
    if (key === 'Escape') {
      event.preventDefault();
      exitOnTab.current = true;
      if (onSelection(singleSelection(anchor))) reveal(anchor);
      return;
    }
    if (control && key === 'Enter') return;
    const extend = shiftKey && key !== 'Enter' && key !== 'Tab';
    const point = navigate(sheet, rows, columns, extend ? selection.end : anchor, key, control, shiftKey,
      Math.max(1, rowTracks.length - 1));
    if (!point) return;
    event.preventDefault();
    if (onSelection(select(selection, point, extend ? selection.kind : 'cells', extend))) reveal(point);
  }

  // Keep the active descendant mounted even when native scrolling moves it outside the viewport.
  const projected = new Map(rowTracks.slice(0, 64).map(row => [row.index, columnTracks.map(column => column.index)]));
  function mountCell(row: number, column: number) {
    const indexes = projected.get(row) ?? [];
    if (!indexes.includes(column)) projected.set(row, [...indexes, column].sort((a, b) => a - b));
  }
  for (const [row, indexes] of [...projected]) {
    for (const column of indexes) {
      const merge = mergeAt(sheet, row, column);
      if (merge) mountCell(merge.row, merge.column);
    }
  }
  mountCell(selection.anchor.row, selection.anchor.column);
  const editMerge = editAt && mergeAt(sheet, editAt.row, editAt.column);
  const editRow = editAt && (rows.track(editAt.row) ?? (editMerge && rows.tracks.find(track =>
    track.index >= editAt.row && track.index < editAt.row + editMerge.rows)));
  const editColumn = editAt && (columns.track(editAt.column) ?? (editMerge && columns.tracks.find(track =>
    track.index >= editAt.column && track.index < editAt.column + editMerge.columns)));
  const bodyWidth = Math.max(0, view.width - left);
  const bodyHeight = Math.max(0, view.height - top);
  const editorWidth = Math.min(bodyWidth, Math.max(80, editMerge ? columns.span(editMerge.column, editMerge.column + editMerge.columns - 1) : editColumn?.size ?? 0));
  const editorHeight = Math.min(bodyHeight, Math.max(24, editMerge ? rows.span(editMerge.row, editMerge.row + editMerge.rows - 1) : editRow?.size ?? 0));
  const editorLeft = editColumn ? columns.screen(editColumn, view.x) : 0;
  const editorTop = editRow ? rows.screen(editRow, view.y) : 0;

  return (
    <div className="worksheet">
      <div ref={scroll} className="worksheet-scroll" role="grid" aria-label={sheet.name}
        aria-rowcount={worksheetViewport.rows} aria-colcount={worksheetViewport.columns}
        aria-multiselectable="true" aria-activedescendant={cellId(selection.anchor)} tabIndex={0}
        aria-describedby={id + '-help'} onBlur={() => { exitOnTab.current = false; }}
        onScroll={measure} onKeyDown={keyDown} onPointerDown={pointerDown} onPointerMove={pointerMove}
        onPointerLeave={event => { if (!resize.current) event.currentTarget.style.cursor = ''; }}
        onCopy={event => props.onClipboard(event, 'copy')} onCut={event => props.onClipboard(event, 'cut')}
        onPaste={event => props.onClipboard(event, 'paste')}
        onContextMenu={event => {
          const bounds = event.currentTarget.getBoundingClientRect();
          const x = event.clientX - bounds.left, y = event.clientY - bounds.top;
          const header = headerAt(x, y, false);
          const point = !header && x >= left && y >= top && pointAt(x, y, 'cells');
          if (!header && !point) return;
          event.preventDefault();
          if (point && !areas.some(area => includes(area, point)) && !onSelection(singleSelection(point))) return;
          props.onStructureMenu(event.clientX, event.clientY, header?.axis, header?.track.index, point || undefined);
        }}
        onDoubleClick={event => {
          const node = event.currentTarget;
          const bounds = node.getBoundingClientRect();
          const x = event.clientX - bounds.left, y = event.clientY - bounds.top;
          const header = headerAt(x, y, true);
          if (header) {
            props.onHeaderFit(header.axis, header.track.index);
            return;
          }
          if (!props.referenceArea && !props.borderTool && x >= left && y >= top && x < node.clientWidth && y < node.clientHeight &&
              pointAt(x, y, 'cells')) props.onEdit();
        }}
        onPointerUp={endDrag} onPointerCancel={endDrag} onLostPointerCapture={endDrag}>
        <div aria-hidden="true" className="worksheet-spacer" style={{ width: left + columns.total, height: top + rows.total }} />
        <div className="worksheet-projection">
          {[...projected].sort(([a], [b]) => a - b).filter(([, indexes]) => indexes.length > 0).map(([row, indexes]) => (
            <div key={row} role="row" aria-rowindex={row}>
              {indexes.map(column => {
                const point = { row, column };
                const merge = mergeAt(sheet, row, column);
                if (merge && (merge.row !== row || merge.column !== column)) return null;
                const cell = findCell(sheet, point);
                const value = cellDisplay(cell, showFormulas, separator);
                return <div key={column} id={cellId(point)} role="gridcell" aria-colindex={column} aria-rowspan={merge?.rows} aria-colspan={merge?.columns}
                  aria-selected={areas.some(area => includes(area, point))}
                  aria-description={cell && staleCells.has(cell.id) ? 'Calculation needed' : undefined}
                  aria-label={address(point) + (value ? ', ' + value : ', blank')} />;
              })}
            </div>
          ))}
        </div>
      </div>
      <p className="worksheet-projection" id={id + '-help'}>
        Use arrow keys or Tab to move between cells. Press F2 to edit. Press Escape then Tab to leave the worksheet.
      </p>
      <canvas ref={canvas} aria-hidden="true" className="worksheet-canvas" style={{ width: view.width, height: view.height }} />
      {guide && <div aria-hidden="true" className={'dimension-guide ' + guide.axis}
        style={guide.axis === 'row' ? { top: guide.offset } : { left: guide.offset }} />}
      {editRow && editColumn && <div className="inline-editor-layer" style={{
        left, top, width: bodyWidth, height: bodyHeight,
      }}>
        <div className="inline-cell-editor" style={{
          left: editorLeft < bodyWidth ? Math.min(editorLeft, bodyWidth - editorWidth) : editorLeft,
          top: editorTop < bodyHeight ? Math.min(editorTop, bodyHeight - editorHeight) : editorTop,
          width: editorWidth, height: editorHeight, fontSize: Math.max(10, 14 * scale),
        }}>{props.editor?.input}</div>
      </div>}
    </div>
  );
}
