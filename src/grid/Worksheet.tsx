import {
  memo, useEffect, useId, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState,
  type ClipboardEvent, type KeyboardEvent, type PointerEvent, type ReactNode, type Ref,
} from 'react';
import type { BorderMode, BorderSegment } from '../core/formats/borders';
import { mergeAt } from '../application/border-commands';
import { drawGesture, hitBorder, type BorderHit, type DrawGesture } from './border-geometry';
import { findCell, resolvedCellStyle, worksheetViewport, type CellBorder, type Coordinate, type SheetDocument, type WorkbookName } from '../core/workbook';
import { extentContains, lastColumn, lastRow, tableDesign } from '../core/workbook-names';
import { filteredRows, headerLabel } from '../application/data-commands';
import { totalFunction } from '../application/table-commands';
import { TotalRowDropdown } from '../features/tables/TotalRowDropdown';
import type { Axis } from '../core/axis-edit';
import { RowMask } from '../core/row-mask';
import type { AfterEditKey } from '../features/editing/useWorksheetEditing';
import { drawWorksheet, type GridViewport } from './draw-worksheet';
import type { FigurePlacement, PythonFigure } from '../core/workbook';
import { figureLogicalRect, figurePlacementFromRect, figureScreenRect, logicalPoint,
  type FigureRect } from './figure-geometry';
import drawBorderCursor from './cursors/draw-border.svg';
import eraseBorderCursor from './cursors/erase-border.svg';
import selectColumnCursor from './cursors/select-column.svg';
import selectRowCursor from './cursors/select-row.svg';
import { address, capFromScrollport, clamp, contentOffset, GridAxis, gridMetrics, placeScroll, rowHeaderWidth, scrollSpacer } from './grid-geometry';
import {
  cellDisplay, formulaSpillArea, includes, mergeSelection, navigate, primaryArea, select, selectionAreas, singleSelection,
  type Area, type AreaKind, type GridSelection,
} from './grid-selection';

export interface WorksheetHandle {
  focus(): void;
  goTo(point: Coordinate): boolean;
  reveal(point: Coordinate, focus?: boolean): void;
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
  readonly calculatingCells: ReadonlySet<string>;
  readonly separator: ',' | ';';
  readonly tables?: readonly WorkbookName[];
  readonly rangePicking?: boolean;
  readonly onRangePickFinish?: (accept: boolean) => void;
  readonly onResizeTable?: (table: WorkbookName, rows: number, columns: number) => void;
  readonly onTotalFunction?: (tableId: string, column: number, code: number) => void;
  readonly onFilter?: (row: number, column: number, x: number, y: number) => void;
  readonly figures?: readonly WorksheetFigure[];
  readonly pendingFigure?: { readonly ownerId: string; readonly figure: PythonFigure };
  readonly figureStatus?: { readonly ownerId: string; readonly message: string };
  readonly onFigurePlacement?: (ownerId: string, placement: FigurePlacement, original?: FigurePlacement) => void;
  readonly onFigureDelete?: (placementId: string) => void;
  readonly onFigureCancel?: () => void;
  readonly ref: Ref<WorksheetHandle>;
}

export interface WorksheetFigure {
  readonly ownerId: string;
  readonly figure: PythonFigure;
  readonly placement: FigurePlacement;
  readonly label?: string;
}

type FigureEdge = 'left' | 'top' | 'right' | 'bottom' | 'top-left' | 'top-right' | 'bottom-right' | 'bottom-left';
const figureEdges: readonly FigureEdge[] = ['top-left', 'top', 'top-right', 'right', 'bottom-right', 'bottom', 'bottom-left', 'left'];
interface FigureDrag {
  readonly pointerId: number;
  readonly ownerId: string;
  readonly figure: PythonFigure;
  readonly original?: FigurePlacement;
  readonly mode: 'place' | 'move' | 'resize';
  readonly edge?: FigureEdge;
  readonly start: { readonly x: number; readonly y: number };
  readonly initial: FigureRect;
}
interface FigureDraft {
  readonly ownerId: string;
  readonly placementId?: string;
  readonly box: FigureRect;
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
  readonly finish?: () => void;
  readonly cancel: () => void;
}

const initialViewport: GridViewport = { x: 0, y: 0, width: 0, height: 0 };

interface ProjectedRowProps {
  readonly row: number;
  readonly columns: readonly number[];
  readonly sheet: SheetDocument;
  readonly showFormulas: boolean;
  readonly separator: ',' | ';';
  readonly staleCells: ReadonlySet<string>;
  readonly calculatingCells: ReadonlySet<string>;
  readonly idPrefix: string;
  /** >= 1 when a single column is selected in this row; -1 when none; -2 when using `areas`. */
  readonly selectedColumn: number;
  readonly areas: readonly Area[] | undefined;
}

const ProjectedRow = memo(function ProjectedRow({
  row, columns, sheet, showFormulas, separator, staleCells, calculatingCells, idPrefix, selectedColumn, areas,
}: ProjectedRowProps) {
  return (
    <div role="row" aria-rowindex={row}>
      {columns.map(column => {
        const point = { row, column };
        const merge = sheet.merges.length ? mergeAt(sheet, row, column) : undefined;
        if (merge && (merge.row !== row || merge.column !== column)) return null;
        const cell = findCell(sheet, point);
        const value = cellDisplay(cell, showFormulas, separator, resolvedCellStyle(sheet, point, cell)?.numberFormat);
        const labelAddress = address(point);
        const selected = selectedColumn === -1 ? false :
          selectedColumn > 0 ? column === selectedColumn :
          areas!.some(area => includes(area, point));
        return <div key={column} id={idPrefix + '-' + labelAddress} role="gridcell" aria-colindex={column} aria-rowspan={merge?.rows} aria-colspan={merge?.columns}
          aria-selected={selected}
          aria-readonly={cell?.spillOwner ? 'true' : undefined}
          aria-description={cell && calculatingCells.has(cell.id) ? 'Calculating' :
            cell && staleCells.has(cell.id) ? 'Calculation needed' : undefined}
          aria-label={labelAddress + (value ? ', ' + value : ', blank')} />;
      })}
    </div>
  );
});

export function Worksheet(props: WorksheetProps) {
  const { sheet, selection, onSelection, onAdjacentSheet, zoom, headings, gridlines, showFormulas,
    staleCells, calculatingCells, separator } = props;
  const scroll = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const drag = useRef<Drag | null>(null);
  const borderDrag = useRef<BorderDrag | null>(null);
  const [borderPreview, setBorderPreview] = useState<DrawGesture>();
  const [resizePreview, setResizePreview] = useState<Area>();
  const resize = useRef<{ id: number; axis: Axis; origin: number; boundOrigin: number; size: number; apply: (size: number) => void } | null>(null);
  const [guide, setGuide] = useState<{ axis: Axis; offset: number }>();
  const guideRef = useRef<HTMLDivElement>(null);
  const exitOnTab = useRef(false);
  const fitStart = useRef<Coordinate | null>(null);
  const frame = useRef(0);
  const id = useId();
  const [view, setView] = useState(initialViewport);
  const [projectionView, setProjectionView] = useState(initialViewport);
  const lastScroll = useRef({ left: 0, top: 0 });
  const [palette, setPalette] = useState(0);
  const [selectedFigureId, setSelectedFigureId] = useState<string>();
  const pendingFigureSelection = useRef<string | undefined>(undefined);
  const [figureDraft, setFigureDraft] = useState<FigureDraft>();
  const [figureCursor, setFigureCursor] = useState<{ readonly ownerId: string; readonly x: number; readonly y: number }>();
  const figureDrag = useRef<FigureDrag | null>(null);
  const scale = zoom / 100;
  const [fontGeneration, setFontGeneration] = useState(0);
  useEffect(() => {
    const fonts = document.fonts;
    if (!fonts?.ready) return undefined;
    let active = true;
    const refresh = () => { if (active) setFontGeneration(generation => generation + 1); };
    void fonts.ready.then(refresh);
    fonts.addEventListener('loadingdone', refresh);
    return () => {
      active = false;
      fonts.removeEventListener('loadingdone', refresh);
    };
  }, []);
  const rowHeader = useMemo(() => {
    if (typeof document === 'undefined') return gridMetrics.rowHeader * scale;
    const probe = document.createElement('canvas').getContext('2d');
    if (!probe) return gridMetrics.rowHeader * scale;
    probe.font = `${12 * scale}px ${getComputedStyle(document.documentElement).fontFamily}`;
    return rowHeaderWidth(probe.measureText(String(worksheetViewport.rows)).width, scale);
  }, [scale, fontGeneration]);
  const left = headings ? rowHeader : 0;
  const top = headings ? gridMetrics.columnHeader * scale : 0;
  // Filter-hidden rows keep their index but leave the projected axis, like the desktop worksheet.
  const filteredRef = useRef<RowMask>(RowMask.none());
  const filtered = useMemo(() => {
    const next = filteredRows(sheet);
    if (next.equals(filteredRef.current)) return filteredRef.current;
    return (filteredRef.current = next);
  }, [sheet]);
  const rows = useMemo(() => new GridAxis(worksheetViewport.rows, gridMetrics.row, sheet.rows, scale,
    sheet.frozenPanes.rows, filtered.spans()), [sheet.rows, scale, sheet.frozenPanes.rows, filtered]);
  const columns = useMemo(() => new GridAxis(worksheetViewport.columns, gridMetrics.column, sheet.columns, scale, sheet.frozenPanes.columns), [sheet.columns, scale, sheet.frozenPanes.columns]);
  function scrollOrigin(node: HTMLElement) {
    const portX = Math.max(0, node.clientWidth - left);
    const portY = Math.max(0, node.clientHeight - top);
    return {
      x: contentOffset(node.scrollLeft, columns.total, capFromScrollport(node.scrollWidth, left, columns.total), portX),
      y: contentOffset(node.scrollTop, rows.total, capFromScrollport(node.scrollHeight, top, rows.total), portY),
      portX, portY,
    };
  }
  function placeHorizontal(node: HTMLElement, offset: number) {
    const port = Math.max(0, node.clientWidth - left);
    const next = placeScroll(offset, columns.total, node.scrollLeft, capFromScrollport(node.scrollWidth, left, columns.total), port);
    if (next !== node.scrollLeft) node.scrollLeft = next;
  }
  function placeVertical(node: HTMLElement, offset: number) {
    const port = Math.max(0, node.clientHeight - top);
    const next = placeScroll(offset, rows.total, node.scrollTop, capFromScrollport(node.scrollHeight, top, rows.total), port);
    if (next !== node.scrollTop) node.scrollTop = next;
  }
  useEffect(() => {
    if (projectionView.width === 0 || projectionView.width !== view.width || projectionView.height !== view.height) {
      setProjectionView(view);
      return undefined;
    }
    if (projectionView.x === view.x && projectionView.y === view.y) return undefined;
    const timer = window.setTimeout(() => setProjectionView(view), 80);
    return () => window.clearTimeout(timer);
  }, [view, projectionView.x, projectionView.y, projectionView.width, projectionView.height]);
  const rowTracks = useMemo(() => rows.projected(view.y, view.height - top), [rows, view.y, view.height, top]);
  const columnTracks = useMemo(() => columns.projected(view.x, view.width - left), [columns, view.x, view.width, left]);
  const projectedRowTracks = useMemo(() => rows.projected(projectionView.y, projectionView.height - top),
    [rows, projectionView.y, projectionView.height, top]);
  const projectedColumnIndexes = useMemo(() => columns.projected(projectionView.x, projectionView.width - left).map(column => column.index),
    [columns, projectionView.x, projectionView.width, left]);
  const areas = selectionAreas(selection);
  const cellId = (point: Coordinate) => id + '-' + address(point);

  function measure() {
    const node = scroll.current;
    if (!node) return;
    lastScroll.current = { left: node.scrollLeft, top: node.scrollTop };
    const origin = scrollOrigin(node);
    const next = { x: origin.x, y: origin.y, width: node.clientWidth, height: node.clientHeight };
    setView(current => current.x === next.x && current.y === next.y &&
      current.width === next.width && current.height === next.height ? current : next);
  }

  function handleScroll() {
    const node = scroll.current;
    if (!node) return;
    if (node.scrollLeft === lastScroll.current.left && node.scrollTop === lastScroll.current.top) return;
    measure();
  }

  useEffect(() => {
    if (pendingFigureSelection.current && props.figures?.some(item => item.placement.id === pendingFigureSelection.current)) {
      setSelectedFigureId(pendingFigureSelection.current);
      pendingFigureSelection.current = undefined;
    }
    if (selectedFigureId && !props.figures?.some(item => item.placement.id === selectedFigureId)) setSelectedFigureId(undefined);
  }, [props.figures, selectedFigureId]);

  useEffect(() => {
    const draft = figureDraft;
    if (!draft?.placementId) return;
    const placement = props.figures?.find(item => item.placement.id === draft.placementId)?.placement;
    const message = props.figureStatus?.ownerId === draft.ownerId ? props.figureStatus.message : '';
    if (/cancelled|unchanged|remains|no longer available|could not recalculate|failed|was not applied/iu.test(message)) {
      setFigureDraft(undefined);
      return;
    }
    if (!placement) return;
    const saved = figureLogicalRect(placement, rows, columns, scale);
    if (saved && Math.abs(saved.x - draft.box.x) < 1 && Math.abs(saved.y - draft.box.y) < 1 &&
        Math.abs(saved.width - draft.box.width) < 1 && Math.abs(saved.height - draft.box.height) < 1) {
      setFigureDraft(undefined);
    }
  }, [figureDraft, props.figures, props.figureStatus, rows, columns, scale]);

  function figurePoint(clientX: number, clientY: number) {
    const node = scroll.current;
    if (!node) return { x: 0, y: 0 };
    const bounds = node.getBoundingClientRect();
    const origin = scrollOrigin(node);
    return {
      x: logicalPoint(clientX - bounds.left, left, columns, origin.x, scale),
      y: logicalPoint(clientY - bounds.top, top, rows, origin.y, scale),
    };
  }

  function figureBoxForPoint(active: FigureDrag, point: { readonly x: number; readonly y: number }): FigureRect {
    const dx = point.x - active.start.x, dy = point.y - active.start.y;
    const bounds = { width: columns.total / scale, height: rows.total / scale };
    let box: FigureRect;
    if (active.mode === 'place') {
      box = { x: Math.min(active.start.x, point.x), y: Math.min(active.start.y, point.y),
        width: Math.abs(dx), height: Math.abs(dy) };
    } else if (active.mode === 'move') {
      box = { ...active.initial,
        x: clamp(active.initial.x + dx, 0, Math.max(0, bounds.width - active.initial.width)),
        y: clamp(active.initial.y + dy, 0, Math.max(0, bounds.height - active.initial.height)) };
    } else {
      const edge = active.edge!;
      let x = active.initial.x, y = active.initial.y;
      let right = x + active.initial.width, bottom = y + active.initial.height;
      if (edge.includes('left')) x = clamp(x + dx, 0, right - 80);
      if (edge.includes('right')) right = clamp(right + dx, x + 80, bounds.width);
      if (edge.includes('top')) y = clamp(y + dy, 0, bottom - 60);
      if (edge.includes('bottom')) bottom = clamp(bottom + dy, y + 60, bounds.height);
      box = { x, y, width: right - x, height: bottom - y };
    }
    return box;
  }

  function updateFigureDraft(event: PointerEvent<HTMLDivElement>) {
    const active = figureDrag.current;
    if (!active || active.pointerId !== event.pointerId) return;
    const box = figureBoxForPoint(active, figurePoint(event.clientX, event.clientY));
    setFigureDraft({ ownerId: active.ownerId, ...(active.original && { placementId: active.original.id }), box });
  }

  function beginFigureDrag(event: PointerEvent<HTMLDivElement>, mode: FigureDrag['mode'],
    ownerId: string, figure: PythonFigure, original?: FigurePlacement, edge?: FigureEdge) {
    event.preventDefault();
    event.stopPropagation();
    const point = figurePoint(event.clientX, event.clientY);
    const initial = original ? figureLogicalRect(original, rows, columns, scale) : undefined;
    const sourceAspect = figure.height / Math.max(1, figure.width);
    const box = initial ?? { x: point.x, y: point.y, width: 480, height: 480 * sourceAspect };
    figureDrag.current = { pointerId: event.pointerId, ownerId, figure, ...(original && { original }),
      mode, ...(edge && { edge }), start: point, initial: box };
    setSelectedFigureId(original?.id);
    setFigureDraft({ ownerId, ...(original && { placementId: original.id }), box });
    event.currentTarget.setPointerCapture(event.pointerId);
    scroll.current?.focus({ preventScroll: true });
  }

  function finishFigureDrag(event: PointerEvent<HTMLDivElement>) {
    const active = figureDrag.current;
    if (!active || active.pointerId !== event.pointerId) return;
    const point = figurePoint(event.clientX, event.clientY);
    const latest = figureBoxForPoint(active, point);
    // A click chooses the default aspect-preserving size; a drag uses the live rectangle.
    const box = active.mode === 'place' && Math.hypot(point.x - active.start.x, point.y - active.start.y) < 8 ? active.initial : latest;
    const moved = active.mode !== 'move' || Math.hypot(latest.x - active.initial.x, latest.y - active.initial.y) >= 1;
    const resized = active.mode !== 'resize' || Math.hypot(latest.width - active.initial.width,
      latest.height - active.initial.height) >= 1;
    let keepDraft = false;
    try {
      if (moved && resized) {
        const placementId = active.original?.id ?? crypto.randomUUID();
        const placement = figurePlacementFromRect(placementId, sheet.id, box, rows, columns, scale);
        if (active.original) setSelectedFigureId(placementId);
        else pendingFigureSelection.current = placementId;
        props.onFigurePlacement?.(active.ownerId, placement, active.original);
        keepDraft = !!props.onFigurePlacement && !!active.original;
      }
    } catch { /* An out-of-range pointer leaves the previous workbook placement intact. */ }
    figureDrag.current = null;
    setFigureDraft(keepDraft ? { ownerId: active.ownerId, placementId: active.original!.id, box } : undefined);
    setFigureCursor(undefined);
  }

  function cancelFigureDrag(event?: PointerEvent<HTMLDivElement>) {
    if (event && figureDrag.current?.pointerId !== event.pointerId) return;
    figureDrag.current = null;
    setFigureDraft(undefined);
  }

  function reveal(point: Coordinate) {
    const node = scroll.current;
    if (!node) return;
    const width = node.clientWidth;
    const height = node.clientHeight;
    const portX = Math.max(0, width - left);
    const portY = Math.max(0, height - top);
    const capX = capFromScrollport(node.scrollWidth, left, columns.total);
    const capY = capFromScrollport(node.scrollHeight, top, rows.total);
    const currentLeft = node.scrollLeft;
    const currentTop = node.scrollTop;
    const originX = contentOffset(currentLeft, columns.total, capX, portX);
    const originY = contentOffset(currentTop, rows.total, capY, portY);
    const targetX = columns.reveal(point.column, originX, portX);
    const targetY = rows.reveal(point.row, originY, portY);
    let moved = false;
    if (targetX !== originX) {
      const nextLeft = placeScroll(targetX, columns.total, currentLeft, capX, portX);
      if (nextLeft !== currentLeft) {
        node.scrollLeft = nextLeft;
        moved = true;
      }
    }
    if (targetY !== originY) {
      const nextTop = placeScroll(targetY, rows.total, currentTop, capY, portY);
      if (nextTop !== currentTop) {
        node.scrollTop = nextTop;
        moved = true;
      }
    }
    if (!moved && view.x === originX && view.y === originY && view.width === width && view.height === height) return;
    const finalLeft = moved ? node.scrollLeft : currentLeft;
    const finalTop = moved ? node.scrollTop : currentTop;
    lastScroll.current = { left: finalLeft, top: finalTop };
    const next = {
      x: finalLeft === currentLeft ? originX : contentOffset(finalLeft, columns.total, capX, portX),
      y: finalTop === currentTop ? originY : contentOffset(finalTop, rows.total, capY, portY),
      width,
      height,
    };
    setView(current => current.x === next.x && current.y === next.y &&
      current.width === next.width && current.height === next.height ? current : next);
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
    reveal(point, keepFocus = true) {
      reveal(point);
      if (keepFocus) focus();
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
      showFormulas, staleCells, calculatingCells, separator, referenceArea: props.referenceArea, tables: props.tables,
      borderPreview: props.borderTool && borderPreview && ('area' in borderPreview ? {
        area: { top: borderPreview.area.row, left: borderPreview.area.column,
          bottom: borderPreview.area.row + borderPreview.area.rows - 1, right: borderPreview.area.column + borderPreview.area.columns - 1 },
      } : { segments: borderPreview.segments, pen: props.borderPen ?? { style: 'dashed' } }),
      frozenRows: rows.frozen, frozenColumns: columns.frozen, frozenHeight: rows.frozenSize, frozenWidth: columns.frozenSize,
      resizePreview, rowHeaderWidth: rowHeader, themeVersion: palette,
    });
  }, [sheet, view, rowTracks, columnTracks, selection, areas, scale, headings, gridlines, showFormulas, staleCells, calculatingCells,
    separator, props.referenceArea, props.tables, props.borderTool, props.borderPen, borderPreview, resizePreview, rows, columns, palette, rowHeader]);

  function alignTo(point: Coordinate) {
    const node = scroll.current!;
    const nextRow = rows.trackAtOrAfter(point.row);
    const nextColumn = columns.trackAtOrAfter(point.column);
    if (nextRow) placeVertical(node, Math.max(0, nextRow.start - rows.frozenSize));
    if (nextColumn) placeHorizontal(node, Math.max(0, nextColumn.start - columns.frozenSize));
    measure();
  }

  function pointAt(x: number, y: number, kind: AreaKind, bounded = false): Coordinate | undefined {
    const node = scroll.current!;
    const locate = (axis: GridAxis, position: number, header: number, offset: number, size: number) => {
      const relative = (bounded ? clamp(position, header, size - 0.01) : position) - header;
      return (axis.atScreen(relative, offset) ?? (bounded && relative >= 0 ? axis.last : undefined))?.index;
    };
    const origin = scrollOrigin(node);
    const row = kind === 'columns' || kind === 'sheet' ? rows.first?.index :
      locate(rows, y, top, origin.y, node.clientHeight);
    const column = kind === 'rows' || kind === 'sheet' ? columns.first?.index :
      locate(columns, x, left, origin.x, node.clientWidth);
    return row !== undefined && column !== undefined ? { row, column } : undefined;
  }

  function headerAt(x: number, y: number, edge: boolean) {
    const node = scroll.current!;
    if (x < 0 || y < 0 || x >= node.clientWidth || y >= node.clientHeight) return;
    const axis: Axis | undefined = x < left && y >= top ? 'row' : y < top && x >= left ? 'column' : undefined;
    if (!axis) return;
    const coord = axis === 'row' ? y - top : x - left;
    const origin = scrollOrigin(node);
    const offset = axis === 'row' ? origin.y : origin.x;
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
    const origin = scrollOrigin(node);
    placeHorizontal(node, origin.x + clamp(xOverflow, -35, 35));
    placeVertical(node, origin.y + clamp(yOverflow, -30, 30));
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
    const origin = scrollOrigin(node);
    return hitBorder(sheet, { rows, columns, scrollX: origin.x, scrollY: origin.y }, x - left, y - top, erase);
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
    const origin = scrollOrigin(node);
    placeHorizontal(node, origin.x + clamp(rawX - x, -35, 35));
    placeVertical(node, origin.y + clamp(rawY - y, -30, 30));
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
    setSelectedFigureId(undefined);
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
      const boundOrigin = header.axis === 'row' ? bounds.top : bounds.left;
      resize.current = { id: event.pointerId, axis: header.axis, origin, boundOrigin, size: header.track.size / scale, apply };
      setGuide({ axis: header.axis, offset: origin - boundOrigin });
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
    const paint = props.paintActive && !next.additional.length && !additive;
    if (paint && kind === 'sheet') {
      event.preventDefault();
      props.onPaint(primaryArea(next));
      return;
    }
    if (event.pointerType === 'touch' && !paint || kind === 'sheet' || hitSelected) return;
    event.preventDefault();
    node.setPointerCapture(event.pointerId);
    drag.current = { id: event.pointerId, kind, selection: next, x: event.clientX, y: event.clientY, paint };
  }

  function pointerMove(event: PointerEvent<HTMLDivElement>) {
    const sizing = resize.current;
    if (sizing?.id === event.pointerId) {
      const offset = (sizing.axis === 'row' ? event.clientY : event.clientX) - sizing.boundOrigin;
      if (guideRef.current) {
        guideRef.current.style.transform = sizing.axis === 'row'
          ? `translate3d(0, ${offset}px, 0)`
          : `translate3d(${offset}px, 0, 0)`;
      }
      return;
    }
    const node = event.currentTarget;
    const bounds = node.getBoundingClientRect();
    const header = headerAt(event.clientX - bounds.left, event.clientY - bounds.top, true);
    const axis = header?.axis;
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
      active.reference?.finish?.();
    } else active.reference?.cancel();
    drag.current = null;
    cancelAnimationFrame(frame.current);
    if (active.paint && event.type === 'pointerup') props.onPaint(primaryArea(mergeSelection(sheet, active.selection)));
  }

  function keyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (props.pendingFigure && event.key === 'Escape') {
      event.preventDefault();
      props.onFigureCancel?.();
      setFigureCursor(undefined);
      setFigureDraft(undefined);
      return;
    }
    if (selectedFigureId && (event.key === 'Delete' || event.key === 'Backspace')) {
      event.preventDefault();
      props.onFigureDelete?.(selectedFigureId);
      setSelectedFigureId(undefined);
      return;
    }
    if (selectedFigureId && event.key === 'Escape') {
      event.preventDefault();
      setSelectedFigureId(undefined);
      return;
    }
    if (props.rangePicking) {
      if (event.key === 'Enter' || event.key === 'Escape') {
        event.preventDefault();
        props.onRangePickFinish?.(event.key === 'Enter');
        return;
      }
      if (event.key === 'Tab') return;
      if (!['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End', 'PageUp', 'PageDown'].includes(event.key)) {
        event.preventDefault();
        return;
      }
    } else if (props.onShortcut(event)) return;
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
        row: rows.first?.index ?? anchor.row, column: columns.first?.index ?? anchor.column,
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
    if (!point) {
      if (props.rangePicking) event.preventDefault();
      return;
    }
    event.preventDefault();
    if (onSelection(select(selection, point, extend ? selection.kind : 'cells', extend))) reveal(point);
  }

  // Keep the active descendant mounted even when native scrolling moves it outside the viewport.
  const projected = new Map<number, readonly number[]>(projectedRowTracks.slice(0, 64).map(row => [row.index, projectedColumnIndexes]));
  let sortProjectedRows = false;
  function mountCell(row: number, column: number) {
    const indexes = projected.get(row);
    if (!indexes) {
      projected.set(row, [column]);
      sortProjectedRows = true;
    } else if (!indexes.includes(column)) {
      projected.set(row, [...indexes, column].sort((a, b) => a - b));
    }
  }
  if (sheet.merges.length) {
    for (const [row, indexes] of [...projected]) {
      for (const column of indexes) {
        const merge = mergeAt(sheet, row, column);
        if (merge) mountCell(merge.row, merge.column);
      }
    }
  }
  mountCell(selection.anchor.row, selection.anchor.column);
  if (selection.end.row !== selection.anchor.row || selection.end.column !== selection.anchor.column) {
    mountCell(selection.end.row, selection.end.column);
  }
  const editMerge = editAt && mergeAt(sheet, editAt.row, editAt.column);
  const editRowTrack = editAt && rows.trackAtOrAfter(editAt.row);
  const editColumnTrack = editAt && columns.trackAtOrAfter(editAt.column);
  const editRow = editAt && (rows.track(editAt.row) ?? (editMerge && editRowTrack &&
    editRowTrack.index < editAt.row + editMerge.rows ? editRowTrack : undefined));
  const editColumn = editAt && (columns.track(editAt.column) ?? (editMerge && editColumnTrack &&
    editColumnTrack.index < editAt.column + editMerge.columns ? editColumnTrack : undefined));
  const bodyWidth = Math.max(0, view.width - left);
  const bodyHeight = Math.max(0, view.height - top);
  const editorWidth = Math.min(bodyWidth, Math.max(80, editMerge ? columns.span(editMerge.column, editMerge.column + editMerge.columns - 1) : editColumn?.size ?? 0));
  const editorHeight = Math.min(bodyHeight, Math.max(24, editMerge ? rows.span(editMerge.row, editMerge.row + editMerge.rows - 1) : editRow?.size ?? 0));
  const editorLeft = editColumn ? columns.screen(editColumn, view.x) : 0;
  const editorTop = editRow ? rows.screen(editRow, view.y) : 0;

  function cellBounds(row: number, column: number) {
    const rTrack = rows.track(row);
    const cTrack = columns.track(column);
    if (!rTrack || !cTrack) return undefined;
    const screenX = left + columns.screen(cTrack, view.x);
    const screenY = top + rows.screen(rTrack, view.y);
    return {
      left: screenX,
      top: screenY,
      right: screenX + cTrack.size,
      bottom: screenY + rTrack.size,
      width: cTrack.size,
      height: rTrack.size,
    };
  }

  function onStartResizeHandle(event: PointerEvent<HTMLButtonElement>, table: WorkbookName) {
    event.stopPropagation();
    event.preventDefault();
    const button = event.currentTarget;
    button.setPointerCapture(event.pointerId);
    const initialTable = table;
    let latestRows = table.area.rows;
    let latestCols = table.area.columns;

    function handlePointerMove(e: globalThis.PointerEvent) {
      const node = scroll.current;
      if (!node) return;
      const rect = node.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      const pt = pointAt(x, y, 'cells', true);
      if (!pt) return;
      latestRows = Math.max(1, pt.row - initialTable.area.row + 1);
      latestCols = Math.max(1, pt.column - initialTable.area.column + 1);
      setResizePreview({
        top: initialTable.area.row,
        left: initialTable.area.column,
        bottom: initialTable.area.row + latestRows - 1,
        right: initialTable.area.column + latestCols - 1,
      });
    }

    function handlePointerUp(e: globalThis.PointerEvent) {
      button.releasePointerCapture(e.pointerId);
      button.removeEventListener('pointermove', handlePointerMove);
      button.removeEventListener('pointerup', handlePointerUp);
      button.removeEventListener('pointercancel', handlePointerUp);
      setResizePreview(undefined);
      if (latestRows !== initialTable.area.rows || latestCols !== initialTable.area.columns) {
        props.onResizeTable?.(initialTable, latestRows, latestCols);
      }
    }

    button.addEventListener('pointermove', handlePointerMove);
    button.addEventListener('pointerup', handlePointerUp);
    button.addEventListener('pointercancel', handlePointerUp);
  }

  function onHandleKeyDown(event: KeyboardEvent<HTMLButtonElement>, table: WorkbookName) {
    let dr = 0, dc = 0;
    switch (event.key) {
      case 'ArrowUp': dr = -1; break;
      case 'ArrowDown': dr = 1; break;
      case 'ArrowLeft': dc = -1; break;
      case 'ArrowRight': dc = 1; break;
      default: return;
    }
    event.preventDefault();
    event.stopPropagation();
    const newR = Math.max(1, table.area.rows + dr);
    const newC = Math.max(1, table.area.columns + dc);
    if (newR !== table.area.rows || newC !== table.area.columns) {
      props.onResizeTable?.(table, newR, newC);
    }
  }

  function filterControls() {
    const filter = sheet.filter;
    const tables = (props.tables ?? []).filter(table => !table.area.invalid && table.area.sheetId === sheet.id);
    if (!filter && !tables.length) return [];
    const headers = new Set<number>();
    if (filter) headers.add(filter.row);
    for (const table of tables) {
      const design = tableDesign(table);
      if (design.headerRow && design.filterButton) headers.add(table.area.row);
    }
    function owns(row: number, column: number) {
      return !!filter && filter.row === row && column >= filter.column && column < filter.column + filter.columns ||
        tables.some(table => table.area.row === row && extentContains(table.area, { row, column }) &&
          tableDesign(table).headerRow && tableDesign(table).filterButton);
    }
    const controls: { row: number; column: number; label: string; active: boolean;
      sort: { readonly ascending: boolean } | undefined; left: number; top: number; width: number; height: number }[] = [];
    const dividerThickness = Math.max(1, 2 * scale);
    const splitX = left + columns.frozenSize;
    const splitY = top + rows.frozenSize;
    for (const row of rowTracks) {
      if (!headers.has(row.index)) continue;
      for (const column of columnTracks) {
        if (!owns(row.index, column.index)) continue;
        const bounds = cellBounds(row.index, column.index);
        if (!bounds) continue;
        const sort = filter && filter.row === row.index && filter.sort?.column === column.index ? filter.sort : undefined;
        const intendedWidth = Math.min(bounds.width, (sort ? 27.5 : 18) * scale);
        const controlLeft = bounds.right - intendedWidth;
        const overlapsFrozenColumn = columns.frozen > 0 && bounds.left < splitX && bounds.right >= splitX - dividerThickness;
        const overlapsFrozenRow = rows.frozen > 0 && bounds.top < splitY && bounds.bottom >= splitY - dividerThickness;
        const controlRight = overlapsFrozenColumn ? Math.min(bounds.right, splitX - dividerThickness) : bounds.right;
        const controlBottom = overlapsFrozenRow ? Math.min(bounds.bottom, splitY - dividerThickness) : bounds.bottom;
        const width = controlRight - controlLeft;
        const height = controlBottom - bounds.top;
        if (width < 8 || height < 8) continue;
        if (controlLeft < left || controlRight > view.width || bounds.top < top || controlBottom > view.height) continue;
        controls.push({ row: row.index, column: column.index, label: headerLabel(sheet, row.index, column.index),
          active: !!filter && filter.row === row.index && filter.criteria.some(item => item.column === column.index),
          sort, left: controlLeft, top: bounds.top, width, height });
      }
    }
    return controls;
  }

  const activeSpill = formulaSpillArea(sheet, selection.anchor);
  const activeSpillRange = activeSpill
    ? `${address({ row: activeSpill.top, column: activeSpill.left })}:${address({ row: activeSpill.bottom, column: activeSpill.right })}`
    : undefined;

  return (
    <div className="worksheet">
      <div ref={scroll} className="worksheet-scroll" role="grid" aria-label={sheet.name} data-row-header={left}
        data-spill-range={activeSpillRange}
        aria-rowcount={worksheetViewport.rows} aria-colcount={worksheetViewport.columns}
        aria-multiselectable="true" aria-activedescendant={cellId(selection.anchor)} tabIndex={0}
        aria-describedby={id + '-help'} onBlur={() => { exitOnTab.current = false; }}
        onScroll={handleScroll} onKeyDown={keyDown} onPointerDown={pointerDown} onPointerMove={pointerMove}
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
        <div aria-hidden="true" className="worksheet-spacer" style={{ width: scrollSpacer(columns.total, left), height: scrollSpacer(rows.total, top) }} />
        <div className="worksheet-projection">
          {(sortProjectedRows ? [...projected].sort(([a], [b]) => a - b) : [...projected]).filter(([, indexes]) => indexes.length > 0).map(([row, indexes]) => {
            let selectedColumn = -1;
            let rowAreas: readonly Area[] | undefined;
            if (areas.length === 1) {
              const area = areas[0]!;
              if (row >= area.top && row <= area.bottom) {
                selectedColumn = area.left === area.right ? area.left : -2;
                if (selectedColumn === -2) rowAreas = areas;
              }
            } else if (areas.some(area => row >= area.top && row <= area.bottom)) {
              selectedColumn = -2;
              rowAreas = areas;
            }
            return <ProjectedRow key={row} row={row} columns={indexes} sheet={sheet} showFormulas={showFormulas}
              separator={separator} staleCells={staleCells} calculatingCells={calculatingCells} idPrefix={id}
              selectedColumn={selectedColumn} areas={rowAreas} />;
          })}
        </div>
      </div>
      <p className="worksheet-projection" id={id + '-help'}>
        Use arrow keys or Tab to move between cells. Press F2 to edit. Press Escape then Tab to leave the worksheet.
      </p>
      <canvas ref={canvas} aria-hidden="true" className="worksheet-canvas" style={{ width: view.width, height: view.height }} />
      <div className="worksheet-figure-layer" role="group" aria-label="Python Figure placements">
        {(props.figures ?? []).map((item, index) => {
          const isDraft = figureDraft?.placementId === item.placement.id;
          const displayPlacement = isDraft
            ? figurePlacementFromRect(item.placement.id, sheet.id, figureDraft.box, rows, columns, scale)
            : item.placement;
          const bounds = figureScreenRect(displayPlacement, rows, columns, scale, view.x, view.y, left, top);
          if (!bounds) return null;
          const image = item.placement.renderedFigure && item.placement.renderedSourceId
            ? item.placement.renderedFigure : item.figure;
          const selected = selectedFigureId === item.placement.id;
          const label = item.label ?? `Figure placement ${index + 1}`;
          return <div key={item.placement.id} className={'figure-placement' + (selected ? ' selected' : '') +
            (props.rangePicking ? ' range-picking' : '')} role="group" aria-label={label}
            style={{ transform: `translate3d(${bounds.x}px, ${bounds.y}px, 0)`, width: bounds.width, height: bounds.height }}
            onPointerDown={event => {
              if (props.rangePicking || (event.target as HTMLElement).closest('[data-figure-delete]')) return;
              const edge = (event.target as HTMLElement).closest<HTMLElement>('[data-figure-edge]')?.dataset.figureEdge as FigureEdge | undefined;
              beginFigureDrag(event, edge ? 'resize' : 'move', item.ownerId, item.figure, item.placement, edge);
            }}
            onPointerMove={updateFigureDraft} onPointerUp={finishFigureDrag}
            onPointerCancel={cancelFigureDrag} onLostPointerCapture={cancelFigureDrag}>
            <img src={`data:image/png;base64,${image.pngBase64}`} alt={label} draggable={false} />
            {selected && !props.rangePicking && <>
              <button type="button" className="figure-delete" data-figure-delete="true"
                aria-label="Delete figure placement" title="Delete figure placement"
                onPointerDown={event => event.stopPropagation()}
                onClick={event => { event.stopPropagation(); props.onFigureDelete?.(item.placement.id); setSelectedFigureId(undefined); }}>×</button>
              {figureEdges.map(edge => <button key={edge} type="button" className={'figure-handle ' + edge}
                data-figure-edge={edge} aria-label={'Resize figure ' + edge.replace('-', ' ')}
                title={'Resize figure ' + edge.replace('-', ' ')} />)}
            </>}
          </div>;
        })}
        {(() => {
          const pending = props.pendingFigure;
          if (!pending) return null;
          const cursorBox = figureCursor?.ownerId === pending.ownerId
            ? { x: figureCursor.x, y: figureCursor.y, width: 480,
              height: 480 * pending.figure.height / Math.max(1, pending.figure.width) }
            : undefined;
          const activeBox = figureDraft?.ownerId === pending.ownerId && !figureDraft.placementId
            ? figureDraft.box : cursorBox;
          const draftPlacement = activeBox && figurePlacementFromRect('figure-preview', sheet.id, activeBox, rows, columns, scale);
          const preview = draftPlacement && figureScreenRect(draftPlacement, rows, columns, scale, view.x, view.y, left, top);
          return <>
            {preview && <div aria-hidden="true" className="figure-placement-preview" style={{
              transform: `translate3d(${preview.x}px, ${preview.y}px, 0)`, width: preview.width, height: preview.height,
            }}><img src={`data:image/png;base64,${pending.figure.pngBase64}`} alt="" /></div>}
            <div className="figure-placement-capture" aria-label="Place Python Figure on worksheet"
              style={{ left, top, width: bodyWidth, height: bodyHeight }}
              onPointerDown={event => {
                if (event.button !== 0 || !event.isPrimary) return;
                beginFigureDrag(event, 'place', pending.ownerId, pending.figure);
              }}
              onPointerMove={event => {
                if (figureDrag.current) updateFigureDraft(event);
                else setFigureCursor({ ownerId: pending.ownerId, ...figurePoint(event.clientX, event.clientY) });
              }}
              onPointerUp={finishFigureDrag} onPointerCancel={cancelFigureDrag}
              onLostPointerCapture={cancelFigureDrag} onPointerLeave={() => { if (!figureDrag.current) setFigureCursor(undefined); }} />
          </>;
        })()}
      </div>
      {guide && <div ref={guideRef} aria-hidden="true" className={'dimension-guide ' + guide.axis}
        style={{ transform: guide.axis === 'row' ? `translate3d(0, ${guide.offset}px, 0)` : `translate3d(${guide.offset}px, 0, 0)` }} />}
      {editRow && editColumn && <div className="inline-editor-layer" style={{
        left, top, width: bodyWidth, height: bodyHeight,
      }}>
        <div className="inline-cell-editor" style={{
          left: editorLeft < bodyWidth ? Math.min(editorLeft, bodyWidth - editorWidth) : editorLeft,
          top: editorTop < bodyHeight ? Math.min(editorTop, bodyHeight - editorHeight) : editorTop,
          width: editorWidth, height: editorHeight, fontSize: Math.max(10, 14 * scale),
        }}>{props.editor?.input}</div>
      </div>}
      {!props.rangePicking && filterControls().map(control => <button
        key={control.row + '-' + control.column}
        type="button"
        className={'filter-button' + (control.active ? ' filter-active' : '')}
        aria-label={'Filter ' + control.label}
        title={'Filter ' + control.label}
        style={{ transform: `translate3d(${control.left}px, ${control.top}px, 0)`, width: control.width, height: control.height }}
        onClick={event => {
          const bounds = event.currentTarget.getBoundingClientRect();
          props.onFilter?.(control.row, control.column, bounds.left, bounds.bottom);
        }}
      >
        <svg width={12 * scale} height={12 * scale} viewBox="0 0 16 16" aria-hidden="true">
          {control.sort && <path d={control.sort.ascending ? 'M2 13 V3 M.5 6 L2 4.5 L3.5 6' : 'M2 3 V13 M.5 10 L2 11.5 L3.5 10'}
            fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" />}
          <path d={control.sort ? 'M6 3 H15 V4.5 L11 8 V12.5 H8 V8 L6 4.5 Z' : 'M5 3 H15 V4.5 L11 8 V12.5 H8 V8 L5 4.5 Z'}
            fill={control.active ? 'var(--accent)' : 'currentColor'} stroke="none" />
        </svg>
      </button>)}
      {!props.rangePicking && props.tables?.filter(t => !t.area.invalid && t.area.sheetId === sheet.id).map(table => {
        const lastR = lastRow(table.area);
        const lastC = lastColumn(table.area);
        const bounds = cellBounds(lastR, lastC);
        if (!bounds) return null;
        if (bounds.right <= left + 9 || bounds.right > view.width ||
            bounds.bottom <= top + 9 || bounds.bottom > view.height) return null;
        return (
          <button
            key={table.id}
            type="button"
            className="table-resize-handle"
            aria-label={'Resize table ' + table.name}
            title={'Resize table ' + table.name}
            style={{
              transform: `translate3d(${bounds.right - 9}px, ${bounds.bottom - 9}px, 0)`,
            }}
            onPointerDown={e => onStartResizeHandle(e, table)}
            onKeyDown={e => onHandleKeyDown(e, table)}
          >
            <svg width="9" height="9" viewBox="0 0 9 9">
              <polygon points="0,9 9,0 9,9" fill="var(--accent)" />
            </svg>
          </button>
        );
      })}
      {!props.rangePicking && (() => {
        const activeTable = props.tables?.find(t =>
          !t.area.invalid && t.area.sheetId === sheet.id && tableDesign(t).totalRow &&
          lastRow(t.area) === selection.anchor.row &&
          selection.anchor.column >= t.area.column && selection.anchor.column <= lastColumn(t.area)
        );
        if (!activeTable) return null;
        const bounds = cellBounds(selection.anchor.row, selection.anchor.column);
        if (!bounds) return null;
        const size = Math.min(17 * scale, bounds.height);
        const x = bounds.right + size <= view.width ? bounds.right : bounds.left - size;
        const y = bounds.top + (bounds.height - size) / 2;
        if (x < left || x + size > view.width || y < top || y + size > view.height) return null;
        const cellDoc = findCell(sheet, selection.anchor);
        const currentCode = totalFunction(cellDoc);
        return (
          <TotalRowDropdown
            currentCode={currentCode}
            position={{ x, y, size }}
            onSelect={code => props.onTotalFunction?.(activeTable.id, selection.anchor.column, code)}
          />
        );
      })()}
    </div>
  );
}
