import { useEffect, useRef, useState, type KeyboardEvent, type PointerEvent } from 'react';
import type { ChartDocument, ChartPlacement } from '../../core/chart';
import type { DocumentId } from '../../core/workbook';
import type { ResolvedChartData } from '../../charts/chart-data';
import {
  chartLogicalRect,
  chartPlacementFromRect,
  chartScreenRect,
  logicalPoint,
  type FigureRect,
} from '../../grid/figure-geometry';
import { clamp, type GridAxis } from '../../grid/grid-geometry';
import type { Area } from '../../grid/grid-selection';
import { ChartCanvas, downloadChartPng, downloadChartSvg } from './ChartCanvas';

export interface WorksheetChartItem {
  readonly chart: ChartDocument;
  readonly data: ResolvedChartData;
}

export interface ChartRangeHandleConfig {
  readonly area: Area;
  readonly color: string;
  readonly onResizeArea: (nextArea: Area) => void;
}

type ChartEdge = 'left' | 'top' | 'right' | 'bottom' | 'top-left' | 'top-right' | 'bottom-right' | 'bottom-left';
const chartEdges: readonly ChartEdge[] = [
  'top-left', 'top', 'top-right', 'right', 'bottom-right', 'bottom', 'bottom-left', 'left',
];

interface ChartDragState {
  readonly pointerId: number;
  readonly chartId: DocumentId;
  readonly mode: 'move' | 'resize';
  readonly edge?: ChartEdge;
  readonly start: { readonly x: number; readonly y: number };
  readonly initial: FigureRect;
}

interface RangeDragState {
  readonly pointerId: number;
  readonly startArea: Area;
  readonly currentArea: Area;
}

export function ChartOverlay({
  charts,
  selectedChartId,
  rangePicking = false,
  rows,
  columns,
  scale,
  scrollX,
  scrollY,
  headerLeft,
  headerTop,
  scrollContainer,
  scrollOrigin,
  cellBounds,
  pointAt,
  chartRangeHandle,
  onSelectChart,
  onChartPlacement,
  onChartDelete,
  onFocusGrid,
}: {
  readonly charts: readonly WorksheetChartItem[];
  readonly selectedChartId?: DocumentId | undefined;
  readonly rangePicking?: boolean;
  readonly rows: GridAxis;
  readonly columns: GridAxis;
  readonly scale: number;
  readonly scrollX: number;
  readonly scrollY: number;
  readonly headerLeft: number;
  readonly headerTop: number;
  readonly scrollContainer: () => HTMLDivElement | null;
  readonly scrollOrigin: (node: HTMLDivElement) => { readonly x: number; readonly y: number };
  readonly cellBounds: (row: number, column: number) => { readonly left: number; readonly top: number; readonly right: number; readonly bottom: number } | undefined;
  readonly pointAt: (x: number, y: number) => { readonly row: number; readonly column: number } | undefined;
  readonly chartRangeHandle?: ChartRangeHandleConfig | undefined;
  readonly onSelectChart?: ((chartId: DocumentId | undefined) => void) | undefined;
  readonly onChartPlacement?: ((chartId: DocumentId, placement: Partial<ChartPlacement>) => void) | undefined;
  readonly onChartDelete?: ((chartId: DocumentId) => void) | undefined;
  readonly onFocusGrid: () => void;
}) {
  const dragRef = useRef<ChartDragState | null>(null);
  const rangeDragRef = useRef<RangeDragState | null>(null);
  const svgRefs = useRef(new Map<DocumentId, SVGSVGElement | null>());
  const [draft, setDraft] = useState<{ readonly chartId: DocumentId; readonly box: FigureRect } | undefined>(undefined);
  const [rangePreview, setRangePreview] = useState<Area | undefined>(undefined);
  const [contextMenu, setContextMenu] = useState<{
    readonly chartId: DocumentId;
    readonly chartName: string;
    readonly width: number;
    readonly height: number;
    readonly x: number;
    readonly y: number;
  } | undefined>(undefined);

  useEffect(() => {
    if (!contextMenu) return;
    const dismiss = () => setContextMenu(undefined);
    window.addEventListener('pointerdown', dismiss);
    window.addEventListener('blur', dismiss);
    return () => {
      window.removeEventListener('pointerdown', dismiss);
      window.removeEventListener('blur', dismiss);
    };
  }, [contextMenu]);

  function logicalFromPointer(clientX: number, clientY: number) {
    const node = scrollContainer();
    if (!node) return { x: 0, y: 0 };
    const bounds = node.getBoundingClientRect();
    const origin = scrollOrigin(node);
    return {
      x: logicalPoint(clientX - bounds.left, headerLeft, columns, origin.x, scale),
      y: logicalPoint(clientY - bounds.top, headerTop, rows, origin.y, scale),
    };
  }

  function boxForPoint(active: ChartDragState, pt: { readonly x: number; readonly y: number }): FigureRect {
    const dx = pt.x - active.start.x;
    const dy = pt.y - active.start.y;
    const maxW = columns.total / scale;
    const maxH = rows.total / scale;
    if (active.mode === 'move') {
      return {
        ...active.initial,
        x: clamp(active.initial.x + dx, 0, Math.max(0, maxW - active.initial.width)),
        y: clamp(active.initial.y + dy, 0, Math.max(0, maxH - active.initial.height)),
      };
    }
    const edge = active.edge!;
    let x = active.initial.x;
    let y = active.initial.y;
    let right = x + active.initial.width;
    let bottom = y + active.initial.height;
    if (edge.includes('left')) x = clamp(x + dx, 0, right - 120);
    if (edge.includes('right')) right = clamp(right + dx, x + 120, Math.min(maxW, x + 4096));
    if (edge.includes('top')) y = clamp(y + dy, 0, bottom - 90);
    if (edge.includes('bottom')) bottom = clamp(bottom + dy, y + 90, Math.min(maxH, y + 4096));
    return { x, y, width: right - x, height: bottom - y };
  }

  function beginDrag(
    event: PointerEvent<HTMLDivElement>,
    chart: ChartDocument,
    mode: 'move' | 'resize',
    edge?: ChartEdge,
  ) {
    if (event.button !== 0 || !event.isPrimary || rangePicking) return;
    event.preventDefault();
    event.stopPropagation();
    setContextMenu(undefined);
    const initial = chartLogicalRect(chart.placement, rows, columns, scale);
    if (!initial) return;
    const start = logicalFromPointer(event.clientX, event.clientY);
    dragRef.current = {
      pointerId: event.pointerId,
      chartId: chart.id,
      mode,
      ...(edge && { edge }),
      start,
      initial,
    };
    onSelectChart?.(chart.id);
    setDraft({ chartId: chart.id, box: initial });
    event.currentTarget.setPointerCapture(event.pointerId);
    event.currentTarget.focus({ preventScroll: true });
  }

  function updateDrag(event: PointerEvent<HTMLDivElement>) {
    const active = dragRef.current;
    if (!active || active.pointerId !== event.pointerId) return;
    const box = boxForPoint(active, logicalFromPointer(event.clientX, event.clientY));
    setDraft({ chartId: active.chartId, box });
  }

  function finishDrag(event: PointerEvent<HTMLDivElement>) {
    const active = dragRef.current;
    if (!active || active.pointerId !== event.pointerId) return;
    const latest = boxForPoint(active, logicalFromPointer(event.clientX, event.clientY));
    const moved = active.mode !== 'move' || Math.hypot(latest.x - active.initial.x, latest.y - active.initial.y) >= 1;
    const resized = active.mode !== 'resize' || Math.hypot(latest.width - active.initial.width, latest.height - active.initial.height) >= 1;
    if (moved && resized) {
      try {
        const nextPlacement = chartPlacementFromRect(latest, rows, columns, scale);
        onChartPlacement?.(active.chartId, nextPlacement);
      } catch {
        // Ignore out-of-bounds drop
      }
    }
    dragRef.current = null;
    setDraft(undefined);
  }

  function cancelDrag(event?: PointerEvent<HTMLDivElement>) {
    if (event && dragRef.current?.pointerId !== event.pointerId) return;
    dragRef.current = null;
    setDraft(undefined);
  }

  function handleKeyDown(event: KeyboardEvent<HTMLDivElement>, chart: ChartDocument) {
    if (event.key === 'Delete' || event.key === 'Backspace') {
      event.preventDefault();
      event.stopPropagation();
      onChartDelete?.(chart.id);
      onFocusGrid();
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      onSelectChart?.(undefined);
      onFocusGrid();
      return;
    }
    const step = event.shiftKey ? 10 : 1;
    let dx = 0;
    let dy = 0;
    if (event.key === 'ArrowLeft') dx = -step;
    else if (event.key === 'ArrowRight') dx = step;
    else if (event.key === 'ArrowUp') dy = -step;
    else if (event.key === 'ArrowDown') dy = step;
    if (dx !== 0 || dy !== 0) {
      event.preventDefault();
      event.stopPropagation();
      const initial = chartLogicalRect(chart.placement, rows, columns, scale);
      if (!initial) return;
      try {
        const nextPlacement = chartPlacementFromRect(
          { ...initial, x: initial.x + dx, y: initial.y + dy },
          rows,
          columns,
          scale,
        );
        onChartPlacement?.(chart.id, nextPlacement);
      } catch {
        // Ignore boundary nudge
      }
    }
  }

  const sortedCharts = [...charts].sort(
    (a, b) => (a.chart.placement.zIndex ?? 0) - (b.chart.placement.zIndex ?? 0),
  );

  const rangeCorner = !rangePicking && chartRangeHandle
    ? (() => {
      const area = rangePreview ?? chartRangeHandle.area;
      const br = cellBounds(area.bottom, area.right);
      const tl = cellBounds(area.top, area.left);
      return br && tl ? { br, tl, area } : undefined;
    })()
    : undefined;

  return (
    <div className="worksheet-chart-layer" role="group" aria-label="Worksheet charts">
      {sortedCharts.map(({ chart, data }) => {
        const isDraft = draft?.chartId === chart.id;
        const displayPlacement: ChartPlacement = isDraft
          ? { ...chart.placement, ...chartPlacementFromRect(draft.box, rows, columns, scale) }
          : chart.placement;
        const bounds = chartScreenRect(
          displayPlacement,
          rows,
          columns,
          scale,
          scrollX,
          scrollY,
          headerLeft,
          headerTop,
        );
        if (!bounds) return null;
        const selected = selectedChartId === chart.id;
        return (
          <div
            key={chart.id}
            tabIndex={0}
            role="group"
            aria-label={chart.name}
            data-chart-id={chart.id}
            className={'chart-frame' + (selected ? ' selected' : '') + (rangePicking ? ' range-picking' : '')}
            style={{
              transform: `translate3d(${bounds.x}px, ${bounds.y}px, 0)`,
              width: bounds.width,
              height: bounds.height,
            }}
            onPointerDown={event => {
              if (rangePicking) return;
              const edge = (event.target as HTMLElement).closest<HTMLElement>('[data-chart-edge]')?.dataset.chartEdge as ChartEdge | undefined;
              beginDrag(event, chart, edge ? 'resize' : 'move', edge);
            }}
            onPointerMove={updateDrag}
            onPointerUp={finishDrag}
            onPointerCancel={cancelDrag}
            onLostPointerCapture={cancelDrag}
            onKeyDown={event => handleKeyDown(event, chart)}
            onContextMenu={event => {
              if (rangePicking) return;
              event.preventDefault();
              event.stopPropagation();
              onSelectChart?.(chart.id);
              setContextMenu({
                chartId: chart.id,
                chartName: chart.name,
                width: Math.round(chart.placement.width),
                height: Math.round(chart.placement.height),
                x: event.clientX,
                y: event.clientY,
              });
            }}
          >
            <ChartCanvas
              chart={chart}
              data={data}
              width={displayPlacement.width}
              height={displayPlacement.height}
              svgRef={node => {
                if (node) svgRefs.current.set(chart.id, node);
                else svgRefs.current.delete(chart.id);
              }}
              interactive={!isDraft && !rangePicking}
            />
            {selected && !rangePicking && chartEdges.map(edge => (
              <button
                key={edge}
                type="button"
                className={'figure-handle ' + edge}
                data-chart-edge={edge}
                aria-label={`Resize ${chart.name} ${edge.replace('-', ' ')}`}
                title={`Resize ${edge.replace('-', ' ')}`}
              />
            ))}
          </div>
        );
      })}

      {rangeCorner && chartRangeHandle && (
        <>
          {rangePreview && (
            <div
              aria-hidden="true"
              className="chart-range-preview"
              style={{
                left: rangeCorner.tl.left,
                top: rangeCorner.tl.top,
                width: Math.max(8, rangeCorner.br.right - rangeCorner.tl.left),
                height: Math.max(8, rangeCorner.br.bottom - rangeCorner.tl.top),
                borderColor: chartRangeHandle.color,
              }}
            />
          )}
          <button
            type="button"
            className="chart-range-handle"
            aria-label="Resize chart data range"
            title="Drag to resize chart data range"
            style={{
              transform: `translate3d(${rangeCorner.br.right}px, ${rangeCorner.br.bottom}px, 0)`,
              borderColor: chartRangeHandle.color,
            }}
            onPointerDown={event => {
              if (event.button !== 0 || !event.isPrimary) return;
              event.preventDefault();
              event.stopPropagation();
              rangeDragRef.current = {
                pointerId: event.pointerId,
                startArea: chartRangeHandle.area,
                currentArea: chartRangeHandle.area,
              };
              setRangePreview(chartRangeHandle.area);
              event.currentTarget.setPointerCapture(event.pointerId);
            }}
            onPointerMove={event => {
              const active = rangeDragRef.current;
              const node = scrollContainer();
              if (!active || active.pointerId !== event.pointerId || !node) return;
              const bounds = node.getBoundingClientRect();
              const pt = pointAt(event.clientX - bounds.left, event.clientY - bounds.top);
              if (!pt) return;
              const nextArea: Area = {
                top: active.startArea.top,
                left: active.startArea.left,
                bottom: Math.max(active.startArea.top, pt.row),
                right: Math.max(active.startArea.left, pt.column),
              };
              rangeDragRef.current = { ...active, currentArea: nextArea };
              setRangePreview(nextArea);
            }}
            onPointerUp={event => {
              const active = rangeDragRef.current;
              if (!active || active.pointerId !== event.pointerId) return;
              rangeDragRef.current = null;
              setRangePreview(undefined);
              if (
                active.currentArea.bottom !== active.startArea.bottom ||
                active.currentArea.right !== active.startArea.right
              ) {
                chartRangeHandle.onResizeArea(active.currentArea);
              }
            }}
            onPointerCancel={() => {
              rangeDragRef.current = null;
              setRangePreview(undefined);
            }}
          />
        </>
      )}

      {contextMenu && (
        <div
          className="ribbon-flyout chart-context-menu"
          role="menu"
          aria-label="Chart context menu"
          style={{ left: contextMenu.x, top: contextMenu.y }}
          onPointerDown={event => event.stopPropagation()}
        >
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              const svg = svgRefs.current.get(contextMenu.chartId);
              if (svg) downloadChartPng(svg, contextMenu.width, contextMenu.height, contextMenu.chartName);
              setContextMenu(undefined);
            }}
          >
            Save as Picture (PNG)...
          </button>
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              const svg = svgRefs.current.get(contextMenu.chartId);
              if (svg) downloadChartSvg(svg, contextMenu.chartName);
              setContextMenu(undefined);
            }}
          >
            Save as Picture (SVG)...
          </button>
          <div className="ribbon-flyout-divider" role="none">
            <button
              type="button"
              role="menuitem"
              onClick={() => {
                onChartDelete?.(contextMenu.chartId);
                setContextMenu(undefined);
              }}
            >
              Delete Chart
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
