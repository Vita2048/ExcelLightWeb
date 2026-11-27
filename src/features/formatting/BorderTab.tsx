import { type MouseEvent, useLayoutEffect, useRef, useState } from 'react';
import type { BorderEditorState } from '../../application/border-commands';
import { borderPen, borderStyles, draftEdges, type BorderDraft, type BorderStyle, type DraftEdge } from '../../core/formats/borders';
import type { CellBorder } from '../../core/workbook';
import { equalDto } from '../../core/workbook-delta';
import { drawBorder } from '../../grid/border-rendering';
import { ColorControl } from './ColorControl';

const labels: Record<DraftEdge, string> = { top: 'Top', bottom: 'Bottom', left: 'Left', right: 'Right',
  horizontal: 'Inside horizontal', vertical: 'Inside vertical', diagonalUp: 'Diagonal up', diagonalDown: 'Diagonal down' };

const presets = {
  None: [], Outline: ['top', 'right', 'bottom', 'left'], Inside: ['horizontal', 'vertical'],
} as const satisfies Record<string, readonly DraftEdge[]>;

const paths: Record<DraftEdge, string> = {
  top: 'M3 3H19', bottom: 'M3 19H19', left: 'M3 3V19', right: 'M19 3V19',
  horizontal: 'M3 11H19', vertical: 'M11 3V19', diagonalUp: 'M3 19L19 3', diagonalDown: 'M3 3L19 19',
};

function BorderIcon({ active }: { active: readonly DraftEdge[] }) {
  return <svg viewBox="0 0 22 22" fill="none" stroke="currentColor" aria-hidden="true">
    <path d={Object.values(paths).slice(0, 6).join(' ')} strokeWidth=".55" strokeDasharray=".65 .95" />
    {active.length > 0 && <path d={active.map(edge => paths[edge]).join(' ')} strokeWidth=".7" />}
  </svg>;
}

export function BorderSample({ pen, width = 66 }: { pen: CellBorder | null; width?: number }) {
  const ref = useRef<HTMLCanvasElement>(null);
  useLayoutEffect(() => {
    const context = ref.current!.getContext('2d')!;
    const media = matchMedia('(forced-colors: active)');
    function paint() {
      context.clearRect(0, 0, width, 14);
      if (pen) drawBorder(context, pen, 1, 7.5, width - 1, 7.5, 1,
        media.matches ? 'CanvasText' : getComputedStyle(ref.current!).color, media.matches);
    }
    paint();
    media.addEventListener('change', paint);
    return () => media.removeEventListener('change', paint);
  }, [pen?.style, pen?.rgb, width]);
  return <canvas ref={ref} width={width} height={14} aria-hidden="true" />;
}

export function BorderTab({ initial, draft, onChange }: {
  initial: BorderEditorState; draft: BorderDraft; onChange: (draft: BorderDraft) => void;
}) {
  const [style, setStyle] = useState<BorderStyle>('thin');
  const [color, setColor] = useState<string>();
  const edges = { ...initial.edges, ...draft };
  const diagram = useRef<HTMLCanvasElement>(null);
  const pen = borderPen(style, color);
  const available = (edge: DraftEdge) => (edge !== 'horizontal' || initial.rows > 1) && (edge !== 'vertical' || initial.columns > 1);

  function change(entries: [DraftEdge, CellBorder | null][]) {
    const next = { ...draft };
    for (const [key, pen] of entries) {
      delete next[key];
      next[key] = pen;
    }
    onChange(next);
  }

  function updatePen(kind: BorderStyle, rgb: string | undefined, recolor: boolean) {
    setStyle(kind);
    setColor(rgb);
    change(draftEdges.flatMap(key => {
      const edge = edges[key];
      if (!available(key) || !edge || edge === 'mixed') return [];
      return [[key, recolor ? borderPen(edge.style as BorderStyle, rgb) : borderPen(kind, edge.rgb)]];
    }));
  }

  function toggle(key: DraftEdge) {
    const value = equalDto(edges[key], pen) ? null : pen;
    const entries: [DraftEdge, CellBorder | null][] = [[key, value]];
    if (value && (key === 'diagonalUp' || key === 'diagonalDown')) {
      const other = key === 'diagonalUp' ? 'diagonalDown' : 'diagonalUp';
      if (edges[other] && edges[other] !== 'mixed') entries.push([other, value]);
    }
    change(entries);
  }

  useLayoutEffect(() => {
    const canvas = diagram.current!, context = canvas.getContext('2d')!;
    const media = matchMedia('(forced-colors: active)');
    function paint() {
      const { width, height } = canvas.getBoundingClientRect();
      if (!width || !height) return;
      const scale = window.devicePixelRatio;
      canvas.width = Math.round(width * scale);
      canvas.height = Math.round(height * scale);
      context.scale(scale, scale);
      const ink = media.matches ? 'CanvasText' : getComputedStyle(canvas).color;
      const guide = media.matches ? 'CanvasText' : '#8a8a8a';
      const l = 6.5, r = width - 6.5, t = 6.5, b = height - 6.5;
      const mx = (l + r) / 2, my = (t + b) / 2;
      const points = {
        top: [l, t, r, t], horizontal: [l, my, r, my], bottom: [l, b, r, b],
        left: [l, t, l, b], vertical: [mx, t, mx, b], right: [r, t, r, b],
      } as const;
      const tick = (x: number, y: number, x2: number, y2: number) =>
        drawBorder(context, { style: 'thin' }, x, y, x2, y2, 1, guide, media.matches);
      for (const x of initial.columns > 1 ? [l, mx, r] : [l, r]) {
        tick(x, t, x, t + 4);
        tick(x, b - 4, x, b);
      }
      for (const y of initial.rows > 1 ? [t, my, b] : [t, b]) {
        tick(l, y, l + 4, y);
        tick(r - 4, y, r, y);
      }
      for (const key of draftEdges) {
        const value = edges[key];
        if (!value || !available(key)) continue;
        const stroke = value === 'mixed' ? { style: 'dotted' } : value;
        const color = value === 'mixed' ? guide : ink;
        if (key !== 'diagonalUp' && key !== 'diagonalDown') {
          const [x, y, x2, y2] = points[key];
          drawBorder(context, stroke, x, y, x2, y2, 1, color, media.matches);
          continue;
        }
        const w = (r - l) / initial.columns, h = (b - t) / initial.rows;
        const up = key === 'diagonalUp';
        for (let row = 0; row < initial.rows; row++) {
          for (let column = 0; column < initial.columns; column++) {
            const x = l + column * w, y = t + row * h;
            drawBorder(context, stroke, x, y + (up ? h : 0), x + w, y + (up ? 0 : h), 1, color, media.matches);
          }
        }
      }
    }
    paint();
    const observer = new ResizeObserver(paint);
    observer.observe(canvas);
    media.addEventListener('change', paint);
    return () => { observer.disconnect(); media.removeEventListener('change', paint); };
  });

  function clickPreview(event: MouseEvent<HTMLCanvasElement>) {
    const bounds = event.currentTarget.getBoundingClientRect();
    const x = event.clientX - bounds.left, y = event.clientY - bounds.top;
    const { width, height } = bounds;
    if (y < 14) return toggle('top');
    if (y > height - 14) return toggle('bottom');
    if (x < 14) return toggle('left');
    if (x > width - 14) return toggle('right');
    if (available('horizontal') && Math.abs(y - height / 2) < 8) return toggle('horizontal');
    if (available('vertical') && Math.abs(x - width / 2) < 8) return toggle('vertical');
    const cellWidth = (width - 12) / initial.columns, cellHeight = (height - 12) / initial.rows;
    const localX = ((x - 6) % cellWidth) / cellWidth, localY = ((y - 6) % cellHeight) / cellHeight;
    if (Math.abs(localY - localX) < .12) toggle('diagonalDown');
    else if (Math.abs(localY + localX - 1) < .12) toggle('diagonalUp');
  }

  return <div className="border-tab">
    <div className="border-dialog-columns">
      <fieldset><legend>Line</legend>
        <div>Style:</div>
        <div className="border-style-picker" role="group" aria-label="Border line style">
          {Object.entries(borderStyles).map(([key, name]) => <button type="button" key={key}
            aria-label={name} title={name} aria-pressed={style === key} onClick={() => updatePen(key as BorderStyle, color, false)}>
            {key === 'none' ? name : <BorderSample pen={borderPen(key as BorderStyle, color)} />}
          </button>)}
        </div>
        <div className="border-color-label">Color:</div>
        <ColorControl label="Border Color" value={color} display="value" automatic onChange={rgb => updatePen(style, rgb, true)} />
      </fieldset>
      <div className="border-dialog-controls">
        <fieldset><legend>Presets</legend><div className="border-presets">
          {Object.entries(presets).map(([name, sides]) => <button type="button" key={name}
            disabled={name === 'Inside' && initial.rows === 1 && initial.columns === 1}
            onClick={() => change((name === 'None' ? draftEdges : sides).filter(available)
              .map(key => [key, name === 'None' ? null : pen]))}>
            <span className="border-preset-icon"><BorderIcon active={sides} /></span>{name}
          </button>)}
        </div></fieldset>
        <fieldset><legend>Border</legend><div className="border-diagram">
          <div className="border-preview">
            <canvas ref={diagram} aria-hidden="true" onClick={clickPreview} />
            <div className="border-preview-text" aria-hidden="true"
              style={{ gridTemplateColumns: 'repeat(' + initial.columns + ', 1fr)' }}>
              {Array.from({ length: initial.rows * initial.columns }, (_, index) => <span key={index}>Text</span>)}
            </div>
          </div>
          {draftEdges.map(key => <button type="button" key={key} className={'border-side-' + key}
            aria-label={labels[key]} title={labels[key] + ' border'} disabled={!available(key)}
            aria-pressed={edges[key] === 'mixed' ? 'mixed' : !!edges[key]} onClick={() => toggle(key)}>
            <BorderIcon active={[key]} />
          </button>)}
        </div></fieldset>
      </div>
    </div>
    <p>The selected border style can be applied by clicking the presets, preview diagram or the buttons above.</p>
  </div>;
}
