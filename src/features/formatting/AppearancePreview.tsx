import { useLayoutEffect, useRef } from 'react';
import type { CellStyle } from '../../core/workbook';
import { drawCellFill, drawCellText } from '../../grid/cell-appearance';

export function AppearancePreview({ style, label, text, width = 320, height = 88 }: {
  style: CellStyle;
  label?: string;
  text?: string;
  width?: number | 'full';
  height?: number;
}) {
  const ref = useRef<HTMLCanvasElement>(null);
  useLayoutEffect(() => {
    const canvas = ref.current!;
    const colors = matchMedia('(forced-colors: active)');
    function draw() {
      const context = canvas.getContext('2d')!;
      const width = canvas.clientWidth || canvas.width;
      const height = canvas.clientHeight || canvas.height;
      const ratio = devicePixelRatio || 1;
      canvas.width = Math.round(width * ratio);
      canvas.height = Math.round(height * ratio);
      context.setTransform(ratio, 0, 0, ratio, 0, 0);
      const css = getComputedStyle(canvas);
      if (colors.matches) {
        context.fillStyle = 'Canvas';
        context.fillRect(0, 0, width, height);
      } else drawCellFill(context, style, 0, 0, width, height, 1, css.backgroundColor);
      if (text) {
        context.fillStyle = colors.matches ? 'CanvasText' : style.textRgb ? '#' + style.textRgb : css.color;
        drawCellText(context, text, style, css.fontFamily, 1, width / 2, height / 2, 'center', height - 2);
      }
    }
    draw();
    const resize = new ResizeObserver(draw);
    resize.observe(canvas);
    colors.addEventListener('change', draw);
    return () => { resize.disconnect(); colors.removeEventListener('change', draw); };
  }, [style, text]);
  return <canvas ref={ref} className="appearance-preview" width={width === 'full' ? undefined : width} height={height}
    style={{ height, width: width === 'full' ? '100%' : width }} role={label ? 'img' : undefined} aria-label={label} aria-hidden={label ? undefined : true} />;
}
