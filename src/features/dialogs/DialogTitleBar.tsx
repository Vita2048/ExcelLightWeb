import { useRef, type ReactNode } from 'react';

export function DialogTitleBar({ children, spaced = false }: { children: ReactNode; spaced?: boolean }) {
  const drag = useRef<{ pointer: number; x: number; y: number } | undefined>(undefined);

  return <div className={'format-dialog-heading' + (spaced ? ' dialog-title-spaced' : '')}
    onPointerDown={event => {
      const target = event.target;
      if (!event.isPrimary || event.button !== 0 || !(target instanceof Element) ||
          target.closest('button, a, input, select, textarea, [contenteditable]')) return;
      const dialog = event.currentTarget.closest('dialog');
      if (!dialog) return;
      const rect = dialog.getBoundingClientRect();
      drag.current = { pointer: event.pointerId, x: event.clientX - rect.left, y: event.clientY - rect.top };
      event.currentTarget.setPointerCapture(event.pointerId);
      event.preventDefault();
    }}
    onPointerMove={event => {
      const placement = drag.current;
      if (!placement || placement.pointer !== event.pointerId) return;
      const dialog = event.currentTarget.closest('dialog');
      if (!dialog) return;
      const viewport = dialog.ownerDocument.documentElement;
      const left = Math.max(0, Math.min(event.clientX - placement.x, viewport.clientWidth - dialog.offsetWidth));
      const top = Math.max(0, Math.min(event.clientY - placement.y, viewport.clientHeight - dialog.offsetHeight));
      Object.assign(dialog.style, { inset: 'auto', margin: '0', left: left + 'px', top: top + 'px' });
    }}
    onLostPointerCapture={() => { drag.current = undefined; }}
  >{children}</div>;
}
