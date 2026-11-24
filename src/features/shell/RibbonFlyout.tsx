import { useId, useLayoutEffect, useRef, useState, type KeyboardEvent, type ReactNode, type RefObject } from 'react';
import { createPortal } from 'react-dom';
import { desktopGlyphs } from './desktop-glyphs';

export function RibbonGlyph({ name, size = 20 }: { name: string; size?: number }) {
  const drawing = desktopGlyphs[name];
  return drawing ? <svg className="ribbon-glyph" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true"
    dangerouslySetInnerHTML={{ __html: drawing }} /> : null;
}

export function positionRibbonPopup(anchor: HTMLElement, panel: HTMLElement, beside = false) {
  const box = (!beside && (anchor.closest('.ribbon-split, .ribbon-font-input') as HTMLElement)) || anchor;
  const rect = box.getBoundingClientRect();
  const gap = 4;
  const width = panel.offsetWidth;
  const below = innerHeight - rect.bottom - gap;
  const above = rect.top - gap;
  const toRight = innerWidth - rect.right - gap;
  const preferredLeft = beside ? toRight >= width ? rect.right : rect.left - width : rect.left;
  const height = panel.offsetHeight;
  const preferredTop = beside ? rect.top : height > below && above > below ? rect.top - height - 2 : rect.bottom + 2;
  panel.style.left = Math.max(gap, Math.min(preferredLeft, innerWidth - width - gap)) + 'px';
  panel.style.top = Math.max(gap, Math.min(preferredTop, innerHeight - height - gap)) + 'px';
}

export function useRibbonPopover(anchor: RefObject<HTMLElement | null>, panel: RefObject<HTMLElement | null>, beside = false) {
  useLayoutEffect(() => {
    const popup = panel.current;
    if (!popup) return;
    function position() {
      const button = anchor.current;
      if (button && popup?.matches(':popover-open')) positionRibbonPopup(button, popup, beside);
    }
    popup.addEventListener('toggle', position);
    window.addEventListener('resize', position);
    document.addEventListener('scroll', position, true);
    return () => {
      popup.removeEventListener('toggle', position);
      window.removeEventListener('resize', position);
      document.removeEventListener('scroll', position, true);
    };
  });
}

export interface RibbonItem {
  readonly label: string;
  readonly icon?: string;
  readonly detail?: string;
  readonly group?: string;
  readonly divider?: boolean;
  readonly checked?: boolean;
  readonly disabled?: boolean;
  readonly select?: () => void;
  readonly children?: readonly RibbonItem[];
}

export function RibbonFlyout({ label, caption, icon, items, large = false }: {
  label: string; caption?: ReactNode; icon?: string; items: readonly RibbonItem[]; large?: boolean;
}) {
  const id = useId();
  const [open, setOpen] = useState(false);
  const [submenu, setSubmenu] = useState<{ item: RibbonItem; anchor: HTMLButtonElement }>();
  const trigger = useRef<HTMLButtonElement>(null);
  const menu = useRef<HTMLDivElement>(null);
  const childMenu = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    if (!open) return;
    const button = trigger.current!, panel = menu.current!;
    const position = () => positionRibbonPopup(button, panel);
    function dismiss(event: PointerEvent | FocusEvent) {
      const target = event.target as Node;
      if (!panel.contains(target) && !childMenu.current?.contains(target) && !button.contains(target)) {
        setOpen(false);
        setSubmenu(undefined);
      }
    }
    position();
    panel.querySelector('button')?.focus({ preventScroll: true });
    window.addEventListener('resize', position);
    document.addEventListener('scroll', position, true);
    document.addEventListener('pointerdown', dismiss);
    document.addEventListener('focusin', dismiss);
    return () => {
      window.removeEventListener('resize', position);
      document.removeEventListener('scroll', position, true);
      document.removeEventListener('pointerdown', dismiss);
      document.removeEventListener('focusin', dismiss);
    };
  }, [open]);

  useLayoutEffect(() => {
    if (!submenu) return;
    const panel = childMenu.current!;
    const position = () => positionRibbonPopup(submenu.anchor, panel, true);
    position();
    panel.querySelector('button')?.focus({ preventScroll: true });
    window.addEventListener('resize', position);
    document.addEventListener('scroll', position, true);
    return () => {
      window.removeEventListener('resize', position);
      document.removeEventListener('scroll', position, true);
    };
  }, [submenu]);

  function close() {
    setOpen(false);
    setSubmenu(undefined);
    trigger.current?.focus();
  }

  function enter(item: RibbonItem, anchor: HTMLButtonElement) {
    if (item.children) setSubmenu({ item, anchor });
    else {
      close();
      item.select?.();
    }
  }

  function handleMenu(event: KeyboardEvent<HTMLDivElement>, choices: readonly RibbonItem[], child = false) {
    if (child && submenu && (event.key === 'ArrowLeft' || event.key === 'Escape')) {
      event.preventDefault();
      submenu.anchor.focus();
      setSubmenu(undefined);
      return;
    }
    if (event.key === 'Escape') { event.preventDefault(); close(); return; }
    const buttons = [...event.currentTarget.querySelectorAll('button')];
    const current = buttons.indexOf(document.activeElement as HTMLButtonElement);
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 :
      event.key === 'ArrowDown' ? (current + 1) % buttons.length :
      event.key === 'ArrowUp' ? (current + buttons.length - 1) % buttons.length : undefined;
    if (event.key === 'ArrowRight' && choices[current]?.children && buttons[current]) {
      event.preventDefault();
      enter(choices[current], buttons[current]);
    }
    if (next !== undefined) { event.preventDefault(); buttons[next]?.focus(); }
    if (event.key === 'Tab') close();
  }

  function rows(choices: readonly RibbonItem[]) {
    return choices.map((item, index) => {
      const { label: text, icon: glyph, detail, group, checked } = item;
      const newGroup = group !== choices[index - 1]?.group;
      return <div key={text} role="none" className={index > 0 && (newGroup || item.divider) ? 'ribbon-flyout-divider' : undefined}>
        {group && newGroup && <div className="ribbon-flyout-heading" aria-hidden="true">{group}</div>}
        <button type="button" className={detail !== undefined ? 'ribbon-flyout-sample' : undefined}
          role={checked === undefined ? 'menuitem' : 'menuitemcheckbox'} aria-checked={checked} disabled={item.disabled}
          aria-haspopup={item.children ? 'menu' : undefined} aria-expanded={item.children ? submenu?.item === item : undefined}
          onClick={event => enter(item, event.currentTarget)}>
          <span className="ribbon-flyout-glyph" aria-hidden="true">{glyph ? <RibbonGlyph name={glyph} size={detail !== undefined ? 32 : 22} /> :
            checked && <span className="ribbon-flyout-check">✓</span>}</span>
          <span className="ribbon-flyout-label">{text}{detail && <small>{detail}</small>}</span>
          {item.children && <span className="ribbon-submenu-chevron" aria-hidden="true" />}
        </button>
      </div>;
    });
  }

  // The accessible name keeps the visible caption, so a changing value like the Number
  // Format sample still satisfies label-in-name while the command stays identifiable.
  const accessibleName = typeof caption === 'string' && caption !== '' && caption !== label ? label + ' ' + caption : label;
  return <>
    <button type="button" ref={trigger} className={'ribbon-command' + (large ? ' ribbon-large' : '')} aria-label={accessibleName}
      aria-haspopup="menu" aria-expanded={open} aria-controls={open ? id : undefined} title={label}
      onClick={() => { if (open) close(); else setOpen(true); }} onKeyDown={event => {
        if (event.key === 'ArrowDown') { event.preventDefault(); setOpen(true); }
        if (event.key === 'Escape' && open) { event.preventDefault(); close(); }
      }}>
      {icon && <RibbonGlyph name={icon} size={large ? 30 : 18} />}
      {caption ?? label}<span className="ribbon-chevron" aria-hidden="true" />
    </button>
    {open && createPortal(<>
      <div ref={menu} id={id} role="menu" aria-label={label} className="ribbon-flyout" onKeyDown={event => handleMenu(event, items)}>
        {rows(items)}
      </div>
      {submenu?.item.children && <div ref={childMenu} role="menu" aria-label={submenu.item.label} className="ribbon-flyout"
        onKeyDown={event => handleMenu(event, submenu.item.children!, true)}>{rows(submenu.item.children)}</div>}
    </>, document.body)}
  </>;
}
