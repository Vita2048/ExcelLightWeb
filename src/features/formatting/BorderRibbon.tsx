import { useId, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { borderModes, borderPen, borderPresets, borderStyles, type BorderAction, type BorderMode, type BorderStyle } from '../../core/formats/borders';
import { BorderSample } from './BorderTab';
import { ColorControl } from './ColorControl';

export interface BorderRibbonProps {
  readonly action: BorderAction;
  readonly style: BorderStyle;
  readonly color: string | undefined;
  readonly mode: BorderMode | undefined;
  readonly onAction: (action: BorderAction) => void;
  readonly onStyle: (style: BorderStyle) => void;
  readonly onColor: (color?: string) => void;
  readonly onDialog: () => void;
}

export function BorderRibbon({ action, style, color, mode, onAction, onStyle, onColor, onDialog }: BorderRibbonProps) {
  const menu = useRef<HTMLDivElement>(null);
  const trigger = useRef<HTMLButtonElement>(null);
  const id = useId();
  const [recent, setRecent] = useState<string[]>([]);
  const label = { ...borderPresets, ...borderModes }[action];
  function close() { if (menu.current?.matches(':popover-open')) menu.current.hidePopover(); }
  return <fieldset className="border-ribbon">
    <legend>Borders</legend>
    <button type="button" aria-label={'Repeat ' + label} title={label} aria-pressed={mode !== undefined} onClick={() => onAction(action)}>▤</button>
    <button type="button" ref={trigger} aria-label="Borders menu" popoverTarget={id} onClick={() => {
      const box = trigger.current!.getBoundingClientRect();
      menu.current!.style.left = Math.max(4, Math.min(box.left, innerWidth - 278)) + 'px';
      menu.current!.style.top = Math.min(box.bottom + 4, innerHeight - 180) + 'px';
    }}>⌄</button>
    {createPortal(<div className="border-menu" id={id} ref={menu} popover="auto" role="group" aria-label="Borders actions">
        {[borderPresets, borderModes].map((choices, index) => <div key={index} className="border-menu-group">
          {Object.entries(choices).map(([action, name]) => <button key={action} type="button"
            onClick={() => { close(); onAction(action as BorderAction); }}>{name}</button>)}
        </div>)}
        <div className="border-menu-group">
          <ColorControl label="Border Line Color" value={color} automatic recentColors={recent} onChange={(rgb, custom) => {
            if (rgb && custom) setRecent(colors => [rgb, ...colors.filter(color => color !== rgb)].slice(0, 8));
            close();
            onColor(rgb);
          }} />
          <details>
            <summary>Line Style</summary>
            <div className="border-styles">
              {Object.entries(borderStyles).map(([key, name]) => <button type="button" key={key}
                aria-pressed={style === key} onClick={() => { close(); onStyle(key as BorderStyle); }}>
                <BorderSample pen={borderPen(key as BorderStyle, color)} />{name}
              </button>)}
            </div>
          </details>
          <button type="button" onClick={() => { close(); onDialog(); }}>More Borders…</button>
        </div>
      </div>, document.body)}
  </fieldset>;
}
