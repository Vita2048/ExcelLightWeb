import { memo, useId, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { borderModes, borderPen, borderPresets, borderStyles, type BorderAction, type BorderMode, type BorderStyle } from '../../core/formats/borders';
import { BorderSample } from './BorderTab';
import { RibbonGlyph, useRibbonPopover } from '../shell/RibbonFlyout';
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

export const BorderRibbon = memo(function BorderRibbon({ action, style, color, mode, onAction, onStyle, onColor, onDialog }: BorderRibbonProps) {
  const menu = useRef<HTMLDivElement>(null);
  const trigger = useRef<HTMLButtonElement>(null);
  const id = useId();
  const styleMenu = useRef<HTMLDivElement>(null);
  const styleTrigger = useRef<HTMLButtonElement>(null);
  useRibbonPopover(trigger, menu);
  useRibbonPopover(styleTrigger, styleMenu, true);
  const [recent, setRecent] = useState<string[]>([]);
  const label = { ...borderPresets, ...borderModes }[action];
  function close() { if (menu.current?.matches(':popover-open')) menu.current.hidePopover(); }
  return <div className="border-ribbon ribbon-split">
    <button type="button" aria-label={'Repeat ' + label} title={label} aria-pressed={mode !== undefined} onClick={() => onAction(action)}><RibbonGlyph name={'border:' + action} /></button>
    <button type="button" ref={trigger} className="ribbon-command" aria-label="Borders menu" popoverTarget={id}><span className="ribbon-chevron" aria-hidden="true" /></button>
    {createPortal(<div className="border-menu" id={id} ref={menu} popover="auto" role="group" aria-label="Borders actions">
        <div className="ribbon-flyout-heading">Borders</div>
        {Object.entries(borderPresets).map(([action, name], index) => <div key={action}
          className={index === 4 || index === 8 ? 'ribbon-flyout-divider' : undefined}>
          <button type="button" onClick={() => { close(); onAction(action as BorderAction); }}>
            <RibbonGlyph name={'border:' + action} size={22} />{name}</button>
        </div>)}
        <div className="ribbon-flyout-divider"><div className="ribbon-flyout-heading">Draw Borders</div>
          {Object.entries(borderModes).map(([action, name]) => <button key={action} type="button"
            onClick={() => { close(); onAction(action as BorderAction); }}><RibbonGlyph name={'border:' + action} size={22} />{name}</button>)}
          <ColorControl label="Border Line Color" value={color} automatic dropdown icon="border:color" recentColors={recent} onChange={(rgb, custom) => {
            if (rgb && custom) setRecent(colors => [rgb, ...colors.filter(color => color !== rgb)].slice(0, 8));
            close();
            onColor(rgb);
          }} />
          <button type="button" ref={styleTrigger} popoverTarget={id + '-style'}><span className="ribbon-flyout-glyph" aria-hidden="true" />Line Style<span className="ribbon-submenu-chevron" aria-hidden="true" /></button>
          <div className="ribbon-flyout border-style-menu" popover="auto" ref={styleMenu} id={id + '-style'} role="group" aria-label="Line Style">
            {(['none', 'thin', 'hair', 'dotted', 'dashed', 'dashDot', 'dashDotDot', 'double', 'medium', 'mediumDashed',
              'mediumDashDot', 'mediumDashDotDot', 'slantDashDot', 'thick'] as const).map(key => <button type="button" key={key}
              aria-label={borderStyles[key]} title={borderStyles[key]} aria-pressed={style === key}
              onClick={() => { close(); onStyle(key); }}>
              {key === 'none' ? 'None' : <BorderSample pen={borderPen(key, color)} width={160} />}
            </button>)}
          </div>
          <div className="ribbon-flyout-divider"><button type="button" onClick={() => { close(); onDialog(); }}>
            <RibbonGlyph name="border:more" size={22} />More Borders…</button></div>
        </div>
      </div>, document.body)}
  </div>;
});
