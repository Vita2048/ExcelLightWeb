import { useId, useRef, useState, type KeyboardEvent } from 'react';

const families = ['File', 'Home', 'Insert', 'Formulas', 'Data', 'View'] as const;
type Family = typeof families[number];

export interface SheetView {
  readonly headings: boolean;
  readonly gridlines: boolean;
  readonly formulaBar: boolean;
  readonly zoom: number;
}

interface RibbonProps {
  readonly view: SheetView;
  readonly onView: (patch: Partial<SheetView>) => void;
  readonly onZoomToSelection: () => void;
}

const unavailable: Record<Exclude<Family, 'View'>, string> = {
  File: 'File commands are unavailable.',
  Home: 'Cell editing is unavailable.',
  Insert: 'Insert commands are unavailable.',
  Formulas: 'Formula commands are unavailable.',
  Data: 'Data commands are unavailable.',
};

export function Ribbon({ view, onView, onZoomToSelection }: RibbonProps) {
  const [active, setActive] = useState<Family>('Home');
  const id = useId();
  const tabs = useRef(new Map<Family, HTMLButtonElement>());

  function keyDown(event: KeyboardEvent<HTMLButtonElement>, index: number) {
    let next: Family;
    switch (event.key) {
      case 'ArrowLeft': next = families[(index + families.length - 1) % families.length]!; break;
      case 'ArrowRight': next = families[(index + 1) % families.length]!; break;
      case 'Home': next = families[0]; break;
      case 'End': next = families[families.length - 1]!; break;
      default: return;
    }
    event.preventDefault();
    setActive(next);
    tabs.current.get(next)?.focus();
  }

  return (
    <div className="ribbon">
      <div role="tablist" aria-label="Ribbon" className="ribbon-tabs">
        {families.map((name, index) => (
          <button type="button" key={name} role="tab" id={id + '-tab-' + name}
            aria-selected={active === name} aria-controls={id + '-panel-' + name} tabIndex={active === name ? 0 : -1}
            ref={node => { if (node) tabs.current.set(name, node); else tabs.current.delete(name); }}
            onClick={() => setActive(name)} onKeyDown={event => keyDown(event, index)}>{name}</button>
        ))}
      </div>
      {families.map(name => (
        <div key={name} className="ribbon-panel" role="tabpanel" hidden={active !== name}
          tabIndex={name === 'View' ? undefined : 0} id={id + '-panel-' + name} aria-labelledby={id + '-tab-' + name}>
          {name === 'View' ? <>
            <fieldset>
              <legend>Show</legend>
              <label><input type="checkbox" checked={view.formulaBar}
                onChange={event => onView({ formulaBar: event.currentTarget.checked })} /> Formula Bar</label>
              <label><input type="checkbox" checked={view.gridlines}
                onChange={event => onView({ gridlines: event.currentTarget.checked })} /> Gridlines</label>
              <label><input type="checkbox" checked={view.headings}
                onChange={event => onView({ headings: event.currentTarget.checked })} /> Headings</label>
            </fieldset>
            <fieldset>
              <legend>Zoom</legend>
              <button type="button" onClick={() => onView({ zoom: 100 })}>100%</button>
              <button type="button" onClick={onZoomToSelection}>Zoom to Selection</button>
            </fieldset>
          </> : <p>{unavailable[name]}</p>}
        </div>
      ))}
    </div>
  );
}
