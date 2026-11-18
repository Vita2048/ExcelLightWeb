import { useId, useRef, useState, type KeyboardEvent } from 'react';
import type { EditAction } from '../editing/useWorksheetEditing';
import { StructureRibbon, type StructureAction } from '../structure/StructureControls';

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
  readonly onEdit: (action: EditAction) => void;
  readonly onStructure: (action: StructureAction) => void;
}

const unavailable: Record<'File' | 'Insert' | 'Data', string> = {
  File: 'File commands are unavailable.',
  Insert: 'Insert commands are unavailable.',
  Data: 'Data commands are unavailable.',
};

function EditMenu({ label, choices, onEdit }: {
  label: string;
  choices: readonly (readonly [string, EditAction])[];
  onEdit: (action: EditAction) => void;
}) {
  return <label className="edit-menu">
    <span className="worksheet-projection">{label}</span>
    <select value="" onChange={event => {
      const item = choices[Number(event.currentTarget.value)];
      if (item) onEdit(item[1]);
    }}>
      <option value="" disabled>{label}</option>
      {choices.map(([name], index) => <option key={name} value={index}>{name}</option>)}
    </select>
  </label>;
}

function AutoSum({ onEdit }: { onEdit: RibbonProps['onEdit'] }) {
  return <>
    <button type="button" onClick={() => onEdit({ kind: 'autosum', aggregate: 'SUM' })}>AutoSum</button>
    <EditMenu label="AutoSum options" onEdit={onEdit} choices={[
      ['Sum', { kind: 'autosum', aggregate: 'SUM' }], ['Average', { kind: 'autosum', aggregate: 'AVERAGE' }],
      ['Count Numbers', { kind: 'autosum', aggregate: 'COUNT' }], ['Max', { kind: 'autosum', aggregate: 'MAX' }],
      ['Min', { kind: 'autosum', aggregate: 'MIN' }],
    ]} />
  </>;
}

export function Ribbon({ view, onView, onZoomToSelection, onEdit, onStructure }: RibbonProps) {
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
          tabIndex={name in unavailable ? 0 : undefined} id={id + '-panel-' + name} aria-labelledby={id + '-tab-' + name}>
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
          </> : name === 'Home' ? <>
            <fieldset>
              <legend>Clipboard</legend>
              <button type="button" onClick={() => onEdit({ kind: 'paste', mode: 'all' })}>Paste</button>
              <EditMenu label="Paste options" onEdit={onEdit} choices={[
                ['Paste Values', { kind: 'paste', mode: 'values' }],
                ['Paste Formulas', { kind: 'paste', mode: 'formulas' }],
              ]} />
              <button type="button" onClick={() => onEdit({ kind: 'cut' })}>Cut</button>
              <button type="button" onClick={() => onEdit({ kind: 'copy' })}>Copy</button>
            </fieldset>
            <StructureRibbon onAction={onStructure} />
            <fieldset>
              <legend>Editing</legend>
              <AutoSum onEdit={onEdit} />
              <EditMenu label="Fill" onEdit={onEdit} choices={[
                ['Down', { kind: 'fill', direction: 'down' }], ['Right', { kind: 'fill', direction: 'right' }],
                ['Up', { kind: 'fill', direction: 'up' }], ['Left', { kind: 'fill', direction: 'left' }],
              ]} />
              <EditMenu label="Clear" onEdit={onEdit} choices={[
                ['Clear All', { kind: 'clear', mode: 'all' }], ['Clear Formats', { kind: 'clear', mode: 'formats' }],
                ['Clear Contents', { kind: 'clear', mode: 'contents' }],
              ]} />
            </fieldset>
          </> : name === 'Formulas' ? <fieldset>
            <legend>Function Library</legend><AutoSum onEdit={onEdit} />
          </fieldset> : <p>{unavailable[name]}</p>}
        </div>
      ))}
    </div>
  );
}
