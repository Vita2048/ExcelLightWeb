import { useEffect, useId, useRef, useState, type KeyboardEvent } from 'react';
import { nativeFunctions, type NativeFunction } from '../../calculation/formula-functions';
import type { CalculationMode } from '../../calculation/native-scheduler';
import type { EditAction } from '../editing/useWorksheetEditing';
import type { CellDocument, TableDesign, WorkbookName } from '../../core/workbook';
import type { NumberFormatChange } from '../../application/number-format-command';
import { RibbonFlyout, RibbonGlyph } from './RibbonFlyout';
import { AlignmentRibbon } from '../formatting/AlignmentRibbon';
import type { AlignmentAction, MergeMode } from '../../application/alignment-commands';
import { BorderRibbon, type BorderRibbonProps } from '../formatting/BorderRibbon';
import { FontRibbon } from '../formatting/FontRibbon';
import type { FormatTab } from '../formatting/FormatCellsDialog';
import type { AppearancePatch } from '../../core/formats/appearance';
import { NumberRibbon } from '../formatting/NumberRibbon';
import { StructureRibbon, type StructureAction } from '../structure/StructureControls';
import { TableDesignRibbon } from '../tables/TableDesignRibbon';
import { DataSortFilterRibbon, SortFilterHomeMenu, type FilterCommands } from '../data/DataRibbon';

const families = ['File', 'Home', 'Insert', 'Formulas', 'Data', 'View'] as const;
type Family = typeof families[number];

export interface SheetView {
  readonly headings: boolean;
  readonly gridlines: boolean;
  readonly formulaBar: boolean;
  readonly showFormulas: boolean;
  readonly zoom: number;
}

interface RibbonProps {
  readonly inert?: boolean;
  readonly fileBackstageOpen: boolean;
  readonly onFileBackstage: () => void;
  readonly onAlign: (action: AlignmentAction) => void;
  readonly onMerge: (mode: MergeMode) => void;
  readonly merged: boolean;
  readonly view: SheetView;
  readonly onView: (patch: Partial<SheetView>) => void;
  readonly onZoomToSelection: () => void;
  readonly onZoomDialog: () => void;
  readonly frozen: boolean;
  readonly onFreeze: (mode: 'panes' | 'topRow' | 'firstColumn' | 'none') => void;
  readonly onEdit: (action: EditAction) => void;
  readonly cell: CellDocument | undefined;
  readonly borders: BorderRibbonProps;
  readonly onAppearance: (patch: AppearancePatch, focus?: boolean) => void;
  readonly onAppearanceError: (error: string, focus: boolean) => void;
  readonly onFormatDialog: (tab: FormatTab) => void;
  readonly painter: 'single' | 'locked' | undefined;
  readonly onPainter: (locked: boolean) => void;
  readonly onNumberFormat: (change: NumberFormatChange) => void;
  readonly onStructure: (action: StructureAction) => void;
  readonly calculationMode: CalculationMode;
  readonly onCalculationMode: (mode: CalculationMode) => void;
  readonly onCalculate: (scope: 'sheet' | 'workbook') => void;
  readonly onInsertFunction: (name: NativeFunction) => void;
  readonly activeTable?: WorkbookName | undefined;
  readonly onTableRename?: (name: string) => void;
  readonly onTableResize?: () => void;
  readonly onTableChangeDesign?: (patch: Partial<TableDesign>) => void;
  readonly onTableClick?: () => void;
  readonly onNameManagerClick?: () => void;
  readonly onDefineNameClick?: () => void;
  readonly filter: FilterCommands;
  readonly onSort: (ascending: boolean, dialog: boolean) => void;
  readonly onToggleFilter: () => void;
  readonly onClearFilter: () => void;
  readonly onReapplyFilter: () => void;
  readonly onFind: () => void;
  readonly onReplace: () => void;
}

function EditMenu({ label, choices, onEdit }: {
  label: string;
  choices: readonly (readonly [string, EditAction])[];
  onEdit: (action: EditAction) => void;
}) {
  return <RibbonFlyout label={label} caption={label === 'Paste options' || label === 'AutoSum options' ? '' : label}
    {...((label === 'Fill' || label === 'Clear') && { icon: label })}
    items={choices.map(([name, action]) => ({ label: name, icon: action.kind === 'paste' ?
      action.mode === 'all' ? 'paste' : action.mode === 'values' ? 'pasteValues' : 'pasteFormulas' :
      name === 'Sum' ? 'AutoSum' : name, select: () => onEdit(action) }))} />;
}

function AutoSum({ onEdit }: { onEdit: RibbonProps['onEdit'] }) {
  return <div className="ribbon-split">
    <button className="ribbon-command" type="button" onClick={() => onEdit({ kind: 'autosum', aggregate: 'SUM' })}><RibbonGlyph name="AutoSum" />AutoSum</button>
    <EditMenu label="AutoSum options" onEdit={onEdit} choices={[
      ['Sum', { kind: 'autosum', aggregate: 'SUM' }], ['Average', { kind: 'autosum', aggregate: 'AVERAGE' }],
      ['Count Numbers', { kind: 'autosum', aggregate: 'COUNT' }], ['Max', { kind: 'autosum', aggregate: 'MAX' }],
      ['Min', { kind: 'autosum', aggregate: 'MIN' }],
    ]} />
  </div>;
}

export function Ribbon({ inert, fileBackstageOpen, onFileBackstage, onAlign, onMerge, merged, view, onView, onZoomToSelection, onZoomDialog, frozen, onFreeze, onEdit, onStructure, calculationMode,
  onCalculationMode, onCalculate, onInsertFunction, cell, onNumberFormat, onAppearance, onAppearanceError, onFormatDialog, painter, onPainter, borders,
  activeTable, onTableRename, onTableResize, onTableChangeDesign, onTableClick, onNameManagerClick, onDefineNameClick,
  filter, onSort, onToggleFilter, onClearFilter, onReapplyFilter, onFind, onReplace }: RibbonProps) {
  const [active, setActive] = useState<string>('Home');
  const [trackedTableId, setTrackedTableId] = useState<string | undefined>(undefined);
  const [knownTables, setKnownTables] = useState<ReadonlySet<string>>(() => new Set());
  const lastNonTableTab = useRef<string>('Home');
  const tableTabWasActive = useRef<Map<string, boolean>>(new Map());
  const id = useId();
  const tabs = useRef(new Map<string, HTMLButtonElement>());
  const currentId = activeTable?.id;
  // Switch the contextual tab during render. An effect paints one frame with every panel hidden, and the grid jumps.
  if (currentId !== trackedTableId) {
    const priorId = trackedTableId;
    if (priorId !== undefined) tableTabWasActive.current.set(priorId, active === 'Table Design');
    let nextActive = active;
    if (currentId !== undefined) {
      const isNew = !knownTables.has(currentId);
      if (isNew) {
        const nextKnown = new Set(knownTables);
        nextKnown.add(currentId);
        setKnownTables(nextKnown);
      }
      if (isNew || tableTabWasActive.current.get(currentId)) nextActive = 'Table Design';
    } else if (active === 'Table Design') {
      nextActive = lastNonTableTab.current;
    }
    if (nextActive !== active) setActive(nextActive);
    setTrackedTableId(currentId);
  }

  const tabList = activeTable ? [...families, 'Table Design'] : families;

  useEffect(() => {
    if (!fileBackstageOpen && active === 'File') setActive('Home');
  }, [active, fileBackstageOpen]);

  function selectTab(name: string) {
    if (name === 'File') {
      setActive('File');
      onFileBackstage();
      return;
    }
    if (name !== 'Table Design') {
      lastNonTableTab.current = name;
    }
    if (activeTable) {
      tableTabWasActive.current.set(activeTable.id, name === 'Table Design');
    }
    setActive(name);
  }

  function keyDown(event: KeyboardEvent<HTMLButtonElement>, index: number) {
    let next: string;
    switch (event.key) {
      case 'ArrowLeft': next = tabList[(index + tabList.length - 1) % tabList.length]!; break;
      case 'ArrowRight': next = tabList[(index + 1) % tabList.length]!; break;
      case 'Home': next = tabList[0]!; break;
      case 'End': next = tabList[tabList.length - 1]!; break;
      default: return;
    }
    event.preventDefault();
    selectTab(next);
    tabs.current.get(next)?.focus();
  }

  return (
    <div className="ribbon" inert={inert}>
      <div role="tablist" aria-label="Ribbon" className="ribbon-tabs">
        {tabList.map((name, index) => (
          <button type="button" key={name} role="tab" id={id + '-tab-' + name}
            aria-selected={active === name} aria-controls={id + '-panel-' + name} tabIndex={active === name ? 0 : -1}
            ref={node => { if (node) tabs.current.set(name, node); else tabs.current.delete(name); }}
            onClick={() => selectTab(name)} onKeyDown={event => keyDown(event, index)}>{name}</button>
        ))}
      </div>
      {tabList.map(name => (
        <div key={name} className="ribbon-panel" role="tabpanel" hidden={active !== name}
          id={id + '-panel-' + name} aria-labelledby={id + '-tab-' + name}>
          {name === 'Table Design' && activeTable ? (
            <TableDesignRibbon
              table={activeTable}
              onRename={onTableRename ?? (() => {})}
              onResize={onTableResize ?? (() => {})}
              onChangeDesign={onTableChangeDesign ?? (() => {})}
            />
          ) : name === 'Insert' ? (
            <fieldset>
              <legend>Tables</legend>
              <button type="button" className="ribbon-command ribbon-large" onClick={onTableClick}>
                <RibbonGlyph name="Table" size={32} />Table
              </button>
            </fieldset>
          ) : name === 'View' ? <>
            <fieldset>
              <legend>Show</legend><div className="ribbon-stack view-checks">
              <label><input type="checkbox" checked={view.formulaBar}
                onChange={event => onView({ formulaBar: event.currentTarget.checked })} /> Formula Bar</label>
              <label><input type="checkbox" checked={view.gridlines}
                onChange={event => onView({ gridlines: event.currentTarget.checked })} /> Gridlines</label>
              <label><input type="checkbox" checked={view.headings}
                onChange={event => onView({ headings: event.currentTarget.checked })} /> Headings</label>
              <label><input type="checkbox" checked={view.showFormulas}
                onChange={event => onView({ showFormulas: event.currentTarget.checked })} /> Show Formulas</label>
              </div>
            </fieldset>
            <fieldset>
              <legend>Zoom</legend>
              <button type="button" className="ribbon-command ribbon-large" onClick={onZoomDialog}><RibbonGlyph name="Zoom" size={30} />Zoom...</button>
              <button type="button" className="ribbon-command ribbon-large" onClick={() => onView({ zoom: 100 })}><RibbonGlyph name="100%" size={30} />100%</button>
              <button type="button" className="ribbon-command ribbon-large" onClick={onZoomToSelection}><RibbonGlyph name="Zoom to Selection" size={30} />Zoom to Selection</button>
            </fieldset>
            <fieldset>
              <legend>Window</legend>
              <RibbonFlyout label="Freeze Panes" icon="Freeze Panes" large items={[
                frozen ? { label: 'Unfreeze Panes', icon: 'Freeze Panes', select: () => onFreeze('none') } :
                  { label: 'Freeze Panes', icon: 'Freeze Panes', select: () => onFreeze('panes') },
                { label: 'Freeze Top Row', icon: 'Freeze Panes', select: () => onFreeze('topRow') },
                { label: 'Freeze First Column', icon: 'Freeze Panes', select: () => onFreeze('firstColumn') },
              ]} />
            </fieldset>
          </> : name === 'Home' ? <>
            <fieldset className="clipboard-ribbon">
              <legend>Clipboard</legend>
              <div className="ribbon-stack">
                <button type="button" className="ribbon-command ribbon-large" onClick={() => onEdit({ kind: 'paste', mode: 'all' })}>
                  <RibbonGlyph name="paste" size={28} />Paste
                </button>
                <EditMenu label="Paste options" onEdit={onEdit} choices={[
                  ['Paste', { kind: 'paste', mode: 'all' }], ['Paste Values', { kind: 'paste', mode: 'values' }],
                  ['Paste Formulas', { kind: 'paste', mode: 'formulas' }],
                ]} />
              </div>
              <div className="ribbon-stack">
                <button type="button" className="ribbon-command" onClick={() => onEdit({ kind: 'cut' })}><RibbonGlyph name="cut" />Cut</button>
                <button type="button" className="ribbon-command" onClick={() => onEdit({ kind: 'copy' })}><RibbonGlyph name="copy" />Copy</button>
                <div className="ribbon-split">
                  <button type="button" className="ribbon-command" title="Double-click to keep painting" aria-pressed={painter !== undefined}
                    onClick={event => onPainter(event.detail > 1)}><RibbonGlyph name="painter" />Format Painter</button>
                  <button type="button" aria-label="Lock Format Painter" title="Lock Format Painter" aria-pressed={painter === 'locked'}
                    onClick={() => onPainter(true)}><RibbonGlyph name="lock" size={12} /></button>
                </div>
              </div>
            </fieldset>
            <FontRibbon style={cell?.style} onApply={onAppearance} onInvalid={onAppearanceError} borders={<BorderRibbon {...borders} />}
              onDialog={() => onFormatDialog('Font')} />
            <AlignmentRibbon style={cell?.style} merged={merged} onAlign={onAlign} onMerge={onMerge} onDialog={() => onFormatDialog('Alignment')} />
            <NumberRibbon cell={cell} onFormat={onNumberFormat} onDialog={() => onStructure('format-cells')} />
            <StructureRibbon onAction={onStructure} />
            <fieldset>
              <legend>Editing</legend><div className="ribbon-stack">
              <AutoSum onEdit={onEdit} />
              <EditMenu label="Fill" onEdit={onEdit} choices={[
                ['Down', { kind: 'fill', direction: 'down' }], ['Right', { kind: 'fill', direction: 'right' }],
                ['Up', { kind: 'fill', direction: 'up' }], ['Left', { kind: 'fill', direction: 'left' }],
              ]} />
              <EditMenu label="Clear" onEdit={onEdit} choices={[
                ['Clear All', { kind: 'clear', mode: 'all' }], ['Clear Formats', { kind: 'clear', mode: 'formats' }],
                ['Clear Contents', { kind: 'clear', mode: 'contents' }],
              ]} />
              </div>
              <SortFilterHomeMenu filter={filter} onSort={onSort} onToggleFilter={onToggleFilter}
                onClearFilter={onClearFilter} onReapplyFilter={onReapplyFilter} />
              <RibbonFlyout label="Find & Select" caption={<>Find &amp;<br />Select</>} icon="Find & Select" large items={[
                { label: 'Find...', icon: 'Find...', select: onFind },
                { label: 'Replace...', icon: 'Replace...', select: onReplace },
              ]} />
            </fieldset>
          </> : name === 'Formulas' ? <>
            <fieldset>
              <legend>Function Library</legend>
              <RibbonFlyout label="Insert Function" icon="Insert Function" large items={nativeFunctions.map(name => ({
                label: name + '(…)', select: () => onInsertFunction(name),
              }))} />
              <AutoSum onEdit={onEdit} />
            </fieldset>
            <fieldset>
              <legend>Defined Names</legend>
              <button type="button" className="ribbon-command ribbon-large" onClick={onNameManagerClick}>
                <RibbonGlyph name="Name Manager" size={32} />Name Manager
              </button>
              <RibbonFlyout
                label="Define Name"
                icon="Define Name"
                items={[{ label: 'Define Name...', icon: 'Define Name', select: onDefineNameClick ?? (() => {}) }]}
              />
            </fieldset>
            <fieldset>
              <legend>Formula Auditing</legend>
              <button type="button" className="ribbon-command" aria-pressed={view.showFormulas}
                onClick={() => onView({ showFormulas: !view.showFormulas })}><RibbonGlyph name="Insert Function" />Show Formulas</button>
            </fieldset>
            <fieldset>
              <legend>Calculation</legend>
              <RibbonFlyout label="Calculation Options" icon="Calculation Options" large items={(['automatic', 'manual'] as const).map(mode => ({
                label: mode === 'manual' ? 'Manual' : 'Automatic', checked: calculationMode === mode, select: () => onCalculationMode(mode),
              }))} />
              <div className="ribbon-stack">
                <button type="button" className="ribbon-command" onClick={() => onCalculate('workbook')}><RibbonGlyph name="Calculate Now" />Calculate Now</button>
                <button type="button" className="ribbon-command" onClick={() => onCalculate('sheet')}><RibbonGlyph name="Calculate Sheet" />Calculate Sheet</button>
              </div>
            </fieldset>
          </> : name === 'Data' ? <DataSortFilterRibbon filter={filter} onSort={onSort} onToggleFilter={onToggleFilter}
            onClearFilter={onClearFilter} onReapplyFilter={onReapplyFilter} /> : null}
        </div>
      ))}
    </div>
  );
}
