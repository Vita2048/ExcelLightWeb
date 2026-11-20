import { RibbonFlyout, RibbonGlyph } from '../shell/RibbonFlyout';

export interface FilterCommands {
  readonly active: boolean;
  readonly clearable: boolean;
  readonly reapplyable: boolean;
}
export interface DataRibbonProps {
  readonly filter: FilterCommands;
  readonly onSort: (ascending: boolean, dialog: boolean) => void;
  readonly onToggleFilter: () => void;
  readonly onClearFilter: () => void;
  readonly onReapplyFilter: () => void;
}

/** Data → Sort & Filter, in the desktop group order. */
export function DataSortFilterRibbon({ filter, onSort, onToggleFilter, onClearFilter, onReapplyFilter }: DataRibbonProps) {
  return <fieldset>
    <legend>Sort &amp; Filter</legend>
    <div className="ribbon-stack">
      <button type="button" className="ribbon-command" onClick={() => onSort(true, false)}>
        <RibbonGlyph name="Sort A to Z" size={18} />A to Z
      </button>
      <button type="button" className="ribbon-command" onClick={() => onSort(false, false)}>
        <RibbonGlyph name="Sort Z to A" size={18} />Z to A
      </button>
    </div>
    <button type="button" className="ribbon-command ribbon-large" onClick={() => onSort(true, true)}>
      <RibbonGlyph name="Sort" size={32} />Sort
    </button>
    <button type="button" className="ribbon-command ribbon-large" aria-pressed={filter.active}
      onClick={onToggleFilter}>
      <RibbonGlyph name="Filter" size={32} />Filter
    </button>
    <div className="ribbon-stack">
      <button type="button" className="ribbon-command" aria-label="Clear filter" disabled={!filter.clearable}
        onClick={onClearFilter}>
        <RibbonGlyph name="Clear Filter" size={18} />Clear
      </button>
      <button type="button" className="ribbon-command" aria-label="Reapply filter" disabled={!filter.reapplyable}
        onClick={onReapplyFilter}>
        <RibbonGlyph name="Reapply" size={18} />Reapply
      </button>
    </div>
  </fieldset>;
}

/** Home → Editing → Sort & Filter, a large menu with the desktop entries. */
export function SortFilterHomeMenu({ filter, onSort, onToggleFilter, onClearFilter, onReapplyFilter }: DataRibbonProps) {
  return <RibbonFlyout label="Sort & Filter" caption={<>Sort &amp;<br />Filter</>} icon="Sort & Filter" large items={[
    { label: 'Sort A to Z', icon: 'Sort A to Z', select: () => onSort(true, false) },
    { label: 'Sort Z to A', icon: 'Sort Z to A', select: () => onSort(false, false) },
    { label: 'Custom Sort...', icon: 'Custom Sort...', select: () => onSort(true, true) },
    { label: 'Filter', icon: 'Filter', divider: true, select: onToggleFilter },
    { label: 'Clear', icon: 'Clear Filter', disabled: !filter.clearable, select: onClearFilter },
    { label: 'Reapply', icon: 'Reapply', disabled: !filter.reapplyable, select: onReapplyFilter },
  ]} />;
}
