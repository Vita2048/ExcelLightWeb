import { useEffect, useLayoutEffect, useRef, useState, type FormEvent } from 'react';
import type { FilterValue } from '../../application/data-commands';

export interface FilterDropdownProps {
  readonly anchor: { readonly x: number; readonly y: number };
  readonly header: string;
  readonly textColumn: boolean;
  readonly values: readonly FilterValue[];
  readonly selected: readonly string[];
  readonly hasColumnFilter: boolean;
  readonly onApply: (selected: readonly string[]) => void;
  readonly onClearColumn: () => void;
  readonly onSort: (ascending: boolean) => void;
  readonly onCancel: () => void;
}

/** The desktop header funnel checklist: sort, clear, search and a value checklist. */
export function FilterDropdown({ anchor, header, textColumn, values, selected, hasColumnFilter,
  onApply, onClearColumn, onSort, onCancel }: FilterDropdownProps) {
  const panel = useRef<HTMLDivElement>(null);
  const search = useRef<HTMLInputElement>(null);
  const selectAll = useRef<HTMLInputElement>(null);
  const [chosen, setChosen] = useState(() => new Set(selected));
  const [query, setQuery] = useState('');
  const shown = query.trim().length === 0 ? values :
    values.filter(value => value.label.toLowerCase().includes(query.trim().toLowerCase()));
  const shownSelected = shown.filter(value => chosen.has(value.key)).length;
  const selectedCount = values.filter(value => value.key.length > 0 && chosen.has(value.key)).length;
  const allChecked = shown.length > 0 && shownSelected === shown.length;
  const someChecked = shownSelected > 0 && !allChecked;

  useLayoutEffect(() => {
    const node = panel.current!;
    node.style.left = Math.max(8, Math.min(anchor.x, innerWidth - node.offsetWidth - 8)) + 'px';
    node.style.top = Math.max(8, Math.min(anchor.y, innerHeight - node.offsetHeight - 8)) + 'px';
    search.current?.focus({ preventScroll: true });
  }, [anchor.x, anchor.y]);

  useEffect(() => {
    if (selectAll.current) selectAll.current.indeterminate = someChecked;
  }, [someChecked]);

  useEffect(() => {
    function dismiss(event: PointerEvent) {
      if (!panel.current?.contains(event.target as Node)) onCancel();
    }
    document.addEventListener('pointerdown', dismiss);
    return () => document.removeEventListener('pointerdown', dismiss);
  }, [onCancel]);

  function toggle(key: string, checked: boolean) {
    setChosen(current => {
      const next = new Set(current);
      if (checked) next.add(key); else next.delete(key);
      return next;
    });
  }

  function toggleAll() {
    const turnOn = shownSelected < shown.length;
    setChosen(current => {
      const next = new Set(current);
      for (const value of shown) {
        if (turnOn) next.add(value.key); else next.delete(value.key);
      }
      return next;
    });
  }

  function submit(event: FormEvent) {
    event.preventDefault();
    if (selectedCount > 0) onApply([...chosen]);
  }

  return <div ref={panel} className="filter-dropdown" role="dialog" aria-label={'Filter ' + header}
    onKeyDown={event => {
      if (event.key === 'Escape') {
        event.preventDefault();
        onCancel();
      }
    }}>
    <form onSubmit={submit}>
      <button type="button" className="filter-action" onClick={() => onSort(true)}>
        {textColumn ? 'Sort A to Z' : 'Sort Smallest to Largest'}
      </button>
      <button type="button" className="filter-action" onClick={() => onSort(false)}>
        {textColumn ? 'Sort Z to A' : 'Sort Largest to Smallest'}
      </button>
      <hr />
      <button type="button" className="filter-action" disabled={!hasColumnFilter} onClick={onClearColumn}>
        Clear Filter From &quot;{header}&quot;
      </button>
      <hr />
      <input ref={search} type="search" className="filter-search" placeholder="Search" aria-label="Search filter values"
        value={query} onChange={event => setQuery(event.currentTarget.value)} />
      <div className="filter-values">
        <label className="filter-select-all">
          <input ref={selectAll} type="checkbox" disabled={shown.length === 0}
            checked={allChecked} onChange={toggleAll} /> (Select All)
        </label>
        <div className="filter-list">
          {shown.map(value => <label key={value.key}>
            <input type="checkbox" checked={chosen.has(value.key)}
              onChange={event => toggle(value.key, event.currentTarget.checked)} /> {value.label}
          </label>)}
        </div>
      </div>
      <div className="filter-actions">
        <button type="submit" disabled={selectedCount === 0}>OK</button>
        <button type="button" onClick={onCancel}>Cancel</button>
      </div>
    </form>
  </div>;
}
