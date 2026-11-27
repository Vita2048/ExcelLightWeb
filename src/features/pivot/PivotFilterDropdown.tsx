import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState, type FormEvent } from 'react';
import { extractPivotRecords, formatPivotItemLabel } from '../../calculation/pivot-engine';
import type { PivotTableDocument } from '../../core/pivot';
import type { WorkbookDocument } from '../../core/workbook';

export interface PivotFilterDropdownProps {
  readonly anchor: { readonly x: number; readonly y: number };
  readonly kind: 'reportFilter' | 'rowLabel' | 'colLabel';
  readonly fieldNames: readonly string[];
  readonly pivot: PivotTableDocument;
  readonly document: WorkbookDocument;
  readonly onApplyFilter: (fieldName: string, hiddenItems: readonly string[]) => void;
  readonly onSort?: (fieldName: string, direction: 'asc' | 'desc') => void;
  readonly onMoreSortOptions?: (fieldName: string) => void;
  readonly onCancel: () => void;
}

export function PivotFilterDropdown({
  anchor,
  kind,
  fieldNames,
  pivot,
  document: workbookDoc,
  onApplyFilter,
  onSort,
  onMoreSortOptions,
  onCancel,
}: PivotFilterDropdownProps) {
  const panel = useRef<HTMLDivElement>(null);
  const searchRef = useRef<HTMLInputElement>(null);
  const selectAllRef = useRef<HTMLInputElement>(null);
  const fieldSelectId = useId();
  const multiCheckId = useId();

  const [selectedField, setSelectedField] = useState(fieldNames[0] ?? '');
  const [query, setQuery] = useState('');

  useEffect(() => {
    if (!fieldNames.includes(selectedField)) {
      setSelectedField(fieldNames[0] ?? '');
    }
  }, [fieldNames, selectedField]);

  const { distinctItems, currentHidden, hiddenKey, distinctKey } = useMemo(() => {
    const { records } = extractPivotRecords(workbookDoc, pivot.source);
    const items = [...new Set(records.map(rec => formatPivotItemLabel(rec[selectedField] ?? null)))];
    items.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
    const config = kind === 'reportFilter'
      ? pivot.filters.find(f => f.fieldName === selectedField)
      : kind === 'rowLabel'
        ? pivot.rows.find(r => r.fieldName === selectedField)
        : pivot.columns.find(c => c.fieldName === selectedField);
    const hidden = config?.hiddenItems ?? [];
    return {
      distinctItems: items,
      currentHidden: new Set(hidden),
      hiddenKey: [...hidden].sort().join('\u0000'),
      distinctKey: items.join('\u0000'),
    };
  }, [workbookDoc, pivot.source, pivot.filters, pivot.rows, pivot.columns, kind, selectedField]);

  const [multiSelect, setMultiSelect] = useState(() => {
    const vis = distinctItems.filter(item => !currentHidden.has(item));
    return kind !== 'reportFilter' || (vis.length > 1 && vis.length < distinctItems.length);
  });
  const [chosen, setChosen] = useState(() => new Set(distinctItems.filter(item => !currentHidden.has(item))));
  const [singleChoice, setSingleChoice] = useState<string>(() => {
    const vis = distinctItems.filter(item => !currentHidden.has(item));
    return vis.length === 1 && distinctItems.length > 1 ? vis[0]! : '(All)';
  });

  useEffect(() => {
    const vis = distinctItems.filter(item => !currentHidden.has(item));
    setChosen(new Set(vis));
    setSingleChoice(vis.length === 1 && distinctItems.length > 1 ? vis[0]! : '(All)');
    if (kind === 'reportFilter') {
      setMultiSelect(vis.length > 1 && vis.length < distinctItems.length);
    }
    setQuery('');
    // Reset only when the active field, kind, item universe, or stored hiddenItems changes.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedField, kind, hiddenKey, distinctKey]);

  useLayoutEffect(() => {
    const node = panel.current;
    if (!node) return;
    node.style.left = Math.max(8, Math.min(anchor.x, window.innerWidth - node.offsetWidth - 8)) + 'px';
    node.style.top = Math.max(8, Math.min(anchor.y, window.innerHeight - node.offsetHeight - 8)) + 'px';
    searchRef.current?.focus({ preventScroll: true });
  }, [anchor.x, anchor.y]);

  useEffect(() => {
    function dismiss(event: PointerEvent) {
      if (!panel.current?.contains(event.target as Node)) onCancel();
    }
    window.document.addEventListener('pointerdown', dismiss);
    return () => window.document.removeEventListener('pointerdown', dismiss);
  }, [onCancel]);

  const needle = query.trim().toLowerCase();
  const shown = needle.length === 0
    ? distinctItems
    : distinctItems.filter(item => item.toLowerCase().includes(needle));
  const shownSelected = shown.filter(item => chosen.has(item)).length;
  const allChecked = shown.length > 0 && shownSelected === shown.length;
  const someChecked = shownSelected > 0 && !allChecked;

  useEffect(() => {
    if (selectAllRef.current) selectAllRef.current.indeterminate = someChecked;
  }, [someChecked]);

  function toggleItem(item: string, checked: boolean) {
    setChosen(current => {
      const next = new Set(current);
      if (checked) next.add(item);
      else next.delete(item);
      return next;
    });
  }

  function toggleAll() {
    const turnOn = shownSelected < shown.length;
    setChosen(current => {
      const next = new Set(current);
      for (const item of shown) {
        if (turnOn) next.add(item);
        else next.delete(item);
      }
      return next;
    });
  }

  function handleSubmit(event: FormEvent) {
    event.preventDefault();
    if (!selectedField) return;
    if (kind === 'reportFilter' && !multiSelect) {
      const hidden = singleChoice === '(All)'
        ? []
        : distinctItems.filter(item => item !== singleChoice);
      onApplyFilter(selectedField, hidden);
      return;
    }
    if (chosen.size === 0) return;
    const hidden = distinctItems.filter(item => !chosen.has(item));
    onApplyFilter(selectedField, hidden);
  }

  const labelTitle = kind === 'reportFilter'
    ? `Filter ${selectedField}`
    : kind === 'rowLabel'
      ? `Filter Row Labels (${selectedField})`
      : `Filter Column Labels (${selectedField})`;

  return (
    <div
      ref={panel}
      className="filter-dropdown pivot-filter-dropdown"
      role="dialog"
      aria-label={labelTitle}
      onKeyDown={event => {
        if (event.key === 'Escape') {
          event.preventDefault();
          onCancel();
        }
      }}
    >
      <form onSubmit={handleSubmit}>
        {fieldNames.length > 1 && (
          <div className="pivot-filter-field-row">
            <label htmlFor={fieldSelectId}>Select field:</label>
            <select
              id={fieldSelectId}
              aria-label="Select field"
              value={selectedField}
              onChange={event => setSelectedField(event.currentTarget.value)}
            >
              {fieldNames.map(name => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </div>
        )}

        {kind !== 'reportFilter' && (
          <>
            <button type="button" className="filter-action" onClick={() => onSort?.(selectedField, 'asc')}>
              Sort A to Z
            </button>
            <button type="button" className="filter-action" onClick={() => onSort?.(selectedField, 'desc')}>
              Sort Z to A
            </button>
            <button type="button" className="filter-action" onClick={() => onMoreSortOptions?.(selectedField)}>
              More Sort Options...
            </button>
            <hr />
          </>
        )}

        <button
          type="button"
          className="filter-action"
          disabled={currentHidden.size === 0}
          onClick={() => onApplyFilter(selectedField, [])}
        >
          Clear Filter From &quot;{selectedField}&quot;
        </button>
        <hr />

        <input
          ref={searchRef}
          type="search"
          className="filter-search"
          placeholder="Search"
          aria-label="Search filter items"
          value={query}
          onChange={event => setQuery(event.currentTarget.value)}
        />

        <div className="filter-values">
          {kind === 'reportFilter' && !multiSelect ? (
            <div className="filter-list" role="listbox" aria-label={`${selectedField} items`}>
              <button
                type="button"
                role="option"
                aria-selected={singleChoice === '(All)'}
                className={'pivot-filter-single-option' + (singleChoice === '(All)' ? ' selected' : '')}
                onClick={() => setSingleChoice('(All)')}
              >
                (All)
              </button>
              {shown.map(item => (
                <button
                  key={item}
                  type="button"
                  role="option"
                  aria-selected={singleChoice === item}
                  className={'pivot-filter-single-option' + (singleChoice === item ? ' selected' : '')}
                  onClick={() => setSingleChoice(item)}
                >
                  {item}
                </button>
              ))}
            </div>
          ) : (
            <>
              <label className="filter-select-all">
                <input
                  ref={selectAllRef}
                  type="checkbox"
                  disabled={shown.length === 0}
                  checked={allChecked}
                  onChange={toggleAll}
                />{' '}
                (Select All)
              </label>
              <div className="filter-list">
                {shown.map(item => (
                  <label key={item}>
                    <input
                      type="checkbox"
                      checked={chosen.has(item)}
                      onChange={event => toggleItem(item, event.currentTarget.checked)}
                    />{' '}
                    {item}
                  </label>
                ))}
              </div>
            </>
          )}
        </div>

        {kind === 'reportFilter' && (
          <label className="pivot-filter-multi-check" htmlFor={multiCheckId}>
            <input
              id={multiCheckId}
              type="checkbox"
              checked={multiSelect}
              onChange={event => setMultiSelect(event.currentTarget.checked)}
            />
            <span>Select Multiple Items</span>
          </label>
        )}

        <div className="filter-actions">
          <button
            type="submit"
            disabled={(kind !== 'reportFilter' || multiSelect) && chosen.size === 0}
          >
            OK
          </button>
          <button type="button" onClick={onCancel}>Cancel</button>
        </div>
      </form>
    </div>
  );
}
