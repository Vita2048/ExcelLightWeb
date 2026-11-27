import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react';
import { extractPivotRecords } from '../../calculation/pivot-engine';
import {
  defaultPivotValueLabel,
  type PivotAxisField,
  type PivotFilterField,
  type PivotTableDocument,
  type PivotValueField,
} from '../../core/pivot';
import type { IdFactory, WorkbookDocument } from '../../core/workbook';
import type { FieldSettingsRequest } from './FieldSettingsDialog';

type PivotArea = 'filters' | 'columns' | 'rows' | 'values';

interface OpenChipMenu {
  readonly key: string;
  readonly area: PivotArea;
  readonly index: number;
  readonly label: string;
  readonly entry: PivotAxisField | PivotValueField | PivotFilterField;
  readonly anchor: { readonly left: number; readonly right: number; readonly top: number; readonly bottom: number };
}

interface StagedLayout {
  readonly filters: readonly PivotFilterField[];
  readonly columns: readonly PivotAxisField[];
  readonly rows: readonly PivotAxisField[];
  readonly values: readonly PivotValueField[];
}

interface DragState {
  readonly kind: 'field' | 'chip';
  readonly fieldName: string;
  readonly label: string;
  readonly fromArea?: PivotArea;
  readonly fromIndex?: number;
  readonly startX: number;
  readonly startY: number;
  readonly x: number;
  readonly y: number;
  readonly active: boolean;
}

interface DropTarget {
  readonly area: PivotArea;
  readonly index: number;
}

function defaultIsNumericField(records: readonly Readonly<Record<string, unknown>>[], fieldName: string): boolean {
  let num = 0;
  let nonBlank = 0;
  for (const row of records) {
    const val = row[fieldName];
    if (val === null || val === undefined || val === '') continue;
    nonBlank++;
    if (typeof val === 'number') num++;
  }
  return nonBlank > 0 && num / nonBlank > 0.5;
}

function clampIndex(index: number, length: number): number {
  return Math.max(0, Math.min(length, index));
}

export function PivotTableFieldsPane({
  width,
  document,
  pivot,
  makeId,
  onUpdateLayout,
  onClose,
  onOpenSettings,
}: {
  readonly width: number;
  readonly document: WorkbookDocument;
  readonly pivot: PivotTableDocument;
  readonly makeId: IdFactory;
  readonly onUpdateLayout: (patch: StagedLayout) => void;
  readonly onClose: () => void;
  readonly onOpenSettings: (request: FieldSettingsRequest) => void;
}) {
  const searchId = useId();
  const deferId = useId();
  const paneRef = useRef<HTMLElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const [search, setSearch] = useState('');
  const [defer, setDefer] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [openMenu, setOpenMenu] = useState<OpenChipMenu>();
  const [staged, setStaged] = useState<StagedLayout>({
    filters: pivot.filters,
    columns: pivot.columns,
    rows: pivot.rows,
    values: pivot.values,
  });
  const [drag, setDrag] = useState<DragState>();
  const [dropTarget, setDropTarget] = useState<DropTarget>();

  const stagedRef = useRef(staged);
  stagedRef.current = staged;
  const deferRef = useRef(defer);
  deferRef.current = defer;
  const onUpdateLayoutRef = useRef(onUpdateLayout);
  onUpdateLayoutRef.current = onUpdateLayout;
  const dragRef = useRef<DragState | undefined>(undefined);
  const suppressClickRef = useRef(false);

  useEffect(() => {
    setStaged({
      filters: pivot.filters,
      columns: pivot.columns,
      rows: pivot.rows,
      values: pivot.values,
    });
    setDirty(false);
    setOpenMenu(undefined);
  }, [pivot.id, pivot.filters, pivot.columns, pivot.rows, pivot.values]);

  useLayoutEffect(() => {
    const menu = menuRef.current;
    if (!menu || !openMenu) return;
    const w = menu.offsetWidth || 186;
    const h = menu.offsetHeight || 240;
    const { anchor } = openMenu;
    let left = anchor.right - w;
    if (left < 8) left = Math.min(anchor.left, window.innerWidth - w - 8);
    left = Math.max(8, Math.min(left, window.innerWidth - w - 8));
    let top = anchor.bottom + 2;
    if (top + h > window.innerHeight - 8 && anchor.top - h - 2 >= 8) {
      top = anchor.top - h - 2;
    }
    top = Math.max(8, Math.min(top, window.innerHeight - h - 8));
    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;
  }, [openMenu]);

  useEffect(() => {
    if (!openMenu) return;
    function dismiss(event: PointerEvent) {
      if (!menuRef.current?.contains(event.target as Node)) setOpenMenu(undefined);
    }
    function onKey(event: globalThis.KeyboardEvent) {
      if (event.key === 'Escape') {
        event.preventDefault();
        setOpenMenu(undefined);
      }
    }
    window.document.addEventListener('pointerdown', dismiss);
    window.document.addEventListener('keydown', onKey);
    return () => {
      window.document.removeEventListener('pointerdown', dismiss);
      window.document.removeEventListener('keydown', onKey);
    };
  }, [openMenu]);

  const { fields, records } = useMemo(
    () => extractPivotRecords(document, pivot.source),
    [document, pivot.source],
  );
  const recordsRef = useRef(records);
  recordsRef.current = records;

  const visibleFields = useMemo(() => {
    const q = search.trim().toLowerCase();
    return q ? fields.filter(f => f.toLowerCase().includes(q)) : fields;
  }, [fields, search]);

  const activeFieldSet = useMemo(() => {
    const set = new Set<string>();
    for (const f of staged.filters) set.add(f.fieldName);
    for (const f of staged.columns) set.add(f.fieldName);
    for (const f of staged.rows) set.add(f.fieldName);
    for (const f of staged.values) set.add(f.fieldName);
    return set;
  }, [staged]);

  function commitLayout(next: StagedLayout) {
    setStaged(next);
    if (deferRef.current) {
      setDirty(true);
    } else {
      setDirty(false);
      onUpdateLayoutRef.current(next);
    }
  }

  function removeFieldFromAll(layout: StagedLayout, fieldName: string): StagedLayout {
    return {
      filters: layout.filters.filter(f => f.fieldName !== fieldName),
      columns: layout.columns.filter(f => f.fieldName !== fieldName),
      rows: layout.rows.filter(f => f.fieldName !== fieldName),
      values: layout.values.filter(f => f.fieldName !== fieldName),
    };
  }

  function toggleFieldCheck(fieldName: string) {
    if (activeFieldSet.has(fieldName)) {
      commitLayout(removeFieldFromAll(staged, fieldName));
      return;
    }
    if (defaultIsNumericField(records, fieldName)) {
      commitLayout({
        ...staged,
        values: [...staged.values, { id: makeId(), fieldName, aggregation: 'sum' }],
      });
    } else {
      commitLayout({
        ...staged,
        rows: [...staged.rows, { fieldName }],
      });
    }
  }

  function placeFieldInArea(fieldName: string, target: PivotArea, insertAt?: number) {
    const current = stagedRef.current;
    if (target === 'values') {
      const cleaned: StagedLayout = {
        ...current,
        filters: current.filters.filter(f => f.fieldName !== fieldName),
        columns: current.columns.filter(f => f.fieldName !== fieldName),
        rows: current.rows.filter(f => f.fieldName !== fieldName),
      };
      const defaultAgg = defaultIsNumericField(recordsRef.current, fieldName) ? 'sum' : 'count';
      const arr = [...cleaned.values];
      const idx = clampIndex(insertAt ?? arr.length, arr.length);
      arr.splice(idx, 0, { id: makeId(), fieldName, aggregation: defaultAgg });
      commitLayout({ ...cleaned, values: arr });
      return;
    }
    const base = removeFieldFromAll(current, fieldName);
    if (target === 'filters') {
      const arr = [...base.filters];
      arr.splice(clampIndex(insertAt ?? arr.length, arr.length), 0, { fieldName });
      commitLayout({ ...base, filters: arr });
    } else if (target === 'columns') {
      const arr = [...base.columns];
      arr.splice(clampIndex(insertAt ?? arr.length, arr.length), 0, { fieldName });
      commitLayout({ ...base, columns: arr });
    } else {
      const arr = [...base.rows];
      arr.splice(clampIndex(insertAt ?? arr.length, arr.length), 0, { fieldName });
      commitLayout({ ...base, rows: arr });
    }
  }

  function moveWithinArea(area: PivotArea, index: number, targetIndex: number) {
    const current = stagedRef.current;
    const arr = [...current[area]];
    const clamped = Math.max(0, Math.min(arr.length - 1, targetIndex));
    if (clamped === index) return;
    const [item] = arr.splice(index, 1);
    arr.splice(clamped, 0, item!);
    commitLayout({ ...current, [area]: arr } as StagedLayout);
  }

  function moveBetweenAreas(from: PivotArea, index: number, to: PivotArea, insertAt?: number) {
    if (from === to) return;
    const current = stagedRef.current;
    const sourceEntry = current[from][index];
    if (!sourceEntry) return;
    setOpenMenu(undefined);

    if (from !== 'values' && to !== 'values') {
      const sourceArr = [...current[from]];
      sourceArr.splice(index, 1);
      const filtered = removeFieldFromAll({ ...current, [from]: sourceArr } as StagedLayout, sourceEntry.fieldName);
      const targetArr = [...filtered[to]];
      const destIdx = clampIndex(insertAt ?? targetArr.length, targetArr.length);
      targetArr.splice(destIdx, 0, sourceEntry);
      commitLayout({ ...filtered, [to]: targetArr } as StagedLayout);
      return;
    }
    placeFieldInArea(sourceEntry.fieldName, to, insertAt);
  }

  function removeAt(area: PivotArea, index: number) {
    setOpenMenu(undefined);
    const current = stagedRef.current;
    const nextArr = [...current[area]];
    nextArr.splice(index, 1);
    commitLayout({ ...current, [area]: nextArr } as StagedLayout);
  }

  function resolveDropTarget(clientX: number, clientY: number): DropTarget | undefined {
    const pane = paneRef.current;
    if (!pane) return undefined;
    const boxes = pane.querySelectorAll<HTMLElement>('[data-pivot-area]');
    for (const box of boxes) {
      const rect = box.getBoundingClientRect();
      if (clientX < rect.left || clientX > rect.right || clientY < rect.top || clientY > rect.bottom) continue;
      const area = box.dataset.pivotArea as PivotArea;
      const chips = box.querySelectorAll<HTMLElement>('[data-pivot-chip-index]');
      let index = 0;
      for (const chip of chips) {
        const chipRect = chip.getBoundingClientRect();
        if (clientY > chipRect.top + chipRect.height / 2) {
          index = Number(chip.dataset.pivotChipIndex) + 1;
        }
      }
      return { area, index };
    }
    return undefined;
  }

  function startPointerDrag(
    event: ReactPointerEvent<HTMLElement>,
    init: Omit<DragState, 'startX' | 'startY' | 'x' | 'y' | 'active'>,
  ) {
    if (event.button !== 0) return;
    if ((event.target as Element).closest('.pivot-chip-trigger')) return;
    const state: DragState = {
      ...init,
      startX: event.clientX,
      startY: event.clientY,
      x: event.clientX,
      y: event.clientY,
      active: false,
    };
    dragRef.current = state;
  }

  useEffect(() => {
    function handleMove(event: PointerEvent) {
      const cur = dragRef.current;
      if (!cur) return;
      const dx = event.clientX - cur.startX;
      const dy = event.clientY - cur.startY;
      if (!cur.active && dx * dx + dy * dy < 20) return;
      const next: DragState = { ...cur, x: event.clientX, y: event.clientY, active: true };
      dragRef.current = next;
      suppressClickRef.current = true;
      setDrag(next);
      setDropTarget(resolveDropTarget(event.clientX, event.clientY));
    }

    function handleEnd(event: PointerEvent) {
      const cur = dragRef.current;
      if (!cur) return;
      dragRef.current = undefined;
      setDrag(undefined);
      setDropTarget(undefined);
      if (!cur.active) return;

      const target = resolveDropTarget(event.clientX, event.clientY);
      if (!target) {
        if (cur.kind === 'chip' && cur.fromArea !== undefined && cur.fromIndex !== undefined) {
          removeAt(cur.fromArea, cur.fromIndex);
        }
      } else if (cur.kind === 'field') {
        placeFieldInArea(cur.fieldName, target.area, target.index);
      } else if (cur.kind === 'chip' && cur.fromArea !== undefined && cur.fromIndex !== undefined) {
        if (cur.fromArea === target.area) {
          const adjusted = target.index > cur.fromIndex ? target.index - 1 : target.index;
          moveWithinArea(target.area, cur.fromIndex, adjusted);
        } else {
          moveBetweenAreas(cur.fromArea, cur.fromIndex, target.area, target.index);
        }
      }
      window.setTimeout(() => { suppressClickRef.current = false; }, 0);
    }

    function handleCancel() {
      if (!dragRef.current) return;
      dragRef.current = undefined;
      suppressClickRef.current = false;
      setDrag(undefined);
      setDropTarget(undefined);
    }

    function handleKey(event: globalThis.KeyboardEvent) {
      if (event.key === 'Escape' && dragRef.current) {
        event.preventDefault();
        handleCancel();
      }
    }

    window.addEventListener('pointermove', handleMove);
    window.addEventListener('pointerup', handleEnd);
    window.addEventListener('pointercancel', handleCancel);
    window.addEventListener('keydown', handleKey);
    return () => {
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleEnd);
      window.removeEventListener('pointercancel', handleCancel);
      window.removeEventListener('keydown', handleKey);
    };
    // Handler refs access current staged state and helper closures cleanly.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const areas: readonly { readonly key: PivotArea; readonly title: string }[] = [
    { key: 'filters', title: 'Filters' },
    { key: 'columns', title: 'Columns' },
    { key: 'rows', title: 'Rows' },
    { key: 'values', title: 'Values' },
  ];

  return (
    <aside ref={paneRef} className="pivot-pane" aria-label="PivotTable Fields" style={{ width }}>
      <div className="python-panel-header pivot-pane-header">
        <span className="python-panel-title">PivotTable Fields</span>
        <button
          type="button"
          className="python-hide"
          aria-label="Close PivotTable Fields"
          onClick={onClose}
        >
          ×
        </button>
      </div>

      <div
        className="pivot-pane-body"
        onScroll={() => { if (openMenu) setOpenMenu(undefined); }}
        onClickCapture={event => {
          if (suppressClickRef.current) {
            event.preventDefault();
            event.stopPropagation();
          }
        }}
      >
        <label className="pivot-pane-caption" htmlFor={searchId}>Choose fields to add to report:</label>
        <input
          id={searchId}
          type="search"
          className="pivot-field-search"
          placeholder="Search"
          value={search}
          onChange={event => setSearch(event.currentTarget.value)}
        />

        <div className="pivot-field-list" role="group" aria-label="Available PivotTable fields">
          {visibleFields.length === 0 ? (
            <p className="pivot-fields-empty">No matching fields.</p>
          ) : (
            visibleFields.map(fieldName => {
              const checked = activeFieldSet.has(fieldName);
              return (
                <label
                  key={fieldName}
                  className={'pivot-field-check' + (checked ? ' checked' : '')}
                  onPointerDown={event => {
                    if ((event.target as Element).tagName === 'INPUT') return;
                    startPointerDrag(event, { kind: 'field', fieldName, label: fieldName });
                  }}
                >
                  <input
                    type="checkbox"
                    checked={checked}
                    onChange={() => toggleFieldCheck(fieldName)}
                  />
                  <span>{fieldName}</span>
                </label>
              );
            })
          )}
        </div>

        <p className="pivot-pane-caption">Drag fields between areas below:</p>

        <div className="pivot-areas-grid">
          {areas.map(({ key, title }) => {
            const list = staged[key];
            const areaHover = drag?.active && dropTarget?.area === key;
            return (
              <section
                key={key}
                data-pivot-area={key}
                className={'pivot-area-box' + (areaHover ? ' drag-over' : '')}
                aria-label={`${title} area`}
              >
                <div className="pivot-area-title">{title}</div>
                <div className="pivot-area-items">
                  {list.map((entry, idx) => {
                    const chipKey = `${key}-${idx}-${entry.fieldName}`;
                    const label = key === 'values'
                      ? (entry as PivotValueField).customName ?? defaultPivotValueLabel(entry.fieldName, (entry as PivotValueField).aggregation)
                      : entry.fieldName;
                    const menuOpen = openMenu?.key === chipKey;
                    const dropBefore = areaHover && dropTarget?.index === idx;
                    const dropAfter = areaHover && idx === list.length - 1 && dropTarget?.index === list.length;
                    const isDragged = drag?.active && drag.kind === 'chip' && drag.fromArea === key && drag.fromIndex === idx;
                    return (
                      <div
                        key={chipKey}
                        data-pivot-chip-index={idx}
                        className={'pivot-chip-wrap' + (dropBefore ? ' drop-before' : '') + (dropAfter ? ' drop-after' : '')}
                      >
                        <div
                          className={'pivot-chip' + (isDragged ? ' dragging' : '')}
                          onPointerDown={event => startPointerDrag(event, {
                            kind: 'chip', fieldName: entry.fieldName, label, fromArea: key, fromIndex: idx,
                          })}
                        >
                          <span className="pivot-chip-label">{label}</span>
                          <button
                            type="button"
                            className="pivot-chip-trigger"
                            aria-label={`${label} options`}
                            aria-haspopup="menu"
                            aria-expanded={menuOpen}
                            onPointerDown={event => event.stopPropagation()}
                            onClick={event => {
                              const rect = event.currentTarget.getBoundingClientRect();
                              setOpenMenu(menuOpen ? undefined : {
                                key: chipKey,
                                area: key,
                                index: idx,
                                label,
                                entry,
                                anchor: { left: rect.left, right: rect.right, top: rect.top, bottom: rect.bottom },
                              });
                            }}
                          >
                            ▾
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </section>
            );
          })}
        </div>
      </div>

      {drag?.active && (
        <div
          className="pivot-drag-ghost"
          aria-hidden="true"
          style={{ left: drag.x + 10, top: drag.y + 10 }}
        >
          {drag.label}
        </div>
      )}

      {openMenu && (
        <div ref={menuRef} className="pivot-chip-menu" role="menu" aria-label={`${openMenu.label} actions`}>
          <button type="button" role="menuitem" disabled={openMenu.index === 0}
            onClick={() => { setOpenMenu(undefined); moveWithinArea(openMenu.area, openMenu.index, openMenu.index - 1); }}>Move Up</button>
          <button type="button" role="menuitem" disabled={openMenu.index === staged[openMenu.area].length - 1}
            onClick={() => { setOpenMenu(undefined); moveWithinArea(openMenu.area, openMenu.index, openMenu.index + 1); }}>Move Down</button>
          <button type="button" role="menuitem" disabled={openMenu.index === 0}
            onClick={() => { setOpenMenu(undefined); moveWithinArea(openMenu.area, openMenu.index, 0); }}>Move to Beginning</button>
          <button type="button" role="menuitem" disabled={openMenu.index === staged[openMenu.area].length - 1}
            onClick={() => { setOpenMenu(undefined); moveWithinArea(openMenu.area, openMenu.index, staged[openMenu.area].length - 1); }}>Move to End</button>
          {openMenu.area !== 'filters' && <button type="button" role="menuitem"
            onClick={() => moveBetweenAreas(openMenu.area, openMenu.index, 'filters')}>Move to Report Filter</button>}
          {openMenu.area !== 'columns' && <button type="button" role="menuitem"
            onClick={() => moveBetweenAreas(openMenu.area, openMenu.index, 'columns')}>Move to Column Labels</button>}
          {openMenu.area !== 'rows' && <button type="button" role="menuitem"
            onClick={() => moveBetweenAreas(openMenu.area, openMenu.index, 'rows')}>Move to Row Labels</button>}
          {openMenu.area !== 'values' && <button type="button" role="menuitem"
            onClick={() => moveBetweenAreas(openMenu.area, openMenu.index, 'values')}>Move to Values</button>}
          <button type="button" role="menuitem"
            onClick={() => removeAt(openMenu.area, openMenu.index)}>Remove Field</button>
          {openMenu.area === 'values' && (
            <button type="button" role="menuitem"
              onClick={() => { setOpenMenu(undefined); onOpenSettings({ mode: 'value', field: openMenu.entry as PivotValueField }); }}>
              Value Field Settings...
            </button>
          )}
          {(openMenu.area === 'rows' || openMenu.area === 'columns') && (
            <button type="button" role="menuitem"
              onClick={() => { setOpenMenu(undefined); onOpenSettings({ mode: 'axis', area: openMenu.area as 'rows' | 'columns', field: openMenu.entry as PivotAxisField }); }}>
              Field Settings...
            </button>
          )}
        </div>
      )}

      <div className="pivot-pane-footer">
        <label className="pivot-defer-check" htmlFor={deferId}>
          <input
            id={deferId}
            type="checkbox"
            checked={defer}
            onChange={event => {
              const checked = event.currentTarget.checked;
              setDefer(checked);
              if (!checked && dirty) {
                setDirty(false);
                onUpdateLayout(staged);
              }
            }}
          />
          Defer Layout Update
        </label>
        <button
          type="button"
          className="file-primary"
          disabled={!defer || !dirty}
          onClick={() => {
            setDirty(false);
            onUpdateLayout(staged);
          }}
        >
          Update
        </button>
      </div>
    </aside>
  );
}
