import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent } from 'react';
import type { NativeFunction } from '../../calculation/formula-functions';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import {
  dialogCategoryOptions, functionCategory, functionDescription, functionSignatureText, searchFunctions,
  type InsertFunctionFilter,
} from './function-library';

export function InsertFunctionDialog({
  initialCategory = 'Most Recently Used',
  recentFunctions,
  separator,
  onInsert,
  onCancel,
}: {
  readonly initialCategory?: InsertFunctionFilter;
  readonly recentFunctions: readonly NativeFunction[];
  readonly separator: ',' | ';';
  readonly onInsert: (name: NativeFunction) => void;
  readonly onCancel: () => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  const searchInput = useRef<HTMLInputElement>(null);
  const listbox = useRef<HTMLDivElement>(null);
  const searchId = useId();
  const categoryId = useId();
  const listLabelId = useId();
  const optionPrefix = useId();

  const [draftQuery, setDraftQuery] = useState('');
  const [query, setQuery] = useState('');
  const [category, setCategory] = useState<InsertFunctionFilter>(initialCategory);

  const matches = useMemo(
    () => searchFunctions(query, category, recentFunctions),
    [query, category, recentFunctions],
  );
  const [selected, setSelected] = useState<NativeFunction | undefined>(() => matches[0]);

  useLayoutEffect(() => {
    if (!dialog.current?.open) dialog.current?.showModal();
    searchInput.current?.focus();
    searchInput.current?.select();
  }, []);

  useEffect(() => {
    setSelected(current => (current && matches.includes(current) ? current : matches[0]));
  }, [matches]);

  useEffect(() => {
    if (!selected || !listbox.current) return;
    const active = listbox.current.querySelector<HTMLElement>('[aria-selected="true"]');
    active?.scrollIntoView({ block: 'nearest' });
  }, [selected]);

  function runSearch() {
    setQuery(draftQuery);
  }

  function handleCategoryChange(next: InsertFunctionFilter) {
    setCategory(next);
    setDraftQuery('');
    setQuery('');
  }

  function handleListKey(event: KeyboardEvent<HTMLDivElement>) {
    if (!matches.length) return;
    const current = selected ? matches.indexOf(selected) : -1;
    let next: number | undefined;
    if (event.key === 'ArrowDown') next = Math.min(matches.length - 1, current + 1);
    else if (event.key === 'ArrowUp') next = Math.max(0, current <= 0 ? 0 : current - 1);
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = matches.length - 1;
    if (next !== undefined) {
      event.preventDefault();
      setSelected(matches[next]);
    }
  }

  return (
    <dialog
      ref={dialog}
      className="paste-dialog insert-function-dialog"
      aria-label="Insert Function"
      onCancel={event => {
        event.preventDefault();
        onCancel();
      }}
    >
      <form
        onSubmit={event => {
          event.preventDefault();
          if (selected) onInsert(selected);
        }}
      >
        <DialogTitleBar>
          <h2>Insert Function</h2>
          <button type="button" className="format-dialog-close" aria-label="Close Insert Function" onClick={onCancel}>
            ×
          </button>
        </DialogTitleBar>

        <div className="insert-function-body">
          <div className="insert-function-field">
            <label htmlFor={searchId}>Search for a function:</label>
            <div className="insert-function-search">
              <input
                ref={searchInput}
                id={searchId}
                type="search"
                placeholder="Type a brief description of what you want to do and then click Go"
                value={draftQuery}
                onChange={event => {
                  const val = event.currentTarget.value;
                  setDraftQuery(val);
                  setQuery(val);
                }}
                onKeyDown={event => {
                  if (event.key === 'Enter') {
                    event.preventDefault();
                    runSearch();
                  }
                }}
              />
              <button type="button" onClick={runSearch}>Go</button>
            </div>
          </div>

          <div className="insert-function-field insert-function-category">
            <label htmlFor={categoryId}>Or select a category:</label>
            <select
              id={categoryId}
              value={query.trim() ? 'All' : category}
              onChange={event => handleCategoryChange(event.currentTarget.value as InsertFunctionFilter)}
            >
              {dialogCategoryOptions.map(option => (
                <option key={option} value={option}>{option}</option>
              ))}
            </select>
          </div>

          <div className="insert-function-field">
            <span id={listLabelId} className="insert-function-label">Select a function:</span>
            <div
              ref={listbox}
              role="listbox"
              aria-labelledby={listLabelId}
              aria-activedescendant={selected ? `${optionPrefix}-${selected}` : undefined}
              tabIndex={0}
              className="insert-function-list"
              onKeyDown={handleListKey}
            >
              {matches.length === 0 ? (
                <div className="insert-function-empty">No matching functions.</div>
              ) : (
                matches.map(name => {
                  const active = selected === name;
                  return (
                    <div
                      key={name}
                      id={`${optionPrefix}-${name}`}
                      role="option"
                      aria-selected={active}
                      className={'insert-function-option' + (active ? ' selected' : '')}
                      onClick={() => setSelected(name)}
                      onDoubleClick={() => onInsert(name)}
                    >
                      <span>{name}</span>
                      <small>{functionCategory(name)}</small>
                    </div>
                  );
                })
              )}
            </div>
          </div>

          <div className="insert-function-preview" aria-live="polite">
            {selected ? (
              <>
                <strong className="insert-function-signature">{functionSignatureText(selected, separator)}</strong>
                <p className="insert-function-desc">{functionDescription(selected)}</p>
              </>
            ) : (
              <p className="insert-function-desc">Select a category or try another search term.</p>
            )}
          </div>
        </div>

        <DialogActions confirmLabel="OK" disabled={!selected} onCancel={onCancel} />
      </form>
    </dialog>
  );
}
