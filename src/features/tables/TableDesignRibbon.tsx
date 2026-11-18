import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import type { TableDesign, WorkbookName } from '../../core/workbook';
import { tableDesign, tableStyles } from '../../core/workbook-names';

interface TableDesignRibbonProps {
  readonly table: WorkbookName;
  readonly onRename: (name: string) => void;
  readonly onResize: () => void;
  readonly onChangeDesign: (patch: Partial<TableDesign>) => void;
}

function ribbonTint(hex: string): string {
  const channels = [0, 2, 4].map(index => Math.trunc((Number.parseInt(hex.slice(index, index + 2), 16) + 510) / 3));
  return '#' + channels.map(channel => channel.toString(16).padStart(2, '0')).join('');
}
const styleMains: Record<string, string> = {
  Blue: '4F81BD', Green: '548235', Orange: 'C65911', Purple: '8064A2',
};

export function TableDesignRibbon({ table, onRename, onResize, onChangeDesign }: TableDesignRibbonProps) {
  const [nameDraft, setNameDraft] = useState(table.name);
  const editing = useRef(false);
  const design = tableDesign(table);

  useEffect(() => {
    if (!editing.current) setNameDraft(table.name);
  }, [table.name]);

  function commitRename() {
    editing.current = false;
    const trimmed = nameDraft.trim();
    if (trimmed && trimmed !== table.name) {
      onRename(trimmed);
    } else {
      setNameDraft(table.name);
    }
  }

  function handleKeyDown(event: KeyboardEvent<HTMLInputElement>) {
    if (event.key === 'Enter') {
      event.preventDefault();
      commitRename();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      editing.current = false;
      setNameDraft(table.name);
    }
  }

  return (
    <>
      <fieldset className="table-properties-group">
        <legend>Properties</legend>
        <div className="table-properties-content">
          <label className="table-name-row">
            <span>Table Name:</span>
            <input
              type="text"
              className="table-name-input"
              value={nameDraft}
              maxLength={255}
              aria-label="Table Name"
              onFocus={() => { editing.current = true; }}
              onChange={e => setNameDraft(e.target.value)}
              onKeyDown={handleKeyDown}
              onBlur={commitRename}
            />
          </label>
          <button
            type="button"
            className="ribbon-command table-resize-btn"
            onClick={onResize}
          >
            Resize Table
          </button>
        </div>
      </fieldset>

      <fieldset className="table-options-group">
        <legend>Table Style Options</legend>
        <div className="table-options-grid">
          <label>
            <input
              type="checkbox"
              aria-label="Header Row"
              checked={design.headerRow}
              onChange={e => onChangeDesign({ headerRow: e.target.checked })}
            />
            Header Row
          </label>
          <label>
            <input
              type="checkbox"
              aria-label="First Column"
              checked={design.firstColumn}
              onChange={e => onChangeDesign({ firstColumn: e.target.checked })}
            />
            First Column
          </label>
          <label>
            <input
              type="checkbox"
              aria-label="Filter Button"
              disabled={!design.headerRow}
              checked={design.filterButton && design.headerRow}
              onChange={e => onChangeDesign({ filterButton: e.target.checked })}
            />
            Filter Button
          </label>
          <label>
            <input
              type="checkbox"
              aria-label="Total Row"
              checked={design.totalRow}
              onChange={e => onChangeDesign({ totalRow: e.target.checked })}
            />
            Total Row
          </label>
          <label>
            <input
              type="checkbox"
              aria-label="Last Column"
              checked={design.lastColumn}
              onChange={e => onChangeDesign({ lastColumn: e.target.checked })}
            />
            Last Column
          </label>
          <span />
          <label>
            <input
              type="checkbox"
              aria-label="Banded Rows"
              checked={design.bandedRows}
              onChange={e => onChangeDesign({ bandedRows: e.target.checked })}
            />
            Banded Rows
          </label>
          <label>
            <input
              type="checkbox"
              aria-label="Banded Columns"
              checked={design.bandedColumns}
              onChange={e => onChangeDesign({ bandedColumns: e.target.checked })}
            />
            Banded Columns
          </label>
        </div>
      </fieldset>

      <fieldset className="table-styles-group">
        <legend>Table Styles</legend>
        <div className="table-styles-row">
          {tableStyles.map(style => {
            const main = styleMains[style] ?? '303844';
            const colors = { main: '#' + main, tint: ribbonTint(main) };
            const isSelected = design.style === style;
            return (
              <button
                key={style}
                type="button"
                className={'table-style-preview-btn' + (isSelected ? ' selected' : '')}
                aria-label={style + ' table style'}
                title={style}
                onClick={() => onChangeDesign({ style })}
              >
                <div className="table-style-swatch">
                  <div className="table-style-swatch-header" style={{ backgroundColor: colors.main }} />
                  <div className="table-style-swatch-row" style={{ backgroundColor: colors.tint }} />
                  <div className="table-style-swatch-row" style={{ backgroundColor: colors.tint }} />
                  <div className="table-style-swatch-row" style={{ backgroundColor: colors.tint }} />
                  <div className="table-style-swatch-row" style={{ backgroundColor: colors.tint }} />
                </div>
              </button>
            );
          })}
        </div>
      </fieldset>
    </>
  );
}
