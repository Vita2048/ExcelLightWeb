import { useEffect, useId, useState } from 'react';
import type { ActivePivotFieldInfo } from '../../calculation/pivot-engine';
import type { PivotTableDocument } from '../../core/pivot';
import { RibbonFlyout, RibbonGlyph } from '../shell/RibbonFlyout';

export interface PivotAnalyzeRibbonProps {
  readonly pivot: PivotTableDocument;
  readonly activeField?: ActivePivotFieldInfo | undefined;
  readonly fieldListOpen: boolean;
  readonly onRename: (name: string) => void;
  readonly onOpenOptions: () => void;
  readonly onOpenFieldSettings: () => void;
  readonly onExpandField: () => void;
  readonly onCollapseField: () => void;
  readonly onRefresh: () => void;
  readonly onRefreshAll: () => void;
  readonly onOpenChangeSource: () => void;
  readonly onClear: (mode: 'all' | 'filters') => void;
  readonly onSelectEntire: () => void;
  readonly onOpenMove: () => void;
  readonly onToggleFieldList: () => void;
  readonly onToggleButtons: () => void;
  readonly onToggleFieldHeaders: () => void;
}

export function PivotAnalyzeRibbon({
  pivot,
  activeField,
  fieldListOpen,
  onRename,
  onOpenOptions,
  onOpenFieldSettings,
  onExpandField,
  onCollapseField,
  onRefresh,
  onRefreshAll,
  onOpenChangeSource,
  onClear,
  onSelectEntire,
  onOpenMove,
  onToggleFieldList,
  onToggleButtons,
  onToggleFieldHeaders,
}: PivotAnalyzeRibbonProps) {
  const nameId = useId();
  const [nameDraft, setNameDraft] = useState(pivot.name);

  useEffect(() => {
    setNameDraft(pivot.name);
  }, [pivot.id, pivot.name]);

  function commitRename() {
    const trimmed = nameDraft.trim();
    if (!trimmed || trimmed === pivot.name) {
      setNameDraft(pivot.name);
      return;
    }
    onRename(trimmed);
  }

  const canOpenFieldSettings = Boolean(activeField?.axisField || activeField?.valueField);
  const canExpandCollapse = Boolean(activeField?.canExpandCollapse);

  return <>
    <fieldset>
      <legend>PivotTable</legend>
      <div className="pivot-properties-content">
        <label className="pivot-name-label" htmlFor={nameId}>
          <span>PivotTable Name:</span>
          <input
            id={nameId}
            type="text"
            className="pivot-name-input"
            value={nameDraft}
            onChange={event => setNameDraft(event.currentTarget.value)}
            onBlur={commitRename}
            onKeyDown={event => {
              if (event.key === 'Enter') {
                commitRename();
                event.currentTarget.blur();
              } else if (event.key === 'Escape') {
                setNameDraft(pivot.name);
                event.currentTarget.blur();
              }
            }}
          />
        </label>
        <button type="button" className="ribbon-command" onClick={onOpenOptions}>
          <RibbonGlyph name="Pivot Options" size={16} />
          Options
        </button>
      </div>
    </fieldset>

    <fieldset>
      <legend>Active Field</legend>
      <div className="pivot-properties-content">
        <span className="pivot-active-field-readout" aria-label="Active Field" title={activeField?.displayName ?? 'None'}>
          {activeField?.displayName ?? 'None'}
        </span>
        <button
          type="button"
          className="ribbon-command"
          disabled={!canOpenFieldSettings}
          onClick={onOpenFieldSettings}
        >
          <RibbonGlyph name="Field Settings" size={16} />
          Field Settings
        </button>
      </div>
      <div className="pivot-ribbon-stack">
        <button
          type="button"
          className="ribbon-command"
          disabled={!canExpandCollapse}
          onClick={onExpandField}
        >
          <RibbonGlyph name="Expand Field" size={16} />
          Expand Field
        </button>
        <button
          type="button"
          className="ribbon-command"
          disabled={!canExpandCollapse}
          onClick={onCollapseField}
        >
          <RibbonGlyph name="Collapse Field" size={16} />
          Collapse Field
        </button>
      </div>
    </fieldset>

    <fieldset>
      <legend>Data</legend>
      <RibbonFlyout
        label="Refresh"
        icon="Refresh PivotTable"
        large
        items={[
          { label: 'Refresh', icon: 'Refresh PivotTable', select: onRefresh },
          { label: 'Refresh All', icon: 'Refresh All', select: onRefreshAll },
        ]}
      />
      <button
        type="button"
        className="ribbon-command ribbon-large"
        aria-label="Change Data Source"
        onClick={onOpenChangeSource}
      >
        <RibbonGlyph name="Change Data Source" size={30} />
        <span>Change Data<br />Source</span>
      </button>
    </fieldset>

    <fieldset>
      <legend>Actions</legend>
      <RibbonFlyout
        label="Clear"
        icon="Clear PivotTable"
        large
        items={[
          { label: 'Clear All', select: () => onClear('all') },
          { label: 'Clear Filters', select: () => onClear('filters') },
        ]}
      />
      <RibbonFlyout
        label="Select"
        icon="Select PivotTable"
        large
        items={[
          { label: 'Entire PivotTable', select: onSelectEntire },
        ]}
      />
      <button
        type="button"
        className="ribbon-command ribbon-large"
        aria-label="Move PivotTable"
        onClick={onOpenMove}
      >
        <RibbonGlyph name="Move PivotTable" size={30} />
        <span>Move<br />PivotTable</span>
      </button>
    </fieldset>

    <fieldset>
      <legend>Show</legend>
      <button
        type="button"
        className="ribbon-command ribbon-large pivot-ribbon-toggle"
        aria-label="Field List"
        aria-pressed={fieldListOpen}
        onClick={onToggleFieldList}
      >
        <RibbonGlyph name="Field List" size={30} />
        <span>Field<br />List</span>
      </button>
      <button
        type="button"
        className="ribbon-command ribbon-large pivot-ribbon-toggle"
        aria-label="+/- Buttons"
        aria-pressed={pivot.showButtons !== false}
        onClick={onToggleButtons}
      >
        <RibbonGlyph name="Pivot +/- Buttons" size={30} />
        <span>+/-<br />Buttons</span>
      </button>
      <button
        type="button"
        className="ribbon-command ribbon-large pivot-ribbon-toggle"
        aria-label="Field Headers"
        aria-pressed={pivot.showFieldHeaders !== false}
        onClick={onToggleFieldHeaders}
      >
        <RibbonGlyph name="Field Headers" size={30} />
        <span>Field<br />Headers</span>
      </button>
    </fieldset>
  </>;
}
