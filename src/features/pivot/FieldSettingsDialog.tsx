import { useId, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { extractPivotRecords, formatPivotItemLabel } from '../../calculation/pivot-engine';
import {
  defaultPivotValueLabel,
  type PivotAggregation,
  type PivotAxisField,
  type PivotShowValuesAs,
  type PivotSortRule,
  type PivotTableDocument,
  type PivotValueField,
} from '../../core/pivot';
import type { WorkbookDocument } from '../../core/workbook';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';

const aggregationChoices: readonly { readonly id: PivotAggregation; readonly label: string }[] = [
  { id: 'sum', label: 'Sum' },
  { id: 'count', label: 'Count' },
  { id: 'average', label: 'Average' },
  { id: 'max', label: 'Max' },
  { id: 'min', label: 'Min' },
  { id: 'product', label: 'Product' },
  { id: 'countNumbers', label: 'Count Numbers' },
  { id: 'stdDev', label: 'StdDev' },
  { id: 'stdDevp', label: 'StdDevp' },
  { id: 'var', label: 'Var' },
  { id: 'varp', label: 'Varp' },
];

const showValuesAsChoices: readonly { readonly id: PivotShowValuesAs; readonly label: string }[] = [
  { id: 'none', label: 'No Calculation' },
  { id: 'percentOfGrandTotal', label: '% of Grand Total' },
  { id: 'percentOfColumnTotal', label: '% of Column Total' },
  { id: 'percentOfRowTotal', label: '% of Row Total' },
  { id: 'percentOfParentRowTotal', label: '% of Parent Row Total' },
  { id: 'percentOfParentColumnTotal', label: '% of Parent Column Total' },
  { id: 'differenceFrom', label: 'Difference From' },
  { id: 'percentDifferenceFrom', label: '% Difference From' },
  { id: 'runningTotalIn', label: 'Running Total In' },
  { id: 'rankAscending', label: 'Rank Smallest to Largest' },
  { id: 'rankDescending', label: 'Rank Largest to Smallest' },
  { id: 'index', label: 'Index' },
];

const numberFormatPresets = ['General', '#,##0.00', '$#,##0.00', '0.00%', '#,##0'] as const;

export type FieldSettingsRequest =
  | { readonly mode: 'value'; readonly field: PivotValueField }
  | { readonly mode: 'axis'; readonly area: 'rows' | 'columns'; readonly field: PivotAxisField };

export function FieldSettingsDialog({
  document,
  pivot,
  request,
  onSaveValueField,
  onSaveAxisField,
  onCancel,
}: {
  readonly document: WorkbookDocument;
  readonly pivot: PivotTableDocument;
  readonly request: FieldSettingsRequest;
  readonly onSaveValueField: (updated: PivotValueField) => void;
  readonly onSaveAxisField: (area: 'rows' | 'columns', updated: PivotAxisField) => void;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const customNameId = useId();
  const formatId = useId();
  const formatListId = useId();
  const showAsId = useId();
  const baseFieldId = useId();
  const baseItemId = useId();
  const sortById = useId();

  const initialValue = request.mode === 'value' ? request.field : undefined;
  const initialAxis = request.mode === 'axis' ? request.field : undefined;

  const [tab, setTab] = useState<'summarize' | 'showAs'>('summarize');
  const [aggregation, setAggregation] = useState<PivotAggregation>(initialValue?.aggregation ?? 'sum');
  const [customName, setCustomName] = useState(
    initialValue?.customName ?? (initialValue ? defaultPivotValueLabel(initialValue.fieldName, initialValue.aggregation) : ''),
  );
  const [customEdited, setCustomEdited] = useState(Boolean(initialValue?.customName));
  const [showValuesAs, setShowValuesAs] = useState<PivotShowValuesAs>(initialValue?.showValuesAs ?? 'none');
  const [baseField, setBaseField] = useState<string>(initialValue?.baseField ?? pivot.rows[0]?.fieldName ?? pivot.columns[0]?.fieldName ?? '');
  const [baseItem, setBaseItem] = useState<string>(initialValue?.baseItem ?? '(previous)');
  const [numberFormat, setNumberFormat] = useState<string>(initialValue?.numberFormat ?? 'General');

  const [subtotal, setSubtotal] = useState<'automatic' | 'none'>(initialAxis?.subtotal ?? 'automatic');
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>(initialAxis?.sort?.direction ?? 'asc');
  const [sortByValueId, setSortByValueId] = useState<string>(initialAxis?.sort?.byValueFieldId ?? '');

  const { fields, records } = useMemo(
    () => extractPivotRecords(document, pivot.source),
    [document, pivot.source],
  );

  const baseFieldChoices = useMemo(() => {
    const axis = [...pivot.rows.map(r => r.fieldName), ...pivot.columns.map(c => c.fieldName)];
    return axis.length ? [...new Set(axis)] : [...fields];
  }, [pivot.rows, pivot.columns, fields]);

  const baseItemChoices = useMemo(() => {
    const items = baseField
      ? [...new Set(records.map(r => formatPivotItemLabel(r[baseField] ?? null)))].sort((a, b) => a.localeCompare(b))
      : [];
    return ['(previous)', '(next)', ...items];
  }, [records, baseField]);

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (dialog && !dialog.open) dialog.showModal();
  }, []);

  const title = request.mode === 'value' ? 'Value Field Settings' : 'Field Settings';
  const needsBaseField = showValuesAs === 'differenceFrom' || showValuesAs === 'percentDifferenceFrom'
    || showValuesAs === 'runningTotalIn' || showValuesAs === 'percentOfParentRowTotal'
    || showValuesAs === 'percentOfParentColumnTotal';
  const needsBaseItem = showValuesAs === 'differenceFrom' || showValuesAs === 'percentDifferenceFrom';

  function handleConfirm() {
    if (request.mode === 'value') {
      const f = request.field;
      const defaultName = defaultPivotValueLabel(f.fieldName, aggregation);
      const trimmedName = customName.trim();
      const trimmedFmt = numberFormat.trim();
      onSaveValueField({
        id: f.id,
        fieldName: f.fieldName,
        ...(trimmedName && trimmedName !== defaultName && { customName: trimmedName }),
        aggregation,
        ...(showValuesAs !== 'none' && { showValuesAs }),
        ...(needsBaseField && baseField && { baseField }),
        ...(needsBaseItem && baseItem && { baseItem }),
        ...(trimmedFmt && trimmedFmt !== 'General' && { numberFormat: trimmedFmt }),
      });
    } else {
      const sort: PivotSortRule = {
        direction: sortDir,
        ...(sortByValueId && { byValueFieldId: sortByValueId }),
      };
      onSaveAxisField(request.area, {
        ...request.field,
        subtotal,
        sort,
      });
    }
  }

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog field-settings-dialog"
      aria-label={title}
      onCancel={event => {
        event.preventDefault();
        onCancel();
      }}
    >
      <form
        onSubmit={event => {
          event.preventDefault();
          handleConfirm();
        }}
      >
        <DialogTitleBar>
          <h2>{title}</h2>
          <button
            type="button"
            className="format-dialog-close"
            aria-label={`Close ${title}`}
            onClick={onCancel}
          >
            ×
          </button>
        </DialogTitleBar>

        <div className="field-settings-meta">
          <span>Source Name:</span>
          <strong>{request.field.fieldName}</strong>
        </div>

        {request.mode === 'value' ? (
          <>
            <div className="field-settings-row">
              <label htmlFor={customNameId}>Custom Name:</label>
              <input
                id={customNameId}
                type="text"
                value={customName}
                onChange={event => {
                  setCustomEdited(true);
                  setCustomName(event.currentTarget.value);
                }}
              />
            </div>

            <div className="find-tabs" role="tablist" aria-label="Value Field Options">
              <button
                type="button"
                role="tab"
                aria-selected={tab === 'summarize'}
                onClick={() => setTab('summarize')}
              >
                Summarize Values By
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={tab === 'showAs'}
                onClick={() => setTab('showAs')}
              >
                Show Values As
              </button>
            </div>

            {tab === 'summarize' ? (
              <div className="field-settings-panel">
                <span className="field-settings-label">Summarize value field by:</span>
                <div role="listbox" aria-label="Summarize value field by" className="field-settings-listbox">
                  {aggregationChoices.map(choice => {
                    const selected = aggregation === choice.id;
                    return (
                      <div
                        key={choice.id}
                        role="option"
                        aria-selected={selected}
                        tabIndex={0}
                        className={'insert-function-option' + (selected ? ' selected' : '')}
                        onClick={() => {
                          setAggregation(choice.id);
                          if (!customEdited) setCustomName(defaultPivotValueLabel(request.field.fieldName, choice.id));
                        }}
                        onKeyDown={event => {
                          if (event.key === 'Enter' || event.key === ' ') {
                            event.preventDefault();
                            setAggregation(choice.id);
                            if (!customEdited) setCustomName(defaultPivotValueLabel(request.field.fieldName, choice.id));
                          }
                        }}
                      >
                        {choice.label}
                      </div>
                    );
                  })}
                </div>
              </div>
            ) : (
              <div className="field-settings-panel">
                <div className="field-settings-row">
                  <label htmlFor={showAsId}>Show values as:</label>
                  <select
                    id={showAsId}
                    value={showValuesAs}
                    onChange={event => setShowValuesAs(event.currentTarget.value as PivotShowValuesAs)}
                  >
                    {showValuesAsChoices.map(choice => (
                      <option key={choice.id} value={choice.id}>{choice.label}</option>
                    ))}
                  </select>
                </div>
                {needsBaseField && (
                  <div className="field-settings-row">
                    <label htmlFor={baseFieldId}>Base field:</label>
                    <select
                      id={baseFieldId}
                      value={baseField}
                      onChange={event => setBaseField(event.currentTarget.value)}
                    >
                      {baseFieldChoices.map(name => (
                        <option key={name} value={name}>{name}</option>
                      ))}
                    </select>
                  </div>
                )}
                {needsBaseItem && (
                  <div className="field-settings-row">
                    <label htmlFor={baseItemId}>Base item:</label>
                    <select
                      id={baseItemId}
                      value={baseItem}
                      onChange={event => setBaseItem(event.currentTarget.value)}
                    >
                      {baseItemChoices.map(item => (
                        <option key={item} value={item}>{item}</option>
                      ))}
                    </select>
                  </div>
                )}
              </div>
            )}

            <div className="field-settings-row">
              <label htmlFor={formatId}>Number Format:</label>
              <input
                id={formatId}
                type="text"
                list={formatListId}
                value={numberFormat}
                onChange={event => setNumberFormat(event.currentTarget.value)}
              />
              <datalist id={formatListId}>
                {numberFormatPresets.map(code => (
                  <option key={code} value={code} />
                ))}
              </datalist>
            </div>
          </>
        ) : (
          <div className="field-settings-panel">
            <fieldset className="pivot-dialog-group">
              <legend>Subtotals</legend>
              <label className="pivot-inline-radio">
                <input
                  type="radio"
                  name="axis-subtotal"
                  checked={subtotal === 'automatic'}
                  onChange={() => setSubtotal('automatic')}
                />
                Automatic
              </label>
              <label className="pivot-inline-radio">
                <input
                  type="radio"
                  name="axis-subtotal"
                  checked={subtotal === 'none'}
                  onChange={() => setSubtotal('none')}
                />
                None
              </label>
            </fieldset>

            <fieldset className="pivot-dialog-group">
              <legend>Sort</legend>
              <label className="pivot-inline-radio">
                <input
                  type="radio"
                  name="axis-sort"
                  checked={sortDir === 'asc'}
                  onChange={() => setSortDir('asc')}
                />
                Ascending (A to Z)
              </label>
              <label className="pivot-inline-radio">
                <input
                  type="radio"
                  name="axis-sort"
                  checked={sortDir === 'desc'}
                  onChange={() => setSortDir('desc')}
                />
                Descending (Z to A)
              </label>
              <div className="field-settings-row">
                <label htmlFor={sortById}>Sort by:</label>
                <select
                  id={sortById}
                  value={sortByValueId}
                  onChange={event => setSortByValueId(event.currentTarget.value)}
                >
                  <option value="">{request.field.fieldName} (label)</option>
                  {pivot.values.map(vf => (
                    <option key={vf.id} value={vf.id}>
                      {vf.customName ?? defaultPivotValueLabel(vf.fieldName, vf.aggregation)}
                    </option>
                  ))}
                </select>
              </div>
            </fieldset>
          </div>
        )}

        <DialogActions confirmLabel="OK" onCancel={onCancel} />
      </form>
    </dialog>
  );
}
