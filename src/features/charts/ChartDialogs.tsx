import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react';
import {
  chartCategoryOf,
  type ChartDocument,
  type ChartHiddenEmptyMode,
  type ChartSeries,
} from '../../core/chart';
import type { DocumentId, IdFactory, SheetDocument, WorkbookDocument } from '../../core/workbook';
import type { ChartUpdatePatch } from '../../application/chart-commands';
import { inferChartSeriesFromSource, resolveChartData } from '../../charts/chart-data';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';

function SwitchArrowLeft() {
  return (
    <svg width="42" height="22" viewBox="0 0 42 22" aria-hidden="true" className="select-data-arrow">
      <path d="M38 6 H14 Q8 6 8 12 V17 M4 13 L8 18 L12 13" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function SwitchArrowRight() {
  return (
    <svg width="42" height="22" viewBox="0 0 42 22" aria-hidden="true" className="select-data-arrow">
      <path d="M4 6 H28 Q34 6 34 12 V17 M30 13 L34 18 L38 13" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

export function SelectDataDialog({
  document,
  chart,
  makeId,
  suspended,
  pickedRange,
  onConsumePickedRange,
  onPickRange,
  onSubmit,
  onCancel,
}: {
  readonly document: WorkbookDocument;
  readonly chart: ChartDocument;
  readonly makeId: IdFactory;
  readonly suspended: boolean;
  readonly pickedRange?: string | undefined;
  readonly onConsumePickedRange?: () => void;
  readonly onPickRange: (currentRange: string) => void;
  readonly onSubmit: (patch: ChartUpdatePatch) => void;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const rangeInputId = useId();
  const initialRange = chart.source.kind === 'range' ? chart.source.formula : '';
  const [rangeText, setRangeText] = useState(initialRange);
  const [seriesBy, setSeriesBy] = useState<'columns' | 'rows'>(chart.seriesBy);
  const [series, setSeries] = useState<readonly ChartSeries[]>(chart.series);
  const [selectedIdx, setSelectedIdx] = useState<number>(0);
  const [hiddenCats, setHiddenCats] = useState<ReadonlySet<number>>(
    () => new Set(chart.hiddenCategories ?? []),
  );
  const [hiddenEmptyMode, setHiddenEmptyMode] = useState<ChartHiddenEmptyMode>(chart.hiddenEmptyMode);
  const [showHiddenCells, setShowHiddenCells] = useState<boolean>(chart.showHiddenCells);

  const [editSeriesIdx, setEditSeriesIdx] = useState<number | 'new' | undefined>(undefined);
  const [seriesNameDraft, setSeriesNameDraft] = useState('');
  const [seriesValuesDraft, setSeriesValuesDraft] = useState('');
  const [seriesXDraft, setSeriesXDraft] = useState('');
  const [seriesSizeDraft, setSeriesSizeDraft] = useState('');

  const [editCatOpen, setEditCatOpen] = useState(false);
  const [catRefDraft, setCatRefDraft] = useState(chart.series[0]?.categoriesRef ?? '');
  const [hiddenCellsOpen, setHiddenCellsOpen] = useState(false);

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (suspended) {
      if (dialog.open) dialog.close();
      return;
    }
    if (!dialog.open) dialog.showModal();
  }, [suspended]);

  useEffect(() => {
    if (!pickedRange) return;
    const normalized = pickedRange.trim().startsWith('=') ? pickedRange.trim() : `=${pickedRange.trim()}`;
    setRangeText(normalized);
    reinferFromRange(normalized, seriesBy);
    onConsumePickedRange?.();
  }, [pickedRange]);

  const isScatter = chartCategoryOf(chart.chartType) === 'scatter';
  const isBubble = chart.chartType === 'bubble';

  // Omit hiddenCategories when previewing inside the dialog so unchecked categories remain visible to toggle back on
  const previewData = useMemo(() => {
    const { hiddenCategories: _hc, ...chartBase } = chart;
    const previewChart: ChartDocument = {
      ...chartBase,
      source: chart.source.kind === 'range' && rangeText.trim()
        ? { kind: 'range', formula: rangeText.trim().startsWith('=') ? rangeText.trim() : `=${rangeText.trim()}` }
        : chart.source,
      seriesBy,
      series,
      hiddenEmptyMode,
      showHiddenCells,
    };
    return resolveChartData(document, previewChart);
  }, [document, chart, rangeText, seriesBy, series, hiddenEmptyMode, showHiddenCells]);

  function reinferFromRange(formula: string, nextOrientation: 'columns' | 'rows') {
    const clean = formula.trim();
    if (!clean) return;
    const normalized = clean.startsWith('=') ? clean : `=${clean}`;
    const inferred = inferChartSeriesFromSource(
      document,
      { kind: 'range', formula: normalized },
      chart.placement.sheetId,
      chart.chartType,
      nextOrientation,
    );
    if (inferred.series.length > 0) {
      setSeriesBy(inferred.seriesBy);
      setSeries(inferred.series.map((s, idx) => ({ id: series[idx]?.id ?? makeId(), ...s })));
      setCatRefDraft(inferred.series[0]?.categoriesRef ?? '');
      setHiddenCats(new Set());
      setSelectedIdx(0);
    }
  }

  function openEditSeries(idx: number | 'new') {
    if (idx === 'new') {
      setSeriesNameDraft('');
      setSeriesValuesDraft('');
      setSeriesXDraft('');
      setSeriesSizeDraft('');
    } else {
      const target = series[idx];
      if (!target) return;
      setSeriesNameDraft(target.nameRef ?? target.name ?? '');
      setSeriesValuesDraft(target.valuesRef);
      setSeriesXDraft(target.xValuesRef ?? '');
      setSeriesSizeDraft(target.bubbleSizesRef ?? '');
    }
    setEditSeriesIdx(idx);
  }

  function saveEditSeries() {
    const val = seriesValuesDraft.trim();
    if (!val) return;
    const normalizedVal = val.startsWith('=') ? val : `=${val}`;
    const rawName = seriesNameDraft.trim();
    const isRefName = rawName.startsWith('=');
    const xVal = seriesXDraft.trim();
    const sVal = seriesSizeDraft.trim();

    if (editSeriesIdx === 'new') {
      const created: ChartSeries = {
        id: makeId(),
        ...(rawName ? (isRefName ? { nameRef: rawName } : { name: rawName }) : { name: `Series ${series.length + 1}` }),
        ...(catRefDraft.trim() && { categoriesRef: catRefDraft.trim() }),
        valuesRef: normalizedVal,
        ...(xVal && { xValuesRef: xVal.startsWith('=') ? xVal : `=${xVal}` }),
        ...(sVal && { bubbleSizesRef: sVal.startsWith('=') ? sVal : `=${sVal}` }),
      };
      setSeries([...series, created]);
      setSelectedIdx(series.length);
    } else if (typeof editSeriesIdx === 'number') {
      setSeries(series.map((item, i) => {
        if (i !== editSeriesIdx) return item;
        const { name: _n, nameRef: _nr, xValuesRef: _x, bubbleSizesRef: _b, ...rest } = item;
        return {
          ...rest,
          ...(rawName ? (isRefName ? { nameRef: rawName } : { name: rawName }) : {}),
          valuesRef: normalizedVal,
          ...(xVal && { xValuesRef: xVal.startsWith('=') ? xVal : `=${xVal}` }),
          ...(sVal && { bubbleSizesRef: sVal.startsWith('=') ? sVal : `=${sVal}` }),
        };
      }));
    }
    setEditSeriesIdx(undefined);
  }

  function moveSeries(delta: -1 | 1) {
    const target = selectedIdx + delta;
    if (target < 0 || target >= series.length) return;
    const copy = [...series];
    const [removed] = copy.splice(selectedIdx, 1);
    copy.splice(target, 0, removed!);
    setSeries(copy);
    setSelectedIdx(target);
  }

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog select-data-dialog"
      aria-label="Select Data Source"
      onCancel={event => {
        event.preventDefault();
        onCancel();
      }}
    >
      <form
        onSubmit={event => {
          event.preventDefault();
          const clean = rangeText.trim();
          const sortedHidden = [...hiddenCats].sort((a, b) => a - b);
          onSubmit({
            ...(chart.source.kind === 'range' && clean && {
              source: { kind: 'range', formula: clean.startsWith('=') ? clean : `=${clean}` },
            }),
            seriesBy,
            series,
            hiddenCategories: sortedHidden.length > 0 ? sortedHidden : undefined,
            hiddenEmptyMode,
            showHiddenCells,
          });
        }}
      >
        <DialogTitleBar>
          <h2>Select Data Source</h2>
          <button type="button" className="format-dialog-close" aria-label="Close Select Data Source" onClick={onCancel}>×</button>
        </DialogTitleBar>

        {chart.source.kind === 'range' && (
          <div className="pivot-dialog-range-row">
            <label htmlFor={rangeInputId}>Chart data range:</label>
            <input
              id={rangeInputId}
              type="text"
              value={rangeText}
              onChange={event => setRangeText(event.currentTarget.value)}
              onBlur={() => reinferFromRange(rangeText, seriesBy)}
            />
            <button
              type="button"
              className="select-data-pick-btn"
              onClick={() => onPickRange(rangeText)}
              title="Select range on worksheet"
              aria-label="Select chart data range on worksheet"
            >
              ↑
            </button>
          </div>
        )}

        <div className="select-data-switch-row">
          <SwitchArrowLeft />
          <button
            type="button"
            className="select-data-switch-btn"
            disabled={chart.source.kind !== 'range'}
            onClick={() => {
              const nextOrientation = seriesBy === 'columns' ? 'rows' : 'columns';
              setSeriesBy(nextOrientation);
              reinferFromRange(rangeText, nextOrientation);
            }}
          >
            Switch Row/Column
          </button>
          <SwitchArrowRight />
        </div>

        <div className="select-data-columns">
          <div className="select-data-box">
            <div className="select-data-box-title">Legend Entries (Series)</div>
            <div className="select-data-box-actions">
              <button type="button" onClick={() => openEditSeries('new')}>+ Add</button>
              <button type="button" disabled={series.length === 0} onClick={() => openEditSeries(selectedIdx)}>✎ Edit</button>
              <button
                type="button"
                disabled={series.length <= 1}
                onClick={() => {
                  const next = series.filter((_, i) => i !== selectedIdx);
                  setSeries(next);
                  setSelectedIdx(Math.max(0, Math.min(selectedIdx, next.length - 1)));
                }}
              >
                ✕ Remove
              </button>
              <button type="button" aria-label="Move series up" disabled={selectedIdx <= 0} onClick={() => moveSeries(-1)}>▲</button>
              <button type="button" aria-label="Move series down" disabled={selectedIdx >= series.length - 1} onClick={() => moveSeries(1)}>▼</button>
            </div>
            <div className="select-data-listbox" role="listbox" aria-label="Legend Entries (Series)">
              {series.map((s, idx) => {
                const resolved = previewData.series[idx];
                const label = resolved?.name || s.name || `Series ${idx + 1}`;
                return (
                  <div
                    key={s.id}
                    role="option"
                    aria-selected={selectedIdx === idx}
                    className={'select-data-list-item' + (selectedIdx === idx ? ' selected' : '')}
                    onClick={() => setSelectedIdx(idx)}
                  >
                    <input
                      type="checkbox"
                      aria-label={`Show ${label}`}
                      checked={!s.hidden}
                      onChange={event => {
                        const show = event.currentTarget.checked;
                        setSeries(series.map((entry, i) => {
                          if (i !== idx) return entry;
                          if (!show) return { ...entry, hidden: true };
                          const { hidden: _h, ...rest } = entry;
                          return rest;
                        }));
                      }}
                    />
                    <span>{label}</span>
                  </div>
                );
              })}
            </div>
          </div>

          <div className="select-data-box">
            <div className="select-data-box-title">Horizontal (Category) Axis Labels</div>
            <div className="select-data-box-actions">
              <button
                type="button"
                onClick={() => {
                  setCatRefDraft(series[0]?.categoriesRef ?? '');
                  setEditCatOpen(open => !open);
                }}
              >
                ✎ Edit
              </button>
            </div>
            {editCatOpen && (
              <div className="select-data-inline-editor">
                <label>
                  <span>Axis label range:</span>
                  <input
                    type="text"
                    value={catRefDraft}
                    onChange={event => setCatRefDraft(event.currentTarget.value)}
                    placeholder="=Sheet1!$A$2:$A$5"
                  />
                </label>
                <div className="select-data-inline-btns">
                  <button
                    type="button"
                    onClick={() => {
                      const clean = catRefDraft.trim();
                      const norm = clean ? (clean.startsWith('=') ? clean : `=${clean}`) : undefined;
                      setSeries(series.map(entry => {
                        const { categoriesRef: _prev, ...rest } = entry;
                        return norm ? { ...rest, categoriesRef: norm } : rest;
                      }));
                      setEditCatOpen(false);
                    }}
                  >
                    Apply
                  </button>
                  <button type="button" onClick={() => setEditCatOpen(false)}>Close</button>
                </div>
              </div>
            )}
            <div className="select-data-listbox" role="list" aria-label="Horizontal Axis Labels">
              {previewData.categories.map((cat, idx) => (
                <label key={idx} role="listitem" className="select-data-list-item">
                  <input
                    type="checkbox"
                    aria-label={`Show category ${cat}`}
                    checked={!hiddenCats.has(idx)}
                    onChange={event => {
                      const next = new Set(hiddenCats);
                      if (event.currentTarget.checked) next.delete(idx);
                      else next.add(idx);
                      setHiddenCats(next);
                    }}
                  />
                  <span>{cat}</span>
                </label>
              ))}
            </div>
          </div>
        </div>

        {editSeriesIdx !== undefined && (
          <div className="select-data-subpanel" role="group" aria-label="Edit Series">
            <div className="select-data-subpanel-title">{editSeriesIdx === 'new' ? 'Add Series' : 'Edit Series'}</div>
            <div className="select-data-subgrid">
              <label>
                <span>Series name:</span>
                <input type="text" value={seriesNameDraft} onChange={e => setSeriesNameDraft(e.currentTarget.value)} />
              </label>
              <label>
                <span>Series values:</span>
                <input type="text" value={seriesValuesDraft} onChange={e => setSeriesValuesDraft(e.currentTarget.value)} />
              </label>
              {isScatter && (
                <label>
                  <span>Series X values:</span>
                  <input type="text" value={seriesXDraft} onChange={e => setSeriesXDraft(e.currentTarget.value)} />
                </label>
              )}
              {isBubble && (
                <label>
                  <span>Series bubble size:</span>
                  <input type="text" value={seriesSizeDraft} onChange={e => setSeriesSizeDraft(e.currentTarget.value)} />
                </label>
              )}
            </div>
            <div className="select-data-inline-btns">
              <button type="button" onClick={saveEditSeries}>Save Series</button>
              <button type="button" onClick={() => setEditSeriesIdx(undefined)}>Cancel</button>
            </div>
          </div>
        )}

        {hiddenCellsOpen && (
          <div className="select-data-subpanel" role="group" aria-label="Hidden and Empty Cell Settings">
            <div className="select-data-subpanel-title">Hidden and Empty Cell Settings</div>
            <div className="select-data-radio-col">
              <span>Show empty cells as:</span>
              <label>
                <input type="radio" name="hiddenEmpty" checked={hiddenEmptyMode === 'gaps'} onChange={() => setHiddenEmptyMode('gaps')} />
                Gaps
              </label>
              <label>
                <input type="radio" name="hiddenEmpty" checked={hiddenEmptyMode === 'zero'} onChange={() => setHiddenEmptyMode('zero')} />
                Zero
              </label>
              <label>
                <input type="radio" name="hiddenEmpty" checked={hiddenEmptyMode === 'connect'} onChange={() => setHiddenEmptyMode('connect')} />
                Connect data points with line
              </label>
            </div>
            <label className="pivot-options-check">
              <input type="checkbox" checked={showHiddenCells} onChange={e => setShowHiddenCells(e.currentTarget.checked)} />
              Show data in hidden rows and columns
            </label>
            <div className="select-data-inline-btns">
              <button type="button" onClick={() => setHiddenCellsOpen(false)}>Done</button>
            </div>
          </div>
        )}

        <div className="select-data-bottom">
          <button type="button" onClick={() => setHiddenCellsOpen(open => !open)}>
            Hidden and Empty Cells
          </button>
          <DialogActions onCancel={onCancel} />
        </div>
      </form>
    </dialog>
  );
}

export function MoveChartDialog({
  sheets,
  chart,
  onSubmit,
  onCancel,
}: {
  readonly sheets: readonly SheetDocument[];
  readonly chart: ChartDocument;
  readonly onSubmit: (targetSheetId: DocumentId) => void;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const selectId = useId();
  const [targetSheetId, setTargetSheetId] = useState<DocumentId>(chart.placement.sheetId);

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (!dialog.open) dialog.showModal();
  }, []);

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog move-chart-dialog"
      aria-label="Move Chart"
      onCancel={event => {
        event.preventDefault();
        onCancel();
      }}
    >
      <form
        onSubmit={event => {
          event.preventDefault();
          onSubmit(targetSheetId);
        }}
      >
        <DialogTitleBar>
          <h2>Move Chart</h2>
          <button type="button" className="format-dialog-close" aria-label="Close Move Chart" onClick={onCancel}>×</button>
        </DialogTitleBar>

        <p className="move-chart-prompt">Choose where you want the chart to be placed:</p>
        <div className="move-chart-row">
          <label htmlFor={selectId}>Object in:</label>
          <select
            id={selectId}
            value={targetSheetId}
            onChange={event => setTargetSheetId(event.currentTarget.value)}
          >
            {sheets.map(s => (
              <option key={s.id} value={s.id}>{s.name}</option>
            ))}
          </select>
        </div>

        <DialogActions onCancel={onCancel} />
      </form>
    </dialog>
  );
}
