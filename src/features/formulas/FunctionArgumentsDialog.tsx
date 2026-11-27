import { useEffect, useId, useLayoutEffect, useMemo, useRef } from 'react';
import { previewFormula } from '../../calculation/native-calculator';
import type { NativeFunction } from '../../calculation/formula-functions';
import type { CellReference, WorkbookDocument } from '../../core/workbook';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import { presentFormula } from './formula-presentation';
import {
  buildFunctionFormula, functionDescription, initialFunctionArguments, nextVariadicArgument,
  type FunctionArgumentSpec,
} from './function-library';

export function FunctionArgumentsDialog({
  functionName,
  document,
  target,
  separator,
  locale,
  suspended = false,
  values,
  activeIndex,
  onChangeValues,
  onChangeActiveIndex,
  onPickRange,
  onConfirm,
  onCancel,
}: {
  readonly functionName: NativeFunction;
  readonly document: WorkbookDocument;
  readonly target: CellReference;
  readonly separator: ',' | ';';
  readonly locale: string;
  readonly suspended?: boolean;
  readonly values: readonly string[];
  readonly activeIndex: number;
  readonly onChangeValues: (next: readonly string[]) => void;
  readonly onChangeActiveIndex: (index: number) => void;
  readonly onPickRange: (index: number, currentValue: string) => void;
  readonly onConfirm: (localizedFormula: string) => void;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const inputRefs = useRef<(HTMLInputElement | null)[]>([]);
  const baseId = useId();

  const specs = useMemo<readonly FunctionArgumentSpec[]>(() => {
    const list = [...initialFunctionArguments(functionName)];
    while (list.length < values.length) {
      const next = nextVariadicArgument(functionName, list);
      if (!next) break;
      list.push(next);
    }
    return list;
  }, [functionName, values.length]);

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (suspended) {
      if (dialog.open) dialog.close();
      return;
    }
    if (!dialog.open) dialog.showModal();
    const input = inputRefs.current[activeIndex] ?? inputRefs.current[0];
    input?.focus();
    input?.select();
  }, [functionName, suspended]);

  useEffect(() => {
    if (specs.length > values.length) {
      onChangeValues([...values, ...Array.from({ length: specs.length - values.length }, () => '')]);
    }
  }, [specs.length, values, onChangeValues]);

  const formulaText = useMemo(
    () => buildFunctionFormula(functionName, values, separator),
    [functionName, values, separator],
  );

  const argPreviews = useMemo(
    () => specs.map((_, index) => {
      const raw = (values[index] ?? '').trim();
      if (!raw) return { text: 'any', isError: false, empty: true };
      const preview = previewFormula(document, target, presentFormula('=' + raw, separator, ','), locale);
      if (!preview) return { text: raw, isError: false, empty: false };
      return { text: preview.text, isError: preview.isError, empty: false };
    }),
    [specs, values, document, target, separator, locale],
  );

  const formulaPreview = useMemo(
    () => previewFormula(document, target, presentFormula(formulaText, separator, ','), locale, false),
    [document, target, formulaText, separator, locale],
  );

  const activeSpec = specs[Math.min(activeIndex, Math.max(0, specs.length - 1))];

  function updateValue(index: number, nextValue: string) {
    const updated = specs.map((_, i) => (i === index ? nextValue : (values[i] ?? '')));
    if (index === specs.length - 1 && nextValue.trim() !== '') {
      const extra = nextVariadicArgument(functionName, specs);
      if (extra) updated.push('');
    }
    onChangeValues(updated);
  }

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog function-arguments-dialog"
      aria-label="Function Arguments"
      onCancel={event => {
        event.preventDefault();
        if (!suspended) onCancel();
      }}
    >
      <form
        onSubmit={event => {
          event.preventDefault();
          onConfirm(formulaText);
        }}
      >
        <DialogTitleBar>
          <h2>Function Arguments</h2>
          <button
            type="button"
            className="format-dialog-close"
            aria-label="Close Function Arguments"
            onClick={onCancel}
          >
            ×
          </button>
        </DialogTitleBar>

        <fieldset className="function-args-fieldset">
          <legend>{functionName}</legend>
          {specs.length === 0 ? (
            <p className="function-args-none">This function takes no arguments.</p>
          ) : (
            <div className="function-args-grid">
              {specs.map((spec, index) => {
                const inputId = `${baseId}-arg-${index}`;
                const preview = argPreviews[index]!;
                return (
                  <div key={`${spec.name}-${index}`} className="function-arg-row">
                    <label htmlFor={inputId} className={spec.required ? 'required' : undefined}>
                      {spec.name}
                    </label>
                    <input
                      ref={el => { inputRefs.current[index] = el; }}
                      id={inputId}
                      type="text"
                      value={values[index] ?? ''}
                      onFocus={() => onChangeActiveIndex(index)}
                      onChange={event => updateValue(index, event.currentTarget.value)}
                    />
                    <button
                      type="button"
                      className="refers-control-btn function-arg-pick"
                      aria-label={`Pick range for ${spec.name}`}
                      title="Select range on worksheet"
                      onClick={() => {
                        onChangeActiveIndex(index);
                        onPickRange(index, values[index] ?? '');
                      }}
                    >
                      ↑
                    </button>
                    <span
                      className={
                        'function-arg-preview'
                        + (preview.empty ? ' empty' : '')
                        + (preview.isError ? ' error' : '')
                      }
                    >
                      = {preview.text}
                    </span>
                  </div>
                );
              })}
            </div>
          )}
        </fieldset>

        <div className="function-args-info">
          <p className="function-args-desc">{functionDescription(functionName)}</p>
          {activeSpec && (
            <p className="function-args-hint">
              <strong>{activeSpec.name}</strong>{' '}
              {activeSpec.required ? 'is a required argument.' : 'is an optional argument.'}
            </p>
          )}
        </div>

        <div className="function-args-result" aria-live="polite">
          <span>Formula result =</span>{' '}
          <strong className={formulaPreview?.isError ? 'error' : undefined}>
            {formulaPreview ? formulaPreview.text : ''}
          </strong>
        </div>

        <DialogActions confirmLabel="OK" onCancel={onCancel} />
      </form>
    </dialog>
  );
}
