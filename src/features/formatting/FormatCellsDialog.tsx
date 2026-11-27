import { useId, useLayoutEffect, useRef, useState } from 'react';
import { formatCellValue, formatNumber } from '../../core/formats/cell-format';
import { isNumberFormat, withoutLocale } from '../../core/formats/format-code';
import { formatLocale, formatLocales, explicitFormatLocale } from '../../core/formats/format-locales';
import { categories, currencies, currencyChoices, formatName, localeForCategory, numericFormat, numericOptions,
  presets, specialLocales, typeChoices, type CurrencyChoice, type NumberCategory } from '../../core/formats/number-catalog';
import { documentLimits, type CellDocument, type CellStyle } from '../../core/workbook';
import { isFontFamily, normalFont, parseFontSize, updateAppearance, type AppearancePatch } from '../../core/formats/appearance';
import { AlignmentTab, alignmentValue } from './AlignmentTab';
import type { AlignmentPatch } from '../../application/alignment-commands';
import { BorderTab } from './BorderTab';
import type { BorderEditorState } from '../../application/border-commands';
import type { BorderDraft } from '../../core/formats/borders';
import { FillTab, FontTab } from './AppearanceTabs';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import { DialogActions } from '../dialogs/DialogActions';

const numericCategories: ReadonlySet<NumberCategory> = new Set(['Number', 'Currency', 'Accounting', 'Percentage', 'Scientific']);
const localeCategories: ReadonlySet<NumberCategory> = new Set(['Date', 'Time', 'Special']);
const localeNames = new Intl.DisplayNames(['en'], { type: 'language' });
const localeLabel = (tag: string) => localeNames.of(tag) ?? tag;

const tabs = ['Number', 'Alignment', 'Font', 'Border', 'Fill'] as const;
const availableTabs = tabs;
export type FormatTab = typeof availableTabs[number];

function available(item: typeof tabs[number]): item is FormatTab {
  return availableTabs.some(tab => tab === item);
}

export function FormatCellsDialog({ cell, style = cell?.style, locale, usedCodes, initialTab, borderEditor, mergeEnabled, merged, onSubmit, onCancel }: {
  cell: CellDocument | undefined;
  style?: CellStyle | undefined;
  locale: string;
  usedCodes: readonly string[];
  initialTab: FormatTab;
  borderEditor: BorderEditorState | undefined;
  mergeEnabled: boolean;
  merged: boolean | undefined;
  onSubmit: (patch: AppearancePatch, code: string | undefined, borders: BorderDraft, alignment: AlignmentPatch, merge: boolean | undefined) => string | undefined;
  onCancel: () => void;
}) {
  const ref = useRef<HTMLDialogElement>(null);
  const tabId = useId();
  const [tab, setTab] = useState(initialTab);
  const [borders, setBorders] = useState<BorderDraft>({});
  const [alignment, setAlignment] = useState<AlignmentPatch>({});
  const [merge, setMerge] = useState<boolean>();
  const enabledTabs: readonly FormatTab[] = borderEditor ? availableTabs : availableTabs.filter(item => item !== 'Border');
  const [patch, setPatch] = useState<AppearancePatch>({});
  const [family, setFamily] = useState(style?.fontFamily ?? 'Segoe UI');
  const [size, setSize] = useState(String(style?.fontSize ?? 11));
  const appearance = updateAppearance(style, patch);
  const validFont = isFontFamily(family) && parseFontSize(size) !== undefined;
  function updatePatch(change: AppearancePatch) { setPatch(current => ({ ...current, ...change })); }
  const original = style?.numberFormat ?? 'General';
  const name = formatName(original);
  const [category, setCategory] = useState<NumberCategory>(name === 'Short Date' || name === 'Long Date' ? 'Date' : name);
  const [code, setCode] = useState(original);
  const [touched, setTouched] = useState(false);
  const [error, setError] = useState('');
  const [typeLocale, setTypeLocale] = useState(() => localeForCategory(category, original, locale));
  const [numericLocale, setNumericLocale] = useState(() => (explicitFormatLocale(original) ?? formatLocale(locale)).tag);
  const [number, setNumber] = useState(() => numericOptions(original, numericLocale));
  const [decimalText, setDecimalText] = useState(() => String(number.decimals));
  const hasLocale = localeCategories.has(category);
  const hasTypes = hasLocale || category === 'Fraction';
  const isNumeric = numericCategories.has(category);
  const choices = hasTypes ? typeChoices(category, typeLocale) : [];
  const match = (candidate: string) => withoutLocale(candidate).toLowerCase() === withoutLocale(code).toLowerCase();
  const choice = choices.findIndex(option => match(option.code));
  const validDecimals = /^(?:[0-9]|10)$/u.test(decimalText);
  const valid = validFont && (!isNumeric || validDecimals) && (!touched || isNumberFormat(code));
  const sample = valid ? formatCellValue(cell, code) : undefined;

  useLayoutEffect(() => {
    const dialog = ref.current!;
    if (!dialog.open) dialog.showModal();
    if (initialTab !== 'Number') dialog.querySelector<HTMLButtonElement>('[role="tab"][aria-selected="true"]')?.focus();
  }, [initialTab]);

  function updateCode(next: string) {
    if (next !== code) setTouched(true);
    setError('');
    setCode(next);
  }

  function chooseCategory(next: NumberCategory) {
    setCategory(next);
    if (next === 'General' || next === 'Text') updateCode(presets[next]);
    else if (numericCategories.has(next)) {
      const culture = (explicitFormatLocale(code) ?? formatLocale(locale)).tag;
      const options = numericOptions(code, culture);
      setNumericLocale(culture);
      setNumber(options);
      setDecimalText(String(options.decimals));
      updateCode(numericFormat(next, options.decimals, options.grouping, options.currency, options.negative, culture));
    } else if (localeCategories.has(next) || next === 'Fraction') {
      const culture = localeForCategory(next, code, locale);
      setTypeLocale(culture);
      const options = typeChoices(next, culture);
      updateCode((options.find(option => match(option.code)) ?? options[0]!).code);
    }
  }

  function adjustNumber(next: typeof number) {
    setNumber(next);
    updateCode(numericFormat(category, next.decimals, next.grouping, next.currency, next.negative, numericLocale));
  }

  return <dialog ref={ref} className="paste-dialog number-dialog" aria-labelledby={tabId + '-title'}
    onCancel={event => { if (event.target === event.currentTarget) { event.preventDefault(); onCancel(); } }}>
    <form noValidate onSubmit={event => { event.preventDefault(); if (valid) setError(onSubmit(patch, touched ? code : undefined, borders, alignment, merge) ?? ''); }}>
      <DialogTitleBar>
        <h2 id={tabId + '-title'}>Format Cells</h2>
        <button type="button" className="format-dialog-close" aria-label="Close Format Cells" onClick={onCancel}>×</button>
      </DialogTitleBar>
      <div className="appearance-tabs" role="tablist" aria-label="Format Cells options">
        {tabs.map(item => available(item) && enabledTabs.includes(item) ? <button type="button" key={item} role="tab" id={tabId + '-' + item}
          aria-controls={tabId + '-panel-' + item} aria-selected={tab === item} tabIndex={tab === item ? 0 : -1}
          onClick={() => setTab(item)} onKeyDown={event => {
            const index = enabledTabs.indexOf(item);
            const last = enabledTabs.length - 1;
            const next = event.key === 'ArrowLeft' ? (index + last) % enabledTabs.length :
              event.key === 'ArrowRight' ? (index + 1) % enabledTabs.length : event.key === 'Home' ? 0 : event.key === 'End' ? last : -1;
            if (next < 0) return;
            event.preventDefault();
            const nextTab = enabledTabs[next]!;
            setTab(nextTab);
            document.getElementById(tabId + '-' + nextTab)?.focus();
          }}>{item}</button> : <button key={item} type="button" role="tab" disabled aria-selected={false}
            title={item === 'Border' ? 'Select one contiguous area for borders' : 'Alignment formatting is unavailable'}>{item}</button>)}
      </div>
      <div className="format-dialog-panes">
        <div role="tabpanel" id={tabId + '-panel-Number'} aria-labelledby={tabId + '-Number'} hidden={tab !== 'Number'} className="number-dialog-body">
          <label>Category
            <select size={categories.length} autoFocus={initialTab === 'Number'} value={category} onChange={event => chooseCategory(event.currentTarget.value as NumberCategory)}>
              {categories.map(item => <option key={item}>{item}</option>)}
            </select>
          </label>
          <div className="number-dialog-fields">
            <div className="number-format-preview">Sample
              <output aria-label="Number format sample" style={sample?.color ? { color: '#' + sample.color } : undefined}>{sample?.text}</output>
            </div>
            {isNumeric && <>
              <label>Decimal places<input type="number" min="0" max="10" value={decimalText} required onChange={event => {
                const text = event.currentTarget.value;
                setDecimalText(text);
                if (/^(?:[0-9]|10)$/u.test(text)) adjustNumber({ ...number, decimals: Number(text) });
              }} /></label>
              {category === 'Number' && <label className="number-checkbox">
                <input type="checkbox" checked={number.grouping}
                  onChange={event => adjustNumber({ ...number, grouping: event.currentTarget.checked })} />
                Use 1000 Separator ({formatLocale(numericLocale).group})
              </label>}
              {(category === 'Currency' || category === 'Accounting') && <label>Symbol
                <select value={number.currency} onChange={event => adjustNumber({ ...number, currency: event.currentTarget.value as CurrencyChoice })}>
                  <option value="None">None</option>
                  {currencies.map(currency => <option key={currency} value={currency}>{currencyChoices[currency]}</option>)}
                </select>
              </label>}
              {(category === 'Number' || category === 'Currency') && <label>Negative numbers
                <select size={4} value={number.negative} onChange={event => adjustNumber({ ...number, negative: Number(event.currentTarget.value) })}>
                  {[0, 1, 2, 3].map(index => {
                    const example = formatNumber(-1234.1,
                      numericFormat(category, number.decimals, number.grouping, number.currency, index, numericLocale));
                    return <option key={index} value={index}>{example.text}{example.color && ' (red)'}</option>;
                  })}
                </select>
              </label>}
            </>}
            {hasTypes && <label>{category === 'Special' ? 'Special type' : category === 'Fraction' ? 'Fraction type' : 'Date/time type'}
              <select size={Math.min(7, choices.length)} value={choice < 0 ? '' : choice} onChange={event => {
                const option = choices[Number(event.currentTarget.value)];
                if (option) updateCode(option.code);
              }}>
                {choice < 0 && <option value="" disabled>Current code</option>}
                {choices.map((option, index) => <option key={index} value={index}>{option.label}</option>)}
              </select>
            </label>}
            {hasLocale && <label>Locale (location)
              <select value={typeLocale} onChange={event => {
                const culture = event.currentTarget.value;
                setTypeLocale(culture);
                const options = typeChoices(category, culture);
                updateCode(options[Math.max(0, Math.min(choice, options.length - 1))]!.code);
              }}>
                {(category === 'Special' ? specialLocales : formatLocales.map(item => item.tag)).map(tag =>
                  <option key={tag} value={tag}>{localeLabel(tag)}</option>)}
              </select>
            </label>}
            {category === 'Custom' && <>
              <label>Type<input maxLength={documentLimits.numberFormatLength} value={code} onChange={event => updateCode(event.currentTarget.value)} /></label>
              <label>Existing formats<select size={7} value={code} onChange={event => updateCode(event.currentTarget.value)}>
                {[...new Set([...Object.values(presets), ...usedCodes, original])].map(item => <option key={item}>{item}</option>)}
              </select></label>
            </>}
          </div>
        </div>
        <div role="tabpanel" id={tabId + '-panel-Alignment'} aria-labelledby={tabId + '-Alignment'} hidden={tab !== 'Alignment'}>
          <AlignmentTab style={alignmentValue(style, alignment)} mergeEnabled={mergeEnabled} merged={merge ?? merged}
            onMerge={setMerge} onChange={change => setAlignment(current => ({ ...current, ...change }))} />
        </div>
        <div role="tabpanel" id={tabId + '-panel-Font'} aria-labelledby={tabId + '-Font'} hidden={tab !== 'Font'}>
          {tab === 'Font' && <FontTab style={appearance} family={family} size={size} setFamily={text => {
            setFamily(text);
            if (isFontFamily(text)) updatePatch({ fontFamily: text.trim() });
          }} setSize={text => {
            setSize(text);
            const points = parseFontSize(text);
            if (points !== undefined) updatePatch({ fontSize: points });
          }} onChange={updatePatch} onNormal={() => {
            setFamily('Segoe UI');
            setSize('11');
            updatePatch(normalFont);
          }} />}
        </div>
        {borderEditor && <div role="tabpanel" id={tabId + '-panel-Border'} aria-labelledby={tabId + '-Border'} hidden={tab !== 'Border'}>
          <BorderTab initial={borderEditor} draft={borders} onChange={setBorders} />
        </div>}
        <div role="tabpanel" id={tabId + '-panel-Fill'} aria-labelledby={tabId + '-Fill'} hidden={tab !== 'Fill'}>
          {tab === 'Fill' && <FillTab style={appearance} onChange={updatePatch} />}
        </div>
        {(!valid || error) && <p role="alert">{error || (!validFont ? 'Enter a font name and a size from 6 to 72 points.' :
          isNumeric && !validDecimals ? 'Enter decimal places from 0 to 10.' : 'Enter a valid Excel number format (1–128 characters).')}</p>}
      </div>
      <DialogActions onCancel={onCancel} disabled={!valid} />
    </form>
  </dialog>;
}
