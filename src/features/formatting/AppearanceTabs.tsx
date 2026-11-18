import { useEffect, useId, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { backgroundPatch, fillPatterns, fillShadings, fontFamilies, fontSizes, isFontFamily, parseFontSize,
  underlineOf, underlineOptions, underlinePatch, type AppearancePatch, type FillShading, type Underline } from '../../core/formats/appearance';
import type { CellStyle } from '../../core/workbook';
import { AppearancePreview } from './AppearancePreview';
import { ColorControl, ColorPalette } from './ColorControl';

const fontStyles = ['Regular', 'Italic', 'Bold', 'Bold Italic'];
const sizes = fontSizes.map(String);
const entries = Object.entries(fillPatterns);
const families = [...fontFamilies].sort();

export function FontLists({ familyId, sizeId }: { familyId: string; sizeId: string }) {
  return <>
    <datalist id={familyId}>{fontFamilies.map(family => <option key={family} value={family} />)}</datalist>
    <datalist id={sizeId}>{fontSizes.map(size => <option key={size} value={size} />)}</datalist>
  </>;
}

function ChoiceList({ label, choices, value, onChange }: {
  label: string; choices: readonly string[]; value: string; onChange: (value: string) => void;
}) {
  const ref = useRef<HTMLSelectElement>(null);
  const selected = choices.includes(value) ? value : '';
  useLayoutEffect(() => {
    const list = ref.current!;
    const option = list.selectedOptions[0];
    if (option && option.value) {
      const bottom = option.offsetTop + option.offsetHeight;
      if (bottom > list.scrollTop + list.clientHeight) list.scrollTop = bottom - list.clientHeight;
      else if (option.offsetTop < list.scrollTop) list.scrollTop = option.offsetTop;
    }
  }, [selected]);
  return <select ref={ref} aria-label={label} size={6} value={selected} onChange={event => onChange(event.currentTarget.value)}>
    {!selected && <option value="" hidden disabled />}
    {choices.map(choice => <option key={choice}>{choice}</option>)}
  </select>;
}

export function FontTab({ style, family, size, setFamily, setSize, onChange, onNormal }: {
  style: CellStyle;
  family: string;
  size: string;
  setFamily: (value: string) => void;
  setSize: (value: string) => void;
  onChange: (patch: AppearancePatch) => void;
  onNormal: () => void;
}) {
  const { fillRgb: _background, fillDetails: _fill, ...fontAppearance } = style;
  const fontStyle = fontStyles[Number(!!style.bold) * 2 + Number(!!style.italic)]!;
  const normal = !style.bold && !style.italic && !style.underline && !style.strikethrough && !style.script && !style.textRgb &&
    (style.fontFamily === undefined || style.fontFamily === 'Segoe UI') && (style.fontSize ?? 11) === 11 &&
    family === 'Segoe UI' && size === '11';
  return <div className="appearance-fields">
    <div className="font-dialog-columns">
      <div className="appearance-choice">
        <label>Font family<input maxLength={100} value={family} aria-invalid={!isFontFamily(family)}
          onChange={event => setFamily(event.currentTarget.value)} /></label>
        <ChoiceList label="Font catalog" choices={families} value={family} onChange={setFamily} />
      </div>
      <div className="appearance-choice">
        <label>Font style<input readOnly value={fontStyle} /></label>
        <ChoiceList label="Font styles" choices={fontStyles} value={fontStyle} onChange={value => {
          const next = fontStyles.indexOf(value);
          onChange({ bold: next >= 2 ? true : undefined, italic: next % 2 ? true : undefined });
        }} />
      </div>
      <div className="appearance-choice">
        <label>Font size<input inputMode="decimal" value={size} aria-invalid={parseFontSize(size) === undefined}
          onChange={event => setSize(event.currentTarget.value)} /></label>
        <ChoiceList label="Font sizes" choices={sizes} value={size} onChange={setSize} />
      </div>
      <label>Underline<select value={underlineOf(style)}
        onChange={event => onChange(underlinePatch(event.currentTarget.value as Underline))}>
        {underlineOptions.map(option => <option key={option}>{option}</option>)}
      </select></label>
      <div className="appearance-choice"><span>Color</span>
        <ColorControl label="Font Color" display="value" value={style.textRgb} automatic onChange={textRgb => onChange({ textRgb })} />
      </div>
      <label className="normal-font"><input type="checkbox" checked={normal} disabled={normal} onChange={onNormal} /> Normal font</label>
      <fieldset className="font-effects"><legend>Effects</legend>
        <label><input type="checkbox" checked={!!style.strikethrough} onChange={event =>
          onChange({ strikethrough: event.currentTarget.checked ? true : undefined })} /> Strikethrough</label>
        {(['superscript', 'subscript'] as const).map(script => <label key={script}>
          <input type="checkbox" checked={style.script === script} onChange={event =>
            onChange({ script: event.currentTarget.checked ? script : undefined })} /> {script === 'superscript' ? 'Superscript' : 'Subscript'}
        </label>)}
      </fieldset>
      <fieldset className="font-preview"><legend>Preview</legend>
        <AppearancePreview style={{ ...fontAppearance, fontSize: style.fontSize ?? 11 }}
          label="Font sample" text={style.fontFamily ?? 'Segoe UI'} width="full" height={103} />
      </fieldset>
    </div>
  </div>;
}

export function FillTab({ style, onChange }: { style: CellStyle; onChange: (patch: AppearancePatch) => void }) {
  const [gradient, setGradient] = useState(false);
  const button = useRef<HTMLButtonElement>(null);
  const picker = useRef<HTMLDetailsElement>(null);
  const [patternsOpen, setPatternsOpen] = useState(false);
  useEffect(() => {
    if (!patternsOpen) return;
    function dismiss(event: PointerEvent) {
      const details = picker.current;
      if (details && event.target instanceof Node && !details.contains(event.target)) details.open = false;
    }
    document.addEventListener('pointerdown', dismiss, true);
    return () => document.removeEventListener('pointerdown', dismiss, true);
  }, [patternsOpen]);
  const fill = style.fillDetails;
  const { patternRgb: _ink, ...existing } = fill ?? { pattern: 'none' };
  const pattern = existing.pattern;
  const patternName = fillPatterns[pattern as keyof typeof fillPatterns];
  const patternSample = (pattern: string): CellStyle => ({ fillRgb: 'FFFFFF',
    fillDetails: { pattern, ...(fill?.patternRgb && { patternRgb: fill.patternRgb }) } });
  function closePatterns() {
    picker.current!.open = false;
    picker.current!.querySelector('summary')?.focus();
  }
  const background = (rgb?: string) => onChange(backgroundPatch(style, rgb));
  return <div className="appearance-fields">
    <div className="fill-dialog-columns">
      <div className="appearance-choice fill-background" role="group" aria-label="Background Color">
        <span>Background Color</span>
        <button type="button" onClick={() => background()}>No Color</button>
        <ColorPalette value={style.fillRgb} onChange={background} />
        <div className="appearance-actions">
          <button type="button" ref={button} onClick={() => setGradient(true)}>Fill Effects…</button>
          <ColorControl label="More Colors…" value={style.fillRgb} display="text" onChange={background} />
        </div>
      </div>
      <div className="appearance-choice">
        <span>Pattern Color</span>
        <ColorControl label="Pattern Color" value={fill?.patternRgb} display="value" automatic onChange={patternRgb =>
          onChange({ fillDetails: { ...existing, ...(patternRgb && { patternRgb }) } })} />
        <span>Pattern Style</span>
        <details ref={picker} className="pattern-picker" onToggle={event => setPatternsOpen(event.currentTarget.open)} onKeyDown={event => {
          if (event.key === 'Escape' && event.currentTarget.open) {
            event.preventDefault();
            event.stopPropagation();
            closePatterns();
          }
        }}>
          <summary className="appearance-combo" aria-label={'Pattern Style: ' + patternName}>
            <AppearancePreview style={patternSample(pattern)} width="full" height={18} />
          </summary>
          {patternsOpen && <div className="pattern-options">
            {entries.map(([code, name]) => <button type="button" key={code} title={name} aria-label={name}
              aria-pressed={pattern === code} onClick={() => {
                onChange({ fillDetails: code === 'none' ? undefined : patternSample(code).fillDetails });
                closePatterns();
              }}><AppearancePreview style={patternSample(code)} width="full" height={22} /></button>)}
          </div>}
        </details>
      </div>
    </div>
    <div className="appearance-choice">
      <span>Sample</span>
      <AppearancePreview style={{ ...style, fillRgb: style.fillRgb ?? 'FFFFFF' }} label="Fill sample" width="full" />
    </div>
    {gradient && <GradientDialog style={style} onChange={onChange} onClose={() => {
      setGradient(false);
      button.current?.focus();
    }} />}
  </div>;
}

function GradientDialog({ style, onChange, onClose }: { style: CellStyle; onChange: (patch: AppearancePatch) => void; onClose: () => void }) {
  const id = useId();
  const ref = useRef<HTMLDialogElement>(null);
  const [first, setFirst] = useState(style.fillRgb ?? 'FFFFFF');
  const [second, setSecond] = useState(style.fillDetails?.gradientRgb ?? '5B9BD5');
  const [shading, setShading] = useState<FillShading>((style.fillDetails?.shading as FillShading | undefined) ?? 'horizontal');
  const variants = (kind: FillShading) => kind === 'corner' ? 4 : kind === 'center' ? 1 : 3;
  const [variant, setVariant] = useState(Math.min(style.fillDetails?.variant ?? 0, variants(shading) - 1));
  const preview = (variant: number) => ({ fillRgb: first, fillDetails: { pattern: 'none', gradientRgb: second, shading, variant } });
  useLayoutEffect(() => { if (!ref.current!.open) ref.current!.showModal(); }, []);
  function close() { ref.current?.close(); onClose(); }
  return createPortal(<dialog ref={ref} aria-labelledby={id} className="paste-dialog gradient-dialog"
    onCancel={event => { event.preventDefault(); event.stopPropagation(); close(); }}>
    <div className="format-dialog-heading"><h2 id={id}>Fill Effects</h2>
      <button type="button" className="format-dialog-close" aria-label="Close Fill Effects" onClick={close}>×</button></div>
    <div className="appearance-choice">
      <strong>Gradient</strong>
      <span>Two colors</span>
      <span>Color 1</span>
      <ColorControl label="Color 1" value={first} display="bar" onChange={value => setFirst(value ?? 'FFFFFF')} />
      <span>Color 2</span>
      <ColorControl label="Color 2" value={second} display="bar" onChange={value => setSecond(value ?? '5B9BD5')} />
    </div>
    <div className="gradient-choices">
      <fieldset className="gradient-shading"><legend>Shading styles</legend>
        {Object.entries(fillShadings).map(([code, label]) => <label key={code}>
          <input type="radio" name={id + '-shading'} checked={shading === code} onChange={() => {
            const next = code as FillShading;
            setShading(next);
            setVariant(Math.min(variant, variants(next) - 1));
          }} /> {label}
        </label>)}
      </fieldset>
      <fieldset><legend>Variants</legend><div className="gradient-variants">
        {Array.from({ length: variants(shading) }, (_, index) => <button key={index} type="button" aria-pressed={variant === index}
          aria-label={'Gradient variant ' + (index + 1)} onClick={() => setVariant(index)}>
          <AppearancePreview style={preview(index)} width="full" height={62} />
        </button>)}
      </div></fieldset>
    </div>
    <div className="appearance-choice"><span>Sample</span>
      <AppearancePreview style={preview(variant)} label="Gradient sample" width="full" />
    </div>
    <div className="dialog-actions">
      <button className="format-primary" type="button" onClick={() => { onChange(preview(variant)); close(); }}>OK</button>
      <button type="button" onClick={close}>Cancel</button>
    </div>
  </dialog>, document.body);
}
