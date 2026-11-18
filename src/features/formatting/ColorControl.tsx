import { useId, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { parseColor } from '../../core/formats/appearance';

const theme = ['FFFFFF', '000000', 'E7E6E6', '1F4E79', '5B9BD5', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646'];
const standards = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];
const ribbonTheme = ['FFFFFF', '000000', 'E7E6E6', '0E2841', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E'];
const tints = [0, .68, .46, .24, .02, -.2];
function tintColor(rgb: string, tint: number) {
  return [0, 2, 4].map(index => {
    const channel = Number.parseInt(rgb.slice(index, index + 2), 16);
    const offset = tint >= 0 ? (255 - channel) * tint : channel * tint;
    return Math.round(channel + offset).toString(16).padStart(2, '0').toUpperCase();
  }).join('');
}
const themeColors = tints.flatMap(tint => theme.map(rgb => tintColor(rgb, tint)));
const ribbonColors = ribbonTheme.concat(Array.from({ length: 5 }, (_, row) => ribbonTheme.map((rgb, column) =>
  tintColor(rgb, (column === 0 ? [-.05, -.15, -.25, -.35, -.5] : column === 1 ? [.5, .35, .25, .15, .05] :
    [.8, .6, .4, -.25, -.5])[row]!))).flat());

export function ColorPalette({ value, colors = [...themeColors, ...standards], onChange }: {
  value: string | undefined;
  colors?: readonly string[];
  onChange: (color: string) => void;
}) {
  return <div className="color-palette">
    {colors.map((color, index) => <button key={index} type="button" title={'#' + color} className="color-swatch"
      aria-label={'Color #' + color} aria-pressed={value === color} style={{ background: '#' + color }}
      onClick={() => onChange(color)} />)}
  </div>;
}

export function ColorControl({ label, value, automatic, ribbon = false, display = 'label', onChange }: {
  label: string;
  value: string | undefined;
  automatic?: boolean;
  ribbon?: boolean;
  display?: 'label' | 'value' | 'bar' | 'text';
  onChange: (color?: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const button = useRef<HTMLButtonElement>(null);
  const combo = display === 'value' || display === 'bar';
  const bar = display === 'bar' || (display === 'value' && value !== undefined);
  return <>
    <button type="button" ref={button} className={'appearance-color' + (combo ? ' appearance-combo' : '')} aria-label={label}
      aria-description={combo ? value ? 'Selected color #' + value : automatic ? 'Automatic' : 'No Color' : undefined}
      aria-haspopup="dialog" aria-expanded={open} onClick={() => setOpen(true)}>
      {display !== 'text' && <span aria-hidden="true" className={'color-swatch' + (bar ? ' color-swatch-bar' : '')}
        style={{ background: value ? '#' + value : display === 'value' ? '#fff' : undefined }} />}
      {!combo ? label : !bar && (automatic ? 'Automatic' : 'No Color')}
    </button>
    {open && <ColorPicker ribbon={ribbon} label={label} value={value} automatic={automatic === true} onClose={() => {
      setOpen(false);
      button.current?.focus();
    }} onChange={onChange} />}
  </>;
}

function ColorPicker({ label, value, automatic, ribbon, onChange, onClose }: {
  label: string;
  value: string | undefined;
  automatic: boolean;
  ribbon: boolean;
  onChange: (value?: string) => void;
  onClose: () => void;
}) {
  const ref = useRef<HTMLDialogElement>(null);
  const title = useId();
  const [hex, setHex] = useState(value ? '#' + value : '');
  const rgb = parseColor(hex);
  useLayoutEffect(() => { if (!ref.current!.open) ref.current!.showModal(); }, []);
  function close() { ref.current?.close(); onClose(); }
  function choose(color?: string) { close(); onChange(color); }
  return createPortal(<dialog ref={ref} className="paste-dialog appearance-popover" aria-labelledby={title}
    onCancel={event => { event.preventDefault(); event.stopPropagation(); close(); }}>
    <h2 id={title}>{label}</h2>
    <button type="button" onClick={() => choose()}>{automatic ? 'Automatic' : ribbon ? 'No Fill' : 'No Color'}</button>
    {([['Theme colors', ribbon ? ribbonColors : themeColors], ['Standard colors', standards]] as const).map(([name, colors]) =>
      <fieldset key={name}><legend>{name}</legend><ColorPalette colors={colors} value={value} onChange={choose} /></fieldset>)}
    <div className="custom-color">
      <label>Custom color<input type="color" aria-label="Choose custom color" value={'#' + (rgb ?? '000000')}
        onChange={event => setHex(event.currentTarget.value)} /></label>
      <label>Custom color hex<input type="text" maxLength={7} placeholder="#RRGGBB" value={hex}
        onChange={event => setHex(event.currentTarget.value)}
        onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); if (rgb) choose(rgb); } }} /></label>
    </div>
    {hex && !rgb && <p role="alert">Enter six hexadecimal color digits.</p>}
    <div className="dialog-actions">
      <button type="button" disabled={!rgb} onClick={() => rgb && choose(rgb)}>Apply custom color</button>
      <button type="button" onClick={close}>Cancel</button>
    </div>
  </dialog>, document.body);
}
