import { useLayoutEffect, useRef, useState } from 'react';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';

const presets = [200, 100, 75, 50, 25] as const;

/** The desktop Zoom dialog: magnification presets, Fit selection and a 10–400% custom value. */
export function ZoomDialog({ current, onApply, onCancel }: {
  current: number;
  onApply: (zoom: number | 'fit') => void;
  onCancel: () => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  const [preset, setPreset] = useState<number | 'fit' | 'custom'>(
    presets.some(value => value === current) ? current : 'custom');
  const [custom, setCustom] = useState(String(current));
  useLayoutEffect(() => { if (!dialog.current!.open) dialog.current!.showModal(); }, []);

  function apply() {
    if (preset === 'fit') { onApply('fit'); return; }
    if (preset === 'custom') {
      const parsed = Number(custom);
      // The desktop NumberBox keeps the current zoom for an empty value and clamps typed values to 10–400.
      onApply(custom.trim() === '' || !Number.isFinite(parsed) ? current : Math.min(400, Math.max(10, Math.round(parsed))));
      return;
    }
    onApply(preset);
  }

  return <dialog ref={dialog} className="paste-dialog zoom-dialog" aria-label="Zoom"
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={event => { event.preventDefault(); apply(); }}>
      <DialogTitleBar><h2>Zoom</h2>
        <button type="button" className="format-dialog-close" aria-label="Close Zoom" onClick={onCancel}>×</button>
      </DialogTitleBar>
      <h3 className="zoom-heading">Magnification</h3>
      <div className="zoom-presets" role="radiogroup" aria-label="Magnification">
        {presets.map(value => <label key={value} className="zoom-check">
          <input type="radio" name="zoom-magnification" aria-label={value + '%'} checked={preset === value}
            onChange={() => setPreset(value)} />{value}%
        </label>)}
        <label className="zoom-check">
          <input type="radio" name="zoom-magnification" aria-label="Fit selection" checked={preset === 'fit'}
            onChange={() => setPreset('fit')} />Fit selection
        </label>
        <div className="zoom-custom">
          <label className="zoom-check">
            <input type="radio" name="zoom-magnification" aria-label="Custom" checked={preset === 'custom'}
              onChange={() => setPreset('custom')} />Custom:
          </label>
          <input type="number" aria-label="Custom zoom percentage, 10 to 400" step={1}
            value={custom} onChange={event => { setCustom(event.currentTarget.value); setPreset('custom'); }} />
          <span aria-hidden="true">%</span>
        </div>
      </div>
      <DialogActions onCancel={onCancel} />
    </form>
  </dialog>;
}
