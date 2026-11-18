import { useLayoutEffect, useRef, useState, type PointerEvent } from 'react';
import type { AlignmentPatch } from '../../application/alignment-commands';
import type { CellStyle } from '../../core/workbook';

const horizontal = [['', 'General'], ['left', 'Left (Indent)'], ['center', 'Center'], ['right', 'Right (Indent)'],
  ['fill', 'Fill'], ['justify', 'Justify'], ['centerContinuous', 'Center Across Selection'], ['distributed', 'Distributed (Indent)']] as const;
const vertical = [['top', 'Top'], ['center', 'Center'], ['', 'Bottom'], ['justify', 'Justify'], ['distributed', 'Distributed']] as const;
const directions = [['', 'Context'], ['ltr', 'Left-to-Right'], ['rtl', 'Right-to-Left']] as const;

function Choice({ label, name, options, value, onChange }: {
  label: string; name: string; options: readonly (readonly [string, string])[]; value?: string | undefined;
  onChange: (value: string | undefined) => void;
}) {
  return <label>{label}<select aria-label={name} value={value ?? ''} onChange={event => onChange(event.currentTarget.value || undefined)}>
    {options.map(([key, text]) => <option key={key} value={key}>{text}</option>)}
  </select></label>;
}

export function AlignmentTab({ style, onChange, merged, mergeEnabled, onMerge }: {
  style: AlignmentPatch; onChange: (patch: AlignmentPatch) => void;
  merged: boolean | undefined; mergeEnabled: boolean; onMerge: (merge: boolean) => void;
}) {
  const toggle = useRef<HTMLInputElement>(null);
  useLayoutEffect(() => { toggle.current!.indeterminate = mergeEnabled && merged === undefined; }, [merged, mergeEnabled]);
  const align = style.alignment ?? '';
  const stacked = style.textRotation === 255;
  const [angle, keepAngle] = useState(stacked ? 0 : style.textRotation ?? 0);
  const degrees = stacked ? angle : style.textRotation ?? 0;
  const [degreeText, setDegreeText] = useState<string>();
  const orient = !['fill', 'centerContinuous', 'distributed'].includes(align);
  function rotate(value: number) {
    const next = Math.max(-90, Math.min(90, Math.round(value)));
    setDegreeText(undefined);
    keepAngle(next);
    onChange({ textRotation: next || undefined });
  }
  function moveDial(event: PointerEvent<SVGSVGElement>) {
    const bounds = event.currentTarget.getBoundingClientRect();
    const dx = (event.clientX - bounds.left) * 132 / bounds.width - 10;
    const dy = 90 - (event.clientY - bounds.top) * 180 / bounds.height;
    rotate(Math.atan2(dy, Math.max(0, dx)) * 180 / Math.PI);
  }
  function checked(field: 'wrap' | 'shrinkToFit' | 'justifyLastLine', label: string, disabled = false) {
    return <label className="alignment-check"><input type="checkbox" checked={style[field] ?? false} disabled={disabled}
      onChange={event => onChange({ [field]: event.currentTarget.checked || undefined,
        ...(field === 'justifyLastLine' && event.currentTarget.checked && { indent: undefined }) })} />{label}</label>;
  }

  return <div className="alignment-tab">
    <fieldset className="alignment-text"><legend>Text alignment</legend>
      <div className="alignment-row">
        <Choice label="Horizontal:" name="Horizontal alignment" options={horizontal} value={style.alignment}
          onChange={alignment => onChange({ alignment })} />
        <label>Indent:<input aria-label="Indent" type="number" min="0" max="250" value={style.indent ?? 0}
          disabled={!['left', 'right', 'distributed'].includes(align) || style.justifyLastLine}
          onChange={event => onChange({ indent: Math.min(250, Math.max(0, Math.trunc(+event.currentTarget.value))) || undefined })} /></label>
      </div>
      <div className="alignment-row">
        <Choice label="Vertical:" name="Vertical alignment" options={vertical}
          value={style.verticalAlignment === 'bottom' ? undefined : style.verticalAlignment} onChange={verticalAlignment => onChange({ verticalAlignment })} />
      </div>
      {checked('justifyLastLine', 'Justify distributed', align !== 'distributed')}
    </fieldset>
    <div className="alignment-lower">
      <fieldset><legend>Text control</legend>
        {checked('wrap', 'Wrap text')}
        {checked('shrinkToFit', 'Shrink to fit', !!style.wrap || ['fill', 'justify', 'distributed'].includes(align))}
        <label className="alignment-check"><input ref={toggle} type="checkbox" checked={merged === true} disabled={!mergeEnabled}
          onChange={event => onMerge(event.currentTarget.checked)} />Merge cells</label>
      </fieldset>
      <fieldset><legend>Right-to-left</legend>
        <Choice label="Text direction:" name="Text direction" options={directions} value={style.readingOrder}
          onChange={readingOrder => onChange({ readingOrder })} />
      </fieldset>
    </div>
    <fieldset className="alignment-orientation"><legend>Orientation</legend>
      <div className="orientation-controls">
        <button type="button" aria-label="Vertical Text" aria-pressed={stacked} disabled={!orient} onClick={() => {
          if (!stacked) keepAngle(degrees);
          onChange({ textRotation: stacked ? degrees || undefined : 255 });
        }}>T<br />e<br />x<br />t</button>
        <svg className="orientation-dial" viewBox="0 0 132 180" aria-hidden="true" data-disabled={!orient || stacked}
          onPointerDown={event => {
            if (!orient || stacked || event.button !== 0) return;
            event.currentTarget.setPointerCapture(event.pointerId);
            moveDial(event);
          }}
          onPointerMove={event => {
            if (orient && !stacked && event.currentTarget.hasPointerCapture(event.pointerId)) moveDial(event);
          }}>
          {Array.from({ length: 13 }, (_, i) => {
            const rad = (i - 6) * Math.PI / 12;
            const radius = i % 3 === 0 ? 2.5 : 1;
            return <circle key={i} cx={11 + Math.cos(rad) * 104} cy={90 - Math.sin(rad) * 80} r={radius} fill="currentColor" />;
          })}
          <g transform={'rotate(' + -degrees + ' 10 90)'}>
            <path d="M10 90H112" className="orientation-needle" /><text x="10" y="87" fill="currentColor">Text</text>
          </g>
        </svg>
      </div>
      <label className="orientation-degrees"><input aria-label="Degrees" type="number" min="-90" max="90" value={degreeText ?? degrees}
        disabled={!orient || stacked} onBlur={() => setDegreeText(undefined)} onChange={event => {
          const entered = event.currentTarget.value;
          setDegreeText(entered);
          const number = Number(entered);
          if (entered && Number.isInteger(number) && number >= -90 && number <= 90) {
            keepAngle(number);
            onChange({ textRotation: number || undefined });
          }
        }} />Degrees</label>
    </fieldset>
  </div>;
}

export function alignmentValue(style: CellStyle | undefined, patch: AlignmentPatch): AlignmentPatch {
  return { ...style, ...patch };
}
