import { useId, useState } from 'react';
import { fontSizes, isFontFamily, parseFontSize, underlineOf, underlineOptions, underlinePatch,
  type AppearancePatch, type Underline } from '../../core/formats/appearance';
import type { CellStyle } from '../../core/workbook';
import { ColorControl } from './ColorControl';
import { FontLists } from './AppearanceTabs';

function FontInput({ label, value, list, onCommit }: {
  label: string; value: string; list: string; onCommit: (text: string, focus: boolean) => void;
}) {
  const [draft, setDraft] = useState(value);
  return <input aria-label={label} title={label} list={list} value={draft} onChange={event => setDraft(event.currentTarget.value)}
    onBlur={() => { if (draft !== value) { setDraft(value); onCommit(draft, false); } }}
    onKeyDown={event => {
      if (event.key === 'Enter' || event.key === 'Escape') {
        event.preventDefault();
        const next = event.key === 'Enter' ? draft : value;
        setDraft(value);
        onCommit(next, true);
      }
    }} />;
}

export function FontRibbon({ style, onApply, onDialog, onInvalid }: {
  style: CellStyle | undefined;
  onApply: (patch: AppearancePatch, focus?: boolean) => void;
  onDialog: () => void;
  onInvalid: (message: string, focus: boolean) => void;
}) {
  const id = useId();
  const [lastUnderline, setLastUnderline] = useState<Underline>('Single');
  const [text, setText] = useState<string | undefined>('FF0000');
  const [fill, setFill] = useState<string | undefined>('FFFF00');
  const family = style?.fontFamily ?? 'Segoe UI';
  const size = style?.fontSize ?? 11;
  const underline = underlineOf(style);
  function changeSize(direction: number) {
    const next = direction > 0 ? fontSizes.find(value => value > size) ?? 72 : [...fontSizes].reverse().find(value => value < size) ?? 6;
    onApply({ fontSize: next });
  }
  return <fieldset className="font-ribbon">
    <legend>Font</legend>
    <FontLists familyId={id + '-font'} sizeId={id + '-size'} />
    <div className="font-ribbon-row">
      <FontInput key={'family-' + family} label="Font family" value={family} list={id + '-font'} onCommit={(value, focus) => {
        if (!isFontFamily(value)) onInvalid('Enter a valid font family.', focus);
        else onApply(value.trim() === family ? {} : { fontFamily: value.trim() }, focus);
      }} />
      <FontInput key={'size-' + size} label="Font size" value={String(size)} list={id + '-size'} onCommit={(value, focus) => {
        const points = parseFontSize(value);
        if (points === undefined) onInvalid('Font size must be 6 to 72 points.', focus);
        else onApply(points === size ? {} : { fontSize: points }, focus);
      }} />
      <button type="button" aria-label="Increase Font Size" title="Increase Font Size" onClick={() => changeSize(1)}>A↑</button>
      <button type="button" aria-label="Decrease Font Size" title="Decrease Font Size" onClick={() => changeSize(-1)}>A↓</button>
      <button type="button" aria-label="Format Cell Font" title="Format Cell Font" onClick={() => onDialog()}>↘</button>
    </div>
    <div className="font-ribbon-row">
      <button type="button" aria-label="Bold" aria-pressed={!!style?.bold} onClick={() => onApply({ bold: style?.bold ? undefined : true })}><b>B</b></button>
      <button type="button" aria-label="Italic" aria-pressed={!!style?.italic} onClick={() => onApply({ italic: style?.italic ? undefined : true })}><i>I</i></button>
      <button type="button" aria-label="Underline" aria-pressed={underline !== 'None'} onClick={() =>
        onApply(underlinePatch(underline === 'None' ? lastUnderline : 'None'))}><u>U</u></button>
      <select aria-label="Underline options" value="" onChange={event => {
        const choice = event.currentTarget.value as Underline;
        if (choice !== 'None') setLastUnderline(choice);
        onApply(underlinePatch(choice));
      }}><option value="" disabled>⌄</option>
        {underlineOptions.map(option => <option key={option}>{option}</option>)}
      </select>
      <button type="button" aria-label="Apply last fill color" title="Apply last fill color" onClick={() => onApply({ fillRgb: fill, fillDetails: undefined })}>■</button>
      <ColorControl label="Fill Color" value={fill} ribbon onChange={color => { setFill(color); onApply({ fillRgb: color, fillDetails: undefined }); }} />
      <button type="button" aria-label="Apply last font color" title="Apply last font color" onClick={() => onApply({ textRgb: text })}>A</button>
      <ColorControl label="Font Color" value={text} ribbon automatic onChange={color => { setText(color); onApply({ textRgb: color }); }} />
    </div>
  </fieldset>;
}
