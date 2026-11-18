import { useState, type ReactNode } from 'react';
import { fontSizes, isFontFamily, parseFontSize, underlineOf, underlineOptions, underlinePatch,
  type AppearancePatch, type Underline } from '../../core/formats/appearance';
import type { CellStyle } from '../../core/workbook';
import { ColorControl } from './ColorControl';
import { RibbonFlyout, RibbonGlyph } from '../shell/RibbonFlyout';
import { fontFamilies } from '../../core/formats/appearance';

function FontInput({ label, value, choices, onCommit }: {
  label: string; value: string; choices: readonly string[]; onCommit: (text: string, focus: boolean) => void;
}) {
  const [draft, setDraft] = useState(value);
  return <div className="ribbon-font-input"><input role="combobox" aria-expanded="false" aria-label={label} title={label} value={draft} onChange={event => setDraft(event.currentTarget.value)}
    onBlur={() => { if (draft !== value) { setDraft(value); onCommit(draft, false); } }}
    onKeyDown={event => {
      if (event.key === 'Enter' || event.key === 'Escape') {
        event.preventDefault();
        const next = event.key === 'Enter' ? draft : value;
        setDraft(value);
        onCommit(next, true);
      }
    }} /><RibbonFlyout label={label + ' options'} caption="" items={choices.map(choice => ({ label: choice, select: () => onCommit(choice, true) }))} />
  </div>;
}

export function FontRibbon({ style, onApply, onDialog, onInvalid, borders }: {
  style: CellStyle | undefined;
  borders: ReactNode;
  onApply: (patch: AppearancePatch, focus?: boolean) => void;
  onDialog: () => void;
  onInvalid: (message: string, focus: boolean) => void;
}) {
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
    <div className="font-ribbon-row">
      <FontInput key={'family-' + family} label="Font family" value={family} choices={fontFamilies} onCommit={(value, focus) => {
        if (!isFontFamily(value)) onInvalid('Enter a valid font family.', focus);
        else onApply(value.trim() === family ? {} : { fontFamily: value.trim() }, focus);
      }} />
      <FontInput key={'size-' + size} label="Font size" value={String(size)} choices={fontSizes.map(String)} onCommit={(value, focus) => {
        const points = parseFontSize(value);
        if (points === undefined) onInvalid('Font size must be 6 to 72 points.', focus);
        else onApply(points === size ? {} : { fontSize: points }, focus);
      }} />
      <button type="button" aria-label="Increase Font Size" title="Increase Font Size" onClick={() => changeSize(1)}><RibbonGlyph name="fontIncrease" /></button>
      <button type="button" aria-label="Decrease Font Size" title="Decrease Font Size" onClick={() => changeSize(-1)}><RibbonGlyph name="fontDecrease" /></button>
      <button type="button" className="ribbon-launcher" aria-label="Format Cell Font" title="Format Cell Font" onClick={() => onDialog()}><RibbonGlyph name="launcher" size={12} /></button>
    </div>
    <div className="font-ribbon-row">
      <button type="button" aria-label="Bold" aria-pressed={!!style?.bold} onClick={() => onApply({ bold: style?.bold ? undefined : true })}><RibbonGlyph name="bold" /></button>
      <button type="button" aria-label="Italic" aria-pressed={!!style?.italic} onClick={() => onApply({ italic: style?.italic ? undefined : true })}><RibbonGlyph name="italic" /></button>
      <button type="button" aria-label="Underline" aria-pressed={underline !== 'None'} onClick={() =>
        onApply(underlinePatch(underline === 'None' ? lastUnderline : 'None'))}><RibbonGlyph name={underline === 'Double' ? 'doubleUnderline' : 'underline'} /></button>
      <RibbonFlyout label="Underline options" caption="" items={underlineOptions.map(choice => ({
        label: choice, icon: choice.includes('Double') ? 'doubleUnderline' : 'underline', checked: underline === choice, select: () => {
          if (choice !== 'None') setLastUnderline(choice);
          onApply(underlinePatch(choice));
        },
      }))} />
      {borders}
      <button type="button" aria-label="Apply last fill color" title="Apply last fill color" onClick={() => onApply({ fillRgb: fill, fillDetails: undefined })}><span className="ribbon-color-icon" style={{ borderColor: fill ? '#' + fill : 'transparent' }}><RibbonGlyph name="fillColor" /></span></button>
      <ColorControl label="Fill Color" value={fill} ribbon onChange={color => { setFill(color); onApply({ fillRgb: color, fillDetails: undefined }); }} />
      <button type="button" aria-label="Apply last font color" title="Apply last font color" onClick={() => onApply({ textRgb: text })}><span className="ribbon-color-icon" style={{ borderColor: text ? '#' + text : 'currentColor' }}><RibbonGlyph name="fontColor" /></span></button>
      <ColorControl label="Font Color" value={text} ribbon automatic onChange={color => { setText(color); onApply({ textRgb: color }); }} />
    </div>
  </fieldset>;
}
