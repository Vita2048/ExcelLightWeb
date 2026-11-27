import { memo, useState } from 'react';
import { mergeLabels, type AlignmentAction, type MergeMode } from '../../application/alignment-commands';
import type { CellStyle } from '../../core/workbook';
import { RibbonFlyout, RibbonGlyph } from '../shell/RibbonFlyout';

const rotations = [
  ['Angle Counterclockwise', 45, 'orientation'], ['Angle Clockwise', -45, 'clockwise'],
  ['Vertical Text', 255, 'stacked'], ['Rotate Text Up', 90, 'up'], ['Rotate Text Down', -90, 'down'],
] as const;
const mergeIcons = { center: 'merge', across: 'across', cells: 'cells', unmerge: 'unmerge' } as const;

export const AlignmentRibbon = memo(function AlignmentRibbon({ style, merged, onAlign, onMerge, onDialog }: {
  style: CellStyle | undefined; merged: boolean; onAlign: (action: AlignmentAction) => void;
  onMerge: (mode: MergeMode) => void; onDialog: () => void;
}) {
  const [last, setLast] = useState<Exclude<MergeMode, 'unmerge'>>('center');
  function choice(kind: 'horizontal' | 'vertical', label: string, value: string, icon = value) {
    const current = kind === 'horizontal' ? style?.alignment : style?.verticalAlignment ?? 'bottom';
    return <button type="button" aria-label={label} title={label} aria-pressed={current === value}
      onClick={() => onAlign({ kind, value })}><RibbonGlyph name={'alignment:' + icon} /></button>;
  }
  return <fieldset className="alignment-ribbon">
    <legend>Alignment</legend>
    <div className="alignment-ribbon-body">
      <div className="ribbon-stack">
        <div className="alignment-ribbon-row">
          {choice('vertical', 'Top Align', 'top')}
          {choice('vertical', 'Middle Align', 'center', 'middle')}
          {choice('vertical', 'Bottom Align', 'bottom')}
          <RibbonFlyout label="Orientation" caption="" icon="alignment:orientation" items={[
            ...rotations.map(([label, value, icon]) => ({ label, icon: 'alignment:' + icon, checked: style?.textRotation === value,
              select: () => onAlign({ kind: 'rotation', value }) })),
            { label: 'Format Cell Alignment', icon: 'alignment:formatOrientation', select: onDialog },
          ]} />
        </div>
        <div className="alignment-ribbon-row">
          {choice('horizontal', 'Align Left', 'left')}
          {choice('horizontal', 'Center', 'center')}
          {choice('horizontal', 'Align Right', 'right')}
          {([-1, 1] as const).map(value => <button key={value} type="button" aria-label={value < 0 ? 'Decrease Indent' : 'Increase Indent'}
            title={value < 0 ? 'Decrease Indent' : 'Increase Indent'} onClick={() => onAlign({ kind: 'indent', value })}>
            <RibbonGlyph name={'alignment:' + (value < 0 ? 'indentLess' : 'indentMore')} />
          </button>)}
        </div>
      </div>
      <div className="ribbon-stack alignment-ribbon-controls">
        <button type="button" className="ribbon-command" aria-pressed={style?.wrap ?? false} onClick={() => onAlign({ kind: 'wrap' })}>
          <RibbonGlyph name="alignment:wrap" />Wrap Text
        </button>
        <div className="ribbon-split">
          <button type="button" className="ribbon-command" aria-pressed={merged} onClick={() => onMerge(merged ? 'unmerge' : last)}>
            <RibbonGlyph name={'alignment:' + mergeIcons[last]} />{mergeLabels[last]}
          </button>
          <RibbonFlyout label="Merge options" caption="" items={(Object.keys(mergeLabels) as MergeMode[]).map(mode => ({
            label: mergeLabels[mode], icon: 'alignment:' + mergeIcons[mode], select: () => {
              if (mode !== 'unmerge') setLast(mode);
              onMerge(mode);
            },
          }))} />
        </div>
      </div>
    </div>
    <button type="button" className="ribbon-launcher" aria-label="Format Cell Alignment" title="Format Cell Alignment" onClick={onDialog}>
      <RibbonGlyph name="launcher" size={12} />
    </button>
  </fieldset>;
});
