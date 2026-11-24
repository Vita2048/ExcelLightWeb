import type { CellDocument } from '../../core/workbook';
import { formatCellValue } from '../../core/formats/cell-format';
import { commaStyle, currencyChoices, currencies, currencyFormat, formatName, presets } from '../../core/formats/number-catalog';
import type { NumberFormatChange } from '../../application/number-format-command';
import { RibbonFlyout, RibbonGlyph } from '../shell/RibbonFlyout';

export function NumberRibbon({ cell, onFormat, onDialog }: {
  cell: CellDocument | undefined;
  onFormat: (change: NumberFormatChange) => void;
  onDialog: () => void;
}) {
  const tools = [['Percentage', { code: presets.Percentage }], ['Comma Style', { code: commaStyle }],
    ['Increase Decimal', { decimals: 1 }], ['Decrease Decimal', { decimals: -1 }]] as const;
  return <fieldset className="number-ribbon"><legend>Number</legend>
    <div className="ribbon-stack number-ribbon-stack">
      <RibbonFlyout label="Number Format" caption={formatName(cell?.style?.numberFormat)} items={[
        ...Object.entries(presets).map(([label, code]) => ({ label, icon: label,
          detail: label === 'General' ? 'No specific format' : formatCellValue(cell, code).text,
          select: () => onFormat({ code }) })),
        { label: 'More Number Formats…', select: onDialog },
      ]} />
      <div className="ribbon-tools">
        <RibbonFlyout label="Currency" caption="" icon="Currency" items={currencies.map(currency => ({
          label: currencyChoices[currency], icon: 'Currency', select: () => onFormat({ code: currencyFormat(currency) }),
        }))} />
        {tools.map(([label, change]) => <button type="button" key={label} title={label} aria-label={label} onClick={() => onFormat(change)}>
          <RibbonGlyph name={label} />
        </button>)}
      </div>
    </div>
    <button type="button" className="ribbon-launcher" title="Format Cells" aria-label="Format Cells" onClick={onDialog}><RibbonGlyph name="launcher" size={12} /></button>
  </fieldset>;
}
