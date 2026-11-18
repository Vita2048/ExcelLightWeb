import type { CellDocument } from '../../core/workbook';
import { formatCellValue } from '../../core/formats/cell-format';
import { commaStyle, currencyChoices, currencies, currencyFormat, formatName, presets } from '../../core/formats/number-catalog';
import type { NumberFormatChange } from '../../application/number-format-command';

export function NumberRibbon({ cell, onFormat, onDialog }: {
  cell: CellDocument | undefined;
  onFormat: (change: NumberFormatChange) => void;
  onDialog: () => void;
}) {
  return <fieldset className="number-ribbon">
    <legend>Number</legend>
    <select aria-label="Number Format" value="" onChange={event => {
      const code = event.currentTarget.value;
      if (code === 'more') onDialog();
      else onFormat({ code });
    }}>
      <option value="" disabled>{formatName(cell?.style?.numberFormat)}</option>
      {Object.entries(presets).map(([label, code]) => {
        const sample = formatCellValue(cell, code).text;
        return <option key={code} value={code}>{label}{sample && label !== 'General' && '  ·  ' + sample}</option>;
      })}
      <option value="more">More Number Formats…</option>
    </select>
    <select aria-label="Currency" value="" onChange={event =>
      onFormat({ code: currencyFormat(event.currentTarget.value as keyof typeof currencyChoices) })}>
      <option value="" disabled>Currency</option>
      {currencies.map(currency => <option key={currency} value={currency}>{currencyChoices[currency]}</option>)}
    </select>
    <button type="button" title="Percentage" aria-label="Percentage" onClick={() => onFormat({ code: presets.Percentage })}>%</button>
    <button type="button" title="Comma Style" aria-label="Comma Style" onClick={() => onFormat({ code: commaStyle })}>,</button>
    <button type="button" title="Increase Decimal" aria-label="Increase Decimal" onClick={() => onFormat({ decimals: 1 })}>.0→.00</button>
    <button type="button" title="Decrease Decimal" aria-label="Decrease Decimal" onClick={() => onFormat({ decimals: -1 })}>.00→.0</button>
    <button type="button" title="Format Cells" aria-label="Format Cells" onClick={onDialog}>↘</button>
  </fieldset>;
}
