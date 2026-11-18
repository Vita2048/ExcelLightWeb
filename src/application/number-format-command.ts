import { changeDecimals } from '../core/formats/number-catalog';
import { isNumberFormat } from '../core/formats/format-code';
import { requireDocument } from '../core/workbook-validation';
import type { EditRange } from './editing-commands';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { SetCellsCommand } from './workbook-session';
import { changeStyles, formatStyles } from './style-commands';

export type NumberFormatChange = { readonly code: string } | { readonly decimals: 1 | -1 };

export function applyNumberFormat(basis: WorkbookSnapshot, ranges: readonly EditRange[], change: NumberFormatChange): SetCellsCommand {
  if ('code' in change) return formatStyles(basis, ranges, {}, change.code);
  return changeStyles(basis, ranges, 'Format Cells', style => {
    const code = changeDecimals(style?.numberFormat, change.decimals);
    if (code === style?.numberFormat) return style;
    requireDocument(isNumberFormat(code), 'The changed number format is invalid or too long.');
    return { ...style, numberFormat: code };
  });
}
