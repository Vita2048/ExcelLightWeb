import { expect, it } from 'vitest';
import { formatNumber } from '../../src/core/formats/cell-format';
import { isNumberFormat } from '../../src/core/formats/format-code';
import { explicitFormatLocale, formatLocales, withFormatLocale } from '../../src/core/formats/format-locales';
import { changeDecimals, commaStyle, formatName, presets, specialLocales, typeChoices } from '../../src/core/formats/number-catalog';
import { parseTemporalEntry, temporalEntry } from '../../src/core/formats/temporal';

it('validates catalog codes and preserves the selected LCID across roundtrips', () => {
  for (const [category, code] of Object.entries(presets)) {
    expect(isNumberFormat(code), code).toBe(true);
    expect(formatName(code), code).toBe(category);
  }
  for (const { tag } of formatLocales) {
    for (const category of ['Date', 'Time'] as const) {
      for (const { code } of typeChoices(category, tag)) {
        expect(isNumberFormat(code), tag + ': ' + code).toBe(true);
        expect(explicitFormatLocale(code)?.tag).toBe(tag);
      }
    }
  }
  for (const tag of specialLocales) {
    for (const { code } of typeChoices('Special', tag)) {
      expect(isNumberFormat(code), code).toBe(true);
      expect(formatName(code), code).toBe('Special');
      expect(changeDecimals(code, 1), code).toBe(code);
    }
  }
  expect(withFormatLocale('[$-405]dd.mm.yyyy', 'en-US')).toBe('[$-409]dd.mm.yyyy');
  expect(withFormatLocale('0" [$-405]"', 'en-US')).toBe('[$-409]0" [$-405]"');
});

it('rewrites decimal precision without changing grouping, literals, section conditions and exponents', () => {
  expect(changeDecimals(undefined, 1)).toBe('0.0');
  expect(changeDecimals('#,##0.00;[Red](#,##0.00)', 1)).toBe('#,##0.000;[Red](#,##0.000)');
  expect(changeDecimals('0.00E+00', -1)).toBe('0.0E+00');
  expect(changeDecimals('0".00"\\0', 1)).toBe('0.0".00"\\0');
  expect(changeDecimals('[>=10]0.0;[Red](0.0)', -1)).toBe('[>=10]0;[Red](0)');
  expect(changeDecimals(commaStyle, 1)).toBe('_(* #,##0.000_);_(* (#,##0.000);_(* "-"??_);_(@_)');
  for (const code of ['# ?/?', '[$-405]dd.mm.yyyy', '[h]:mm:ss', '@', '[$-409]00000']) {
    expect(changeDecimals(code, 1)).toBe(code);
  }
});

it('rejects malformed format codes and safely formats zero in scientific notation', () => {
  for (const code of ['', '0\n', '0.0"missing', '0\\', '0[Red', '0;0;0;@;0', '0'.repeat(129)]) {
    expect(isNumberFormat(code), code).toBe(false);
  }
  expect(formatNumber(0, '0.00E+00')).toEqual({ text: '0.00E+00' });
});

it('round-trips leap-day and localized temporal entry, rejects overflow, and retains hidden dates', () => {
  const csDate = '[$-405]dd.mm.yyyy';
  const usTime = '[$-409]h:mm:ss';
  expect(parseTemporalEntry('29.02.1900 12:00:00', csDate)).toBe(60.5);
  expect(parseTemporalEntry('02/29/1900', usTime)).toBe(60);
  expect(parseTemporalEntry('1900-02-29T06:00', csDate)).toBe(60.25);
  expect(parseTemporalEntry('11.01.1900 6:00:00', csDate)).toBe(11.25);
  expect(parseTemporalEntry('3/14/2012', usTime)).toBe(40982);
  expect(parseTemporalEntry('14. března 2012', csDate)).toBe(40982);
  expect(parseTemporalEntry('14 March 2012', usTime)).toBe(40982);
  expect(parseTemporalEntry('29.02.1900 6:00:00', usTime)).toBe(60.25);
  const bulgarian = withFormatLocale('dd.mm.yyyy', 'bg-BG');
  expect(parseTemporalEntry(temporalEntry(40982.25, bulgarian)!, bulgarian)).toBe(40982.25);
  expect(parseTemporalEntry('12:30 AM', usTime)).toBe(1 / 48);
  expect(parseTemporalEntry('13:30:00', usTime)).toBe(.5625);
  expect(parseTemporalEntry('31.02.2024', csDate)).toBeUndefined();
  expect(parseTemporalEntry('25:00:00', usTime)).toBeUndefined();
  expect(temporalEntry(10, usTime)).toBe('1/10/1900 0:00:00');
  expect(temporalEntry(-1, csDate)).toBeUndefined();
  expect(temporalEntry(2_958_466, csDate)).toBeUndefined();
  expect(parseTemporalEntry('3/14/2012', '0.00" days"')).toBeUndefined();
});
