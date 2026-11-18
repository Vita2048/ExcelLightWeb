import { expect, it } from 'vitest';
import { nativeFunctions } from '../../src/calculation/formula-functions';
import { formulaHint } from '../../src/features/editing/formula-hints';

const atEnd = (source: string) => formulaHint(source, source.length);

it('provides a signature for each supported function in an incomplete draft', () => {
  for (const name of nativeFunctions) expect(atEnd('=' + name.toLowerCase() + '(')?.name).toBe(name);
  expect(atEnd('=TODAY(')).toMatchObject({ parameters: [], active: -1 });
  expect(atEnd('=UNSUPPORTED(')).toBeUndefined();
  expect(atEnd('SUM(')).toBeUndefined();
});

it('ignores separators in escaped quotes, array constants and grouped references', () => {
  const source = `=IF('A, ''(B)'!A1,"x,""y"",z",SUM({1,2;3,4},(A1,A2),`;
  expect(atEnd(source)).toMatchObject({ name: 'SUM', active: 1, variadic: true });
  expect(atEnd(source + '5)')).toMatchObject({ name: 'IF', active: 2 });
  expect(atEnd(source + '5))')).toBeUndefined();
  expect(formulaHint(source, source.indexOf('SUM'))).toMatchObject({ name: 'IF', active: 2 });
});
