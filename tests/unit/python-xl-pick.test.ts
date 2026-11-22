import { expect, test } from 'vitest';
import { insertXlReference, xlReferenceSlot } from '../../src/features/python/xl-pick';

test('recognizes only a caret inside an xl reference literal', () => {
  expect(xlReferenceSlot('xl("")', 4)).toEqual({ start: 4, end: 4 });
  expect(xlReferenceSlot('xl("A1:B2").sum()', 6)).toEqual({ start: 4, end: 9 });
  expect(xlReferenceSlot("xl('A1')", 5)).toEqual({ start: 4, end: 6 });
  for (const [source, offset] of [
    ['xl("")', 0], ['xl("")', 3], ['xl("")', 5], ['value = "A1"', 10],
    ['xl("A1\nB2")', 6], ['no reference here', 9],
  ] as const) {
    expect(xlReferenceSlot(source, offset), source + ' @ ' + offset).toBeUndefined();
  }
});

test('ignores source-shaped text in comments, strings, f-strings and attributes', () => {
  const cases = [
    `# xl("A1")`,
    `message = "xl('A1')"`,
    `message = '''xl("A1")'''`,
    `message = f"{xl('A1')}"`,
    `obj.xl("A1")`,
  ];
  for (const source of cases) {
    expect(xlReferenceSlot(source, source.indexOf('A1') + 1), source).toBeUndefined();
  }
});

test('replaces only the literal, escapes the active quote and leaves the caret after the closing quote', () => {
  const edit = insertXlReference('xl("A1:B2").sum()', 6, 'C3:D4')!;
  expect(edit.text).toBe('xl("C3:D4").sum()');
  expect(edit.text[edit.caret]).toBe(')');
  expect(insertXlReference('xl("")', 4, 'C3:D4')).toEqual({ text: 'xl("C3:D4")', caret: 10 });

  const crossSheet = insertXlReference('xl("")', 4, "'O''Brien'!A1")!;
  expect(crossSheet.text).toBe(`xl("'O''Brien'!A1")`);
  expect(xlReferenceSlot(crossSheet.text, crossSheet.caret)).toBeUndefined();
  const escaped = "xl('" + "\\'O\\'\\'Brien\\'!A1" + "')";
  expect(insertXlReference("xl('')", 4, "'O''Brien'!A1")).toEqual({ text: escaped, caret: 22 });
  expect(insertXlReference('not an xl call', 4, 'A1')).toBeUndefined();
});

test('picks through escaped literals and normalizes raw or triple-quoted literals safely', () => {
  const escaped = "xl('" + "\\'O\\'\\'Brien\\'!A1" + "')";
  const escapedCaret = escaped.indexOf('A1') + 1;
  expect(xlReferenceSlot(escaped, escapedCaret)).toBeDefined();
  expect(insertXlReference(escaped, escapedCaret, 'C3:D4')?.text).toBe(`xl('C3:D4')`);

  const raw = `xl(r'A1')`;
  expect(insertXlReference(raw, raw.indexOf('A1') + 1, "'O''Brien'!A1")?.text)
    .toBe(`xl("'O''Brien'!A1")`);

  const triple = 'xl("""A1""")';
  expect(insertXlReference(triple, triple.indexOf('A1') + 1, 'A1:B2')?.text)
    .toBe(`xl("A1:B2")`);
});
