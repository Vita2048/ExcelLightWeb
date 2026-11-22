import { expect, test } from 'vitest';
import { copyReferences } from '../../src/core/formula-references';
import {
  discoverPythonReferences, isPythonNamedReference, normalizePythonReference, renamePythonReferences, transformPythonReferences,
} from '../../src/python/references';

test('discovers literal xl calls through comments, raw strings, triple strings and escapes', () => {
  expect(discoverPythonReferences('result = xl("A1:B3").sum()')).toEqual([
    { start: 12, length: 7, reference: 'A1:B3', headers: false },
  ]);
  expect(discoverPythonReferences('xl("A1", headers=True)')).toEqual([
    { start: 3, length: 4, reference: 'A1', headers: true },
  ]);
  expect(discoverPythonReferences('xl("A1", headers=False,)')).toEqual([
    { start: 3, length: 4, reference: 'A1', headers: false },
  ]);
  expect(discoverPythonReferences("xl(\n # reference\n r'A1',\n headers=False,\n)")).toEqual([
    { start: 18, length: 5, reference: 'A1', headers: false },
  ]);
  expect(discoverPythonReferences("xl('\\u00411')")[0]?.reference).toBe('A1');
  expect(discoverPythonReferences('xl("SalesTable[#All]", headers=True)')).toEqual([
    { start: 3, length: 18, reference: 'SalesTable[#All]', headers: true },
  ]);
  expect(discoverPythonReferences('xl("TaxRate")')[0]?.reference).toBe('TaxRate');
  expect(discoverPythonReferences('xl("#REF!")')[0]?.reference).toBe('#REF!');
  // Comments and ordinary strings never create dependencies.
  expect(discoverPythonReferences("# xl('A1')\ntext = \"xl('A1')\"\nother = '''xl('A1')'''")).toEqual([]);
  expect(discoverPythonReferences('f"xl(\'A1\')"')).toEqual([]);
});

test('rejects dynamic references, aliases, unsupported headers and shadowing xl', () => {
  for (const source of [
    "xl('A'+str(1))", 'f = xl\nf(\'A1\')', "xl('A1', headers=1)", "xl('A1', True)",
    "xl('A1#')", "xl('A:A')", "xl('A1') if False else xl(ref)", "xl(f'A{1}')", "xl('A1' 'A2')",
    "obj.xl('A1')", 'xl("A1", headers=True, extra=1)', "xl(rb'A1')", 'xl("unterminated)',
  ]) {
    expect(() => discoverPythonReferences(source), source).toThrow();
  }
});

test('transforms only actual calls and preserves comments, strings, names and absolute flags', () => {
  const source = `# xl('Z9')\ntext = "xl('Z9')"\nother = '''xl('Z9')'''\n` +
    "xl('A1') + xl(\"$B$2\") + xl('TaxRate') + xl('Orders[#Data]')";
  const translated = transformPythonReferences(source, reference => copyReferences(reference, 1, 2));
  expect(translated).toContain('xl("C2")');
  expect(translated).toContain('xl("$B$2")');
  expect(translated).toContain("xl('TaxRate')");
  expect(translated).toContain("xl('Orders[#Data]')");
  expect(translated).toContain("# xl('Z9')");
  expect(translated).toContain(`other = '''xl('Z9')'''`);
  // Incomplete drafts remain saveable and are never partially rewritten.
  expect(transformPythonReferences('xl("A1', () => '=B1')).toBe('xl("A1');
});

test('renames workbook names and table selectors without touching unrelated text', () => {
  const source = "# xl('Orders')\nxl('orders[#Data]') + xl('Orders[#All]') + xl('Other')";
  expect(renamePythonReferences(source, 'Orders', 'Renamed'))
    .toBe('# xl(\'Orders\')\nxl("Renamed[#Data]") + xl("Renamed[#All]") + xl(\'Other\')');
  expect(renamePythonReferences("xl('Orders')", 'Orders', '#REF!')).toBe('xl("#REF!")');
});

test('classifies named references and normalizes binding spellings', () => {
  expect(isPythonNamedReference('TaxRate')).toBe(true);
  expect(isPythonNamedReference('Orders[#All]')).toBe(true);
  expect(isPythonNamedReference('Orders[#Data]')).toBe(true);
  expect(isPythonNamedReference('A1')).toBe(false);
  expect(isPythonNamedReference('Orders[Amount]')).toBe(false);
  expect(normalizePythonReference("='O''Brien'!$A$1:$B$2")).toBe("='O''Brien'!$A$1:$B$2");
  expect(normalizePythonReference('=A:A')).toBe('=#REF!');
  expect(normalizePythonReference('=A1#')).toBe('=#REF!');
});
