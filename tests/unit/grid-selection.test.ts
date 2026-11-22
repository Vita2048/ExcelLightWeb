import { describe, expect, it } from 'vitest';
import type { CellDocument } from '../../src/core/workbook';
import { cellDisplay } from '../../src/grid/grid-selection';

const pythonAnchor: CellDocument = {
  id: 'anchor', row: 1, column: 1,
  content: { kind: 'python', blockId: 'block', source: '42\n' },
};

describe('worksheet cell presentation', () => {
  it('shows a Python result in the grid and source only when Show Formulas is on', () => {
    expect(cellDisplay(pythonAnchor)).toBe('');
    expect(cellDisplay(pythonAnchor, true)).toBe('=PY()');
    expect(cellDisplay({ ...pythonAnchor, cache: { kind: 'number', value: 42 } })).toBe('42');
    expect(cellDisplay({ ...pythonAnchor, cache: { kind: 'number', value: 42 } }, true)).toBe('=PY()');
  });
});
