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

  it('shows the cached Python block scalar and a plot marker when the block owns the result', () => {
    const scalar: CellDocument = {
      ...pythonAnchor, content: { kind: 'python', blockId: 'block', source: '42', block: {
        id: 'block', source: '42', bindings: [], hasResult: true, lastValue: { kind: 'number', value: 42 },
      } },
    };
    const figure: CellDocument = {
      ...pythonAnchor, content: { kind: 'python', blockId: 'block', source: 'plot()', block: {
        id: 'block', source: 'plot()', bindings: [], hasResult: true, lastValue: null,
        lastFigure: { pngBase64: 'AA==', width: 1, height: 1 },
      } },
    };
    expect(cellDisplay(scalar)).toBe('42');
    expect(cellDisplay(scalar, true)).toBe('=PY()');
    expect(cellDisplay(figure)).toBe('[plot]');
    expect(cellDisplay(figure, true)).toBe('=PY()');
  });
});
