import { createRoot } from 'react-dom/client';
import { App } from '../../src/app/App';
import { emptySheet, newWorkbook, type WorkbookDocument } from '../../src/core/workbook';
import '../../src/styles/tokens.css';
import '../../src/styles/app.css';

const fresh = newWorkbook(() => crypto.randomUUID());
const base = fresh.sheets[0]!;
const query = new URLSearchParams(location.search);
const merged = query.has('merged');
const frozenPanes = query.has('frozen') ? { rows: 1, columns: 1 } : base.frozenPanes;
const fixture: WorkbookDocument = { ...fresh, sheets: [
  { ...base, name: 'Inputs', frozenPanes, rows: [{ index: 4, size: 38 }, { index: 5, hidden: true }],
    columns: [{ index: 3, size: 112 }],
    ...(merged && { merges: [{ row: 1, column: 1, rows: 2, columns: 2 }] }),
    cells: merged ? [{ id: 'merge', row: 1, column: 1, content: { kind: 'text', value: 'Anchor' }, style: { fillRgb: '156082' } }] : [
      { id: 'first', row: 1, column: 1, content: { kind: 'number', value: 8 }, style: { fillRgb: 'FFFFFF' } },
      { id: 'second', row: 2, column: 2, content: { kind: 'formula', source: '=A1*2' }, cache: { kind: 'number', value: 16 } },
      { id: 'last', row: 1000, column: 26, content: { kind: 'text', value: 'Last cell' } },
    ] },
  { ...emptySheet('totals', 'Totals'), cells: [
    { id: 'summary', row: 1, column: 1, content: { kind: 'text', value: 'Summary' } },
  ] },
] };

createRoot(document.getElementById('root')!).render(<App initialWorkbook={fixture} />);
