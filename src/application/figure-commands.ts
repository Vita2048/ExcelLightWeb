import type { FigurePlacement, WorkbookDocument } from '../core/workbook';
import { validateWorkbook } from '../core/workbook-validation';

function ownerCell(document: WorkbookDocument, ownerId: string) {
  for (const sheet of document.sheets) {
    const cell = sheet.cells.find(item => item.content.kind === 'python' && item.content.block?.id === ownerId);
    if (cell?.content.kind === 'python' && cell.content.block) return { sheet, cell, block: cell.content.block };
  }
  throw new Error('The Python Figure source is no longer in this workbook.');
}

export function putFigurePlacement(document: WorkbookDocument, ownerId: string, placement: FigurePlacement): WorkbookDocument {
  const { sheet, cell, block } = ownerCell(document, ownerId);
  if (!block.lastFigure) throw new Error('Run a Figure result before placing it.');
  const current = block.placements ?? [];
  const existing = current.find(item => item.id === placement.id);
  if (existing && existing.sheetId !== placement.sheetId) throw new Error('A Figure placement cannot move between worksheets.');
  const placements = existing ? current.map(item => item.id === placement.id ? placement : item) : [...current, placement];
  const updated = { ...sheet, cells: sheet.cells.map(item => item.id === cell.id ? {
    ...item, content: { ...item.content, block: { ...block, placements } },
  } : item) };
  const result = { ...document, sheets: document.sheets.map(item => item.id === sheet.id ? updated : item) };
  validateWorkbook(result);
  return result;
}

export function removeFigurePlacement(document: WorkbookDocument, placementId: string): WorkbookDocument {
  let found = false;
  const sheets = document.sheets.map(sheet => {
    let changed = false;
    const cells = sheet.cells.map(cell => {
      if (cell.content.kind !== 'python' || !cell.content.block?.placements?.some(item => item.id === placementId)) return cell;
      const block = cell.content.block;
      found = changed = true;
      return { ...cell, content: { ...cell.content,
        block: { ...block, placements: block.placements!.filter(item => item.id !== placementId) } } };
    });
    return changed ? { ...sheet, cells } : sheet;
  });
  if (!found) return document;
  const result = { ...document, sheets };
  validateWorkbook(result);
  return result;
}
