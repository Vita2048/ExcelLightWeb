import type { FigurePlacement, PythonFigure, SheetDocument, WorkbookDocument } from '../core/workbook';
import { filteredRows } from '../application/data-commands';

const maxSide = 8192;
const maxPixels = 16_000_000;
const defaultColumnWidth = 96;
const defaultRowHeight = 26;

export interface FigureCanvas {
  readonly width: number;
  readonly height: number;
  readonly dpi: number;
}

function axisSize(first: number, count: number, dimensions: SheetDocument['rows'], base: number,
  additionalHidden?: ReadonlySet<number>): number {
  const byIndex = new Map(dimensions.map(item => [item.index, item]));
  let size = 0;
  for (let index = first; index < first + count; index++) {
    const dimension = byIndex.get(index);
    if (dimension?.hidden || additionalHidden?.has(index)) continue;
    size += dimension?.size ?? base;
  }
  return size;
}

export function figurePlacementSize(placement: FigurePlacement, sheet: SheetDocument): { width: number; height: number } {
  const hidden = filteredRows(sheet);
  return {
    width: placement.width && placement.width > 0 ? placement.width :
      axisSize(placement.column, placement.columns, sheet.columns, defaultColumnWidth),
    height: placement.height && placement.height > 0 ? placement.height :
      axisSize(placement.row, placement.rows, sheet.rows, defaultRowHeight, hidden),
  };
}

export function figureCanvasForPlacement(placement: FigurePlacement, sheet: SheetDocument,
  pixelRatio = 1): FigureCanvas {
  const { width, height } = figurePlacementSize(placement, sheet);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    throw new Error('Figure placement has invalid render dimensions.');
  }
  const scale = Math.min(1.5, Math.max(0.1, pixelRatio), maxSide / width, maxSide / height,
    Math.sqrt(maxPixels / (width * height)));
  const pixelWidth = Math.max(1, Math.min(maxSide, Math.round(width * scale)));
  const pixelHeight = Math.max(1, Math.min(maxSide, Math.round(height * scale)));
  if (pixelWidth * pixelHeight > maxPixels) throw new Error('Figure placement exceeds the render pixel limit.');
  return { width: pixelWidth, height: pixelHeight, dpi: 96 * scale };
}

export async function figureCacheKey(figure: PythonFigure): Promise<string> {
  if (figure.figureId && /^[0-9A-F]{64}$/i.test(figure.figureId)) return figure.figureId.toUpperCase();
  const binary = atob(figure.pngBase64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest), value => value.toString(16).padStart(2, '0')).join('').toUpperCase();
}

export function forEachPlacedFigure(document: WorkbookDocument, visit: (ownerId: string,
  figure: PythonFigure, placement: FigurePlacement, sheet: SheetDocument) => void): void {
  for (const sourceSheet of document.sheets) {
    for (const cell of sourceSheet.cells) {
      if (cell.content.kind !== 'python' || !cell.content.block?.lastFigure) continue;
      for (const placement of cell.content.block.placements ?? []) {
        const sheet = document.sheets.find(item => item.id === placement.sheetId);
        if (sheet) visit(cell.content.block.id, cell.content.block.lastFigure, placement, sheet);
      }
    }
  }
}
