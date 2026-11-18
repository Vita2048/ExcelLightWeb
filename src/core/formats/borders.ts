import type { CellBorder, CellStyle } from '../workbook';

export const borderStyles = {
  none: 'None', hair: 'Hairline', dotted: 'Dotted', dashed: 'Dashed', dashDot: 'Dash dot', dashDotDot: 'Dash dot dot',
  mediumDashDotDot: 'Medium dash dot dot', slantDashDot: 'Slanted dash dot', mediumDashDot: 'Medium dash dot',
  mediumDashed: 'Medium dashed', thin: 'Thin', medium: 'Medium', thick: 'Thick', double: 'Double',
} as const;
export type BorderStyle = keyof typeof borderStyles;
export const borderEdges = ['top', 'right', 'bottom', 'left', 'diagonalUp', 'diagonalDown'] as const;
export type BorderEdge = typeof borderEdges[number];
export type CellBorders = Partial<Record<BorderEdge, CellBorder>>;
export const draftEdges = ['top', 'horizontal', 'bottom', 'left', 'vertical', 'right', 'diagonalUp', 'diagonalDown'] as const;
export type DraftEdge = typeof draftEdges[number];
export type BorderDraft = Readonly<Partial<Record<DraftEdge, CellBorder | null>>>;

export function cellBorders(style?: CellStyle): Readonly<CellBorders> {
  if (style?.borderDetails) return style.borderDetails;
  const result: CellBorders = {};
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    if (style?.borders?.includes(side[0]!)) result[side] = { style: 'thin' };
  }
  return result;
}

export function setBorder(edges: CellBorders, side: BorderEdge, pen: CellBorder | null) {
  if (!pen) { delete edges[side]; return; }
  edges[side] = pen;
  if (side === 'diagonalUp' && edges.diagonalDown) edges.diagonalDown = pen;
  if (side === 'diagonalDown' && edges.diagonalUp) edges.diagonalUp = pen;
}

export function withBorders(style: CellStyle | undefined, edges: CellBorders): CellStyle {
  const { borders: _legacy, borderDetails: _details, ...rest } = style ?? {};
  return Object.keys(edges).length ? { ...rest, borderDetails: edges } : rest;
}

export const borderPresets = {
  bottom: 'Bottom Border', top: 'Top Border', left: 'Left Border', right: 'Right Border', none: 'No Border',
  all: 'All Borders', outside: 'Outside Borders', thickOutside: 'Thick Outside Borders', bottomDouble: 'Bottom Double Border',
  bottomThick: 'Thick Bottom Border', topBottom: 'Top and Bottom Border', topThickBottom: 'Top and Thick Bottom Border',
  topDoubleBottom: 'Top and Double Bottom Border',
} as const;
export const borderModes = { draw: 'Draw Border', grid: 'Draw Border Grid', erase: 'Erase Border' } as const;
export type BorderPreset = keyof typeof borderPresets;
export type BorderMode = keyof typeof borderModes;
export type BorderAction = BorderPreset | BorderMode;
export interface BorderSegment {
  readonly edge: 'top' | 'left' | 'diagonalUp' | 'diagonalDown';
  readonly row: number;
  readonly column: number;
}

export function borderPen(style: BorderStyle, color?: string): CellBorder | null {
  return style === 'none' ? null : { style, ...(color && { rgb: color }) };
}
