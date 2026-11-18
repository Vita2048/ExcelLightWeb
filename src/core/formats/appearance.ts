import type { CellStyle } from '../workbook';

export const fontFamilies = ['Segoe UI', 'Aptos', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New',
  'Georgia', 'Helvetica', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'] as const;
export const fontSizes = [6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72] as const;
export const fillPatterns = {
  none: 'None', darkGray: 'Dark gray', mediumGray: 'Medium gray', lightGray: 'Light gray', gray125: '12.5% gray',
  gray0625: '6.25% gray', darkHorizontal: 'Dark horizontal', darkVertical: 'Dark vertical', darkDown: 'Dark diagonal down',
  darkUp: 'Dark diagonal up', darkGrid: 'Dark grid', darkTrellis: 'Dark trellis', lightHorizontal: 'Light horizontal',
  lightVertical: 'Light vertical', lightDown: 'Light diagonal down', lightUp: 'Light diagonal up', lightGrid: 'Light grid',
  lightTrellis: 'Light trellis',
} as const;
export const fillShadings = {
  horizontal: 'Horizontal', vertical: 'Vertical', diagonalUp: 'Diagonal up', diagonalDown: 'Diagonal down',
  corner: 'From corner', center: 'From center',
} as const;
export type FillShading = keyof typeof fillShadings;

type AppearanceField = 'fontFamily' | 'fontSize' | 'bold' | 'italic' | 'underline' | 'doubleUnderline' | 'underlineKind' |
  'strikethrough' | 'script' | 'textRgb' | 'fillRgb' | 'fillDetails';
export type AppearancePatch = { readonly [Key in AppearanceField]?: CellStyle[Key] | undefined };

export function updateAppearance(style: CellStyle | undefined, patch: AppearancePatch): CellStyle {
  return Object.fromEntries(Object.entries({ ...style, ...patch }).filter(([, value]) => value !== undefined));
}

export function isFontFamily(text: string): boolean { return !!text.trim() && text.length <= 100 && !/\p{Cc}/u.test(text); }

export function parseFontSize(text: string): number | undefined {
  if (!/^[+]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/iu.test(text.trim())) return;
  const size = Number(text);
  return size >= 6 && size <= 72 ? size : undefined;
}

export function parseColor(text: string): string | undefined {
  return /^#?([A-F0-9]{6})$/iu.exec(text.trim())?.[1]?.toUpperCase();
}

export const underlineOptions = ['None', 'Single', 'Double', 'Single Accounting', 'Double Accounting'] as const;
export type Underline = typeof underlineOptions[number];
export function underlineOf(style: CellStyle | undefined): Underline {
  if (!style?.underline) return 'None';
  if (style.underlineKind === 'singleAccounting') return 'Single Accounting';
  if (style.underlineKind === 'doubleAccounting') return 'Double Accounting';
  return style.doubleUnderline ? 'Double' : 'Single';
}

export function underlinePatch(kind: Underline): AppearancePatch {
  return {
    underline: kind === 'None' ? undefined : true,
    doubleUnderline: kind === 'Double' || kind === 'Double Accounting' ? true : undefined,
    underlineKind: kind === 'Single Accounting' ? 'singleAccounting' : kind === 'Double Accounting' ? 'doubleAccounting' : undefined,
  };
}

export const normalFont: AppearancePatch = {
  fontFamily: undefined, fontSize: 11, bold: undefined, italic: undefined, ...underlinePatch('None'),
  strikethrough: undefined, script: undefined, textRgb: undefined,
};

export function backgroundPatch(style: CellStyle | undefined, rgb?: string): AppearancePatch {
  return { fillRgb: rgb, fillDetails: rgb && !style?.fillDetails?.shading ? style?.fillDetails : undefined };
}
