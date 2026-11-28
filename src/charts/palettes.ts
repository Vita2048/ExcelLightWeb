export interface ChartPalette {
  readonly id: string;
  readonly name: string;
  readonly colors: readonly string[];
}

export const chartPalettes: readonly ChartPalette[] = [
  {
    id: 'colorful-1',
    name: 'Colorful Palette 1',
    colors: ['#4ea1ff', '#f38c4e', '#a5aab4', '#f7c331', '#69b3e7', '#6ec071', '#3b73b9', '#b85c27'],
  },
  {
    id: 'colorful-2',
    name: 'Colorful Palette 2',
    colors: ['#56b4e9', '#e69f00', '#009e73', '#f0e442', '#cc79a7', '#0072b2', '#d55e00', '#999999'],
  },
  {
    id: 'colorful-3',
    name: 'Colorful Palette 3',
    colors: ['#ff7a5c', '#ffd166', '#06d6a0', '#4ea8de', '#9b72cf', '#f78fb3', '#63cdda', '#e77f67'],
  },
  {
    id: 'colorful-4',
    name: 'Colorful Palette 4',
    colors: ['#7eb6ff', '#76d7c4', '#f7dc6f', '#f0b27a', '#c39bd3', '#85c1e9', '#82e0aa', '#f1948a'],
  },
  {
    id: 'monochrome-1',
    name: 'Monochrome Blue',
    colors: ['#9ecbff', '#6baeff', '#4ea1ff', '#2b7de9', '#1e5cb3', '#164282'],
  },
  {
    id: 'monochrome-2',
    name: 'Monochrome Orange',
    colors: ['#ffc99e', '#ffa96b', '#f38c4e', '#d96c2c', '#a94f1c', '#7a3712'],
  },
  {
    id: 'monochrome-3',
    name: 'Monochrome Green',
    colors: ['#a8e6aa', '#7ed482', '#6ec071', '#4aa14e', '#327b36', '#215724'],
  },
  {
    id: 'monochrome-4',
    name: 'Monochrome Gray',
    colors: ['#e0e3e8', '#c0c5ce', '#a5aab4', '#868c96', '#676d77', '#4b5059'],
  },
];

export function getPaletteColors(paletteId: string): readonly string[] {
  return (chartPalettes.find(p => p.id === paletteId) ?? chartPalettes[0]!).colors;
}

export function getSeriesColor(paletteId: string, seriesIndex: number, overrideRgb?: string): string {
  if (overrideRgb) return overrideRgb;
  const colors = getPaletteColors(paletteId);
  return colors[((seriesIndex % colors.length) + colors.length) % colors.length]!;
}
