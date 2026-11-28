import type { ChartType } from '../../core/chart';

export interface ChartSubtypeOption {
  readonly type: ChartType;
  readonly label: string;
  readonly group: string;
}

export interface ChartFlyoutFamily {
  readonly triggerLabel: string;
  readonly moreLabel: string;
  readonly dialogCategory: CoreChartCategoryTab;
  readonly items: readonly ChartSubtypeOption[];
}

export type CoreChartCategoryTab = 'Column' | 'Line' | 'Pie' | 'Bar' | 'Area' | 'X Y (Scatter)';

export const coreChartFlyoutFamilies: readonly ChartFlyoutFamily[] = [
  {
    triggerLabel: 'Insert Column or Bar Chart',
    moreLabel: 'More Column Charts...',
    dialogCategory: 'Column',
    items: [
      { type: 'clusteredColumn', label: 'Clustered Column', group: '2-D Column' },
      { type: 'stackedColumn', label: 'Stacked Column', group: '2-D Column' },
      { type: 'stackedColumn100', label: '100% Stacked Column', group: '2-D Column' },
      { type: 'clusteredBar', label: 'Clustered Bar', group: '2-D Bar' },
      { type: 'stackedBar', label: 'Stacked Bar', group: '2-D Bar' },
      { type: 'stackedBar100', label: '100% Stacked Bar', group: '2-D Bar' },
    ],
  },
  {
    triggerLabel: 'Insert Line or Area Chart',
    moreLabel: 'More Line Charts...',
    dialogCategory: 'Line',
    items: [
      { type: 'line', label: 'Line', group: '2-D Line' },
      { type: 'stackedLine', label: 'Stacked Line', group: '2-D Line' },
      { type: 'stackedLine100', label: '100% Stacked Line', group: '2-D Line' },
      { type: 'lineMarkers', label: 'Line with Markers', group: '2-D Line' },
      { type: 'stackedLineMarkers', label: 'Stacked Line with Markers', group: '2-D Line' },
      { type: 'stackedLine100Markers', label: '100% Stacked Line with Markers', group: '2-D Line' },
      { type: 'area', label: 'Area', group: '2-D Area' },
      { type: 'stackedArea', label: 'Stacked Area', group: '2-D Area' },
      { type: 'stackedArea100', label: '100% Stacked Area', group: '2-D Area' },
    ],
  },
  {
    triggerLabel: 'Insert Pie or Doughnut Chart',
    moreLabel: 'More Pie Charts...',
    dialogCategory: 'Pie',
    items: [
      { type: 'pie', label: 'Pie', group: '2-D Pie' },
      { type: 'pieOfPie', label: 'Pie of Pie', group: '2-D Pie' },
      { type: 'barOfPie', label: 'Bar of Pie', group: '2-D Pie' },
      { type: 'doughnut', label: 'Doughnut', group: 'Doughnut' },
    ],
  },
  {
    triggerLabel: 'Insert Scatter (X, Y) or Bubble Chart',
    moreLabel: 'More Scatter Charts...',
    dialogCategory: 'X Y (Scatter)',
    items: [
      { type: 'scatter', label: 'Scatter', group: 'Scatter' },
      { type: 'scatterSmoothLinesMarkers', label: 'Scatter with Smooth Lines and Markers', group: 'Scatter' },
      { type: 'scatterSmoothLines', label: 'Scatter with Smooth Lines', group: 'Scatter' },
      { type: 'scatterStraightLinesMarkers', label: 'Scatter with Straight Lines and Markers', group: 'Scatter' },
      { type: 'scatterStraightLines', label: 'Scatter with Straight Lines', group: 'Scatter' },
      { type: 'bubble', label: 'Bubble', group: 'Bubble' },
    ],
  },
];

export const coreDialogCategories: Readonly<Record<CoreChartCategoryTab, readonly ChartSubtypeOption[]>> = {
  Column: coreChartFlyoutFamilies[0]!.items.filter(i => i.group === '2-D Column'),
  Line: coreChartFlyoutFamilies[1]!.items.filter(i => i.group === '2-D Line'),
  Pie: coreChartFlyoutFamilies[2]!.items,
  Bar: coreChartFlyoutFamilies[0]!.items.filter(i => i.group === '2-D Bar'),
  Area: coreChartFlyoutFamilies[1]!.items.filter(i => i.group === '2-D Area'),
  'X Y (Scatter)': coreChartFlyoutFamilies[3]!.items,
};

export function ChartCategoryListGlyph({
  category,
  size = 18,
}: {
  readonly category: CoreChartCategoryTab;
  readonly size?: number;
}) {
  const blue = '#5b9bd5';
  const gray = '#a5aab4';
  const axis = '#cfd5df';
  if (category === 'Column') {
    return (
      <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
        <path d="M2 17 H18 M3 3 V17" stroke={axis} strokeWidth="1.2" fill="none" />
        <rect x="5" y="9" width="3" height="8" fill={blue} />
        <rect x="9" y="5" width="3" height="12" fill={gray} />
        <rect x="13" y="11" width="3" height="6" fill={blue} />
      </svg>
    );
  }
  if (category === 'Line') {
    return (
      <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
        <path d="M2 17 H18 M3 3 V17" stroke={axis} strokeWidth="1.2" fill="none" />
        <polyline points="4,13 8,8 12,11 17,5" fill="none" stroke={blue} strokeWidth="1.7" />
        <polyline points="4,15 8,11 12,14 17,9" fill="none" stroke={gray} strokeWidth="1.5" />
      </svg>
    );
  }
  if (category === 'Pie') {
    return (
      <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
        <circle cx="10" cy="10" r="7" fill={gray} stroke="#3c4048" strokeWidth="1" />
        <path d="M10 10 L10 3 A7 7 0 0 1 16 14 Z" fill={blue} />
      </svg>
    );
  }
  if (category === 'Bar') {
    return (
      <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
        <path d="M3 3 V17 H18" stroke={axis} strokeWidth="1.2" fill="none" />
        <rect x="3" y="5" width="12" height="3" fill={blue} />
        <rect x="3" y="9" width="8" height="3" fill={gray} />
        <rect x="3" y="13" width="10" height="3" fill={blue} />
      </svg>
    );
  }
  if (category === 'Area') {
    return (
      <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
        <path d="M3 16 L3 11 L8 6 L13 10 L17 5 L17 16 Z" fill={blue} fillOpacity="0.65" stroke={blue} strokeWidth="1" />
        <path d="M3 16 L3 13 L8 10 L13 13 L17 10 L17 16 Z" fill={gray} fillOpacity="0.85" />
        <path d="M2 17 H18 M3 3 V17" stroke={axis} strokeWidth="1.2" fill="none" />
      </svg>
    );
  }
  return (
    <svg width={size} height={size} viewBox="0 0 20 20" aria-hidden="true">
      <path d="M2 17 H18 M3 3 V17" stroke={axis} strokeWidth="1.2" fill="none" />
      <circle cx="6" cy="13" r="1.6" fill={blue} />
      <circle cx="10" cy="8" r="1.6" fill={gray} />
      <circle cx="14" cy="11" r="1.6" fill={blue} />
      <circle cx="15" cy="5" r="1.6" fill={gray} />
    </svg>
  );
}

export function ChartTypeGlyph({ type, size = 38 }: { readonly type: ChartType; readonly size?: number }) {
  const s = size;
  const blue = '#5b9bd5';
  const blueDark = '#2c5d8f';
  const light = '#e4e8ef';
  const gray = '#9aa1ad';
  const grayDark = '#525864';
  const axis = '#cfd5df';

  if (type === 'clusteredColumn') {
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <path d="M4 31 H36" stroke={axis} strokeWidth="1.4" />
        <rect x="6" y="13" width="5.5" height="18" fill={blue} stroke={blueDark} strokeWidth="1" />
        <rect x="11.5" y="19" width="5.5" height="12" fill={light} stroke={grayDark} strokeWidth="1" />
        <rect x="21" y="16" width="5.5" height="15" fill={gray} stroke={grayDark} strokeWidth="1" />
        <rect x="26.5" y="7" width="5.5" height="24" fill={light} stroke={grayDark} strokeWidth="1" />
      </svg>
    );
  }
  if (type === 'stackedColumn' || type === 'stackedColumn100') {
    const full = type === 'stackedColumn100';
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <path d="M4 31 H36" stroke={axis} strokeWidth="1.4" />
        <rect x="9" y={full ? 18 : 21} width="8" height={full ? 13 : 10} fill={light} stroke={grayDark} strokeWidth="1" />
        <rect x="9" y={full ? 7 : 12} width="8" height={full ? 11 : 9} fill={blue} stroke={blueDark} strokeWidth="1" />
        <rect x="23" y={full ? 15 : 16} width="8" height={full ? 16 : 15} fill={light} stroke={grayDark} strokeWidth="1" />
        <rect x="23" y={full ? 7 : 8} width="8" height={full ? 8 : 8} fill={gray} stroke={grayDark} strokeWidth="1" />
      </svg>
    );
  }
  if (type === 'clusteredBar') {
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <path d="M6 4 V32 H36" stroke={axis} strokeWidth="1.4" fill="none" />
        <rect x="6" y="6" width="26" height="5" fill={blue} stroke={blueDark} strokeWidth="1" />
        <rect x="6" y="11" width="16" height="5" fill={light} stroke={grayDark} strokeWidth="1" />
        <rect x="6" y="20" width="13" height="5" fill={gray} stroke={grayDark} strokeWidth="1" />
        <rect x="6" y="25" width="21" height="5" fill={light} stroke={grayDark} strokeWidth="1" />
      </svg>
    );
  }
  if (type === 'stackedBar' || type === 'stackedBar100') {
    const full = type === 'stackedBar100';
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <path d="M6 4 V32 H36" stroke={axis} strokeWidth="1.4" fill="none" />
        <rect x="6" y="8" width="12" height="7" fill={light} stroke={grayDark} strokeWidth="1" />
        <rect x="18" y="8" width={full ? 15 : 10} height="7" fill={blue} stroke={blueDark} strokeWidth="1" />
        <rect x="6" y="20" width={full ? 19 : 13} height="7" fill={light} stroke={grayDark} strokeWidth="1" />
        <rect x={full ? 25 : 19} y="20" width={full ? 8 : 13} height="7" fill={gray} stroke={grayDark} strokeWidth="1" />
      </svg>
    );
  }
  if (type.startsWith('line') || type.startsWith('stackedLine')) {
    const markers = type.includes('Markers');
    const is100 = type.includes('100');
    const isStacked = type.startsWith('stackedLine');
    const topPts = is100
      ? [{ x: 7, y: 8 }, { x: 16, y: 8 }, { x: 25, y: 8 }, { x: 34, y: 8 }]
      : isStacked
        ? [{ x: 7, y: 18 }, { x: 16, y: 11 }, { x: 25, y: 15 }, { x: 34, y: 8 }]
        : [{ x: 7, y: 21 }, { x: 16, y: 10 }, { x: 25, y: 17 }, { x: 34, y: 23 }];
    const botPts = is100
      ? [{ x: 7, y: 26 }, { x: 16, y: 19 }, { x: 25, y: 24 }, { x: 34, y: 15 }]
      : isStacked
        ? [{ x: 7, y: 26 }, { x: 16, y: 20 }, { x: 25, y: 24 }, { x: 34, y: 17 }]
        : [{ x: 7, y: 26 }, { x: 16, y: 18 }, { x: 25, y: 22 }, { x: 34, y: 10 }];
    const toStr = (pts: readonly { x: number; y: number }[]) => pts.map(p => `${p.x},${p.y}`).join(' ');
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <path d="M5 31 H36" stroke={axis} strokeWidth="1.4" />
        <polyline points={toStr(botPts)} fill="none" stroke={gray} strokeWidth="1.8" />
        <polyline points={toStr(topPts)} fill="none" stroke={blue} strokeWidth="2" />
        {markers && (
          <>
            {botPts.map((p, i) => (
              <rect key={`b-${i}`} x={p.x - 1.8} y={p.y - 1.8} width="3.6" height="3.6" fill={gray} />
            ))}
            {topPts.map((p, i) => (
              <rect key={`t-${i}`} x={p.x - 1.8} y={p.y - 1.8} width="3.6" height="3.6" fill={blue} />
            ))}
          </>
        )}
      </svg>
    );
  }
  if (type === 'area' || type === 'stackedArea' || type === 'stackedArea100') {
    if (type === 'stackedArea100') {
      return (
        <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
          <rect x="6" y="7" width="28" height="23" fill={blue} fillOpacity="0.72" stroke={blueDark} strokeWidth="1.2" />
          <path d="M6 30 L6 23 L16 14 L25 22 L34 15 L34 30 Z" fill={light} fillOpacity="0.9" stroke={grayDark} strokeWidth="1.1" />
          <path d="M4 31 H36 M6 5 V31" stroke={axis} strokeWidth="1.3" fill="none" />
        </svg>
      );
    }
    if (type === 'stackedArea') {
      return (
        <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
          <path d="M6 30 L6 16 L16 8 L25 14 L34 7 L34 30 Z" fill={blue} fillOpacity="0.75" stroke={blueDark} strokeWidth="1.2" />
          <path d="M6 30 L6 23 L16 17 L25 22 L34 17 L34 30 Z" fill={light} fillOpacity="0.9" stroke={grayDark} strokeWidth="1.1" />
          <path d="M4 31 H36 M6 5 V31" stroke={axis} strokeWidth="1.3" fill="none" />
        </svg>
      );
    }
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <path d="M6 30 L6 19 L17 11 L26 24 L34 14 L34 30 Z" fill={blue} fillOpacity="0.65" stroke={blueDark} strokeWidth="1.2" />
        <path d="M6 30 L6 25 L16 18 L25 21 L34 25 L34 30 Z" fill={light} fillOpacity="0.88" stroke={grayDark} strokeWidth="1.1" />
        <path d="M4 31 H36 M6 5 V31" stroke={axis} strokeWidth="1.3" fill="none" />
      </svg>
    );
  }
  if (type === 'pie') {
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <circle cx="20" cy="18" r="12.5" fill={light} stroke={grayDark} strokeWidth="1.2" />
        <path d="M20 18 L20 5.5 A12.5 12.5 0 0 1 29.5 26 Z" fill={blue} stroke={blueDark} strokeWidth="1.2" />
      </svg>
    );
  }
  if (type === 'doughnut') {
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <circle cx="20" cy="18" r="10.5" fill="none" stroke={light} strokeWidth="6" />
        <path d="M20 7.5 A10.5 10.5 0 0 1 28 24.8" fill="none" stroke={blue} strokeWidth="6" />
        <circle cx="20" cy="18" r="13.5" fill="none" stroke={grayDark} strokeWidth="1" />
        <circle cx="20" cy="18" r="7.5" fill="none" stroke={grayDark} strokeWidth="1" />
      </svg>
    );
  }
  if (type === 'pieOfPie' || type === 'barOfPie') {
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <circle cx="13" cy="18" r="8.5" fill={blue} stroke={blueDark} strokeWidth="1.1" />
        <path d="M13 18 L20 12 L29 13 M13 18 L20 24 L29 23" stroke={axis} strokeWidth="1" fill="none" />
        {type === 'pieOfPie'
          ? <circle cx="30" cy="18" r="5" fill={light} stroke={grayDark} strokeWidth="1.1" />
          : (
            <>
              <rect x="27" y="11.5" width="6.5" height="6.5" fill={light} stroke={grayDark} strokeWidth="1" />
              <rect x="27" y="18" width="6.5" height="6.5" fill={blue} stroke={blueDark} strokeWidth="1" />
            </>
          )}
      </svg>
    );
  }
  if (type === 'bubble') {
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <path d="M4 31 H36 M6 5 V31" stroke={axis} strokeWidth="1.2" fill="none" />
        <circle cx="14" cy="21" r="5" fill={blue} fillOpacity="0.82" stroke={blueDark} strokeWidth="1.1" />
        <circle cx="26" cy="12" r="6" fill={gray} fillOpacity="0.82" stroke={grayDark} strokeWidth="1.1" />
        <circle cx="27" cy="24" r="3.8" fill={blue} fillOpacity="0.82" stroke={blueDark} strokeWidth="1.1" />
      </svg>
    );
  }
  if (type === 'scatterSmoothLines' || type === 'scatterSmoothLinesMarkers') {
    const showMarkers = type === 'scatterSmoothLinesMarkers';
    return (
      <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
        <path d="M4 31 H36 M6 5 V31" stroke={axis} strokeWidth="1.2" fill="none" />
        <path d="M9 24 C15 10, 23 10, 32 23" fill="none" stroke={blue} strokeWidth="1.8" />
        <path d="M9 14 C16 27, 25 27, 32 10" fill="none" stroke={gray} strokeWidth="1.6" />
        {showMarkers && (
          <>
            <circle cx="9" cy="24" r="2.2" fill={blue} />
            <circle cx="20" cy="13" r="2.2" fill={blue} />
            <circle cx="32" cy="23" r="2.2" fill={blue} />
            <circle cx="9" cy="14" r="2.2" fill={gray} />
            <circle cx="20" cy="24" r="2.2" fill={gray} />
            <circle cx="32" cy="10" r="2.2" fill={gray} />
          </>
        )}
      </svg>
    );
  }
  // Scatter dots or Straight lines
  const drawLines = type !== 'scatter';
  const showMarkers = type !== 'scatterStraightLines';
  return (
    <svg width={s} height={s} viewBox="0 0 40 36" aria-hidden="true">
      <path d="M4 31 H36 M6 5 V31" stroke={axis} strokeWidth="1.2" fill="none" />
      {drawLines && (
        <>
          <polyline points="9,25 17,12 24,22 33,9" fill="none" stroke={blue} strokeWidth="1.8" />
          <polyline points="9,16 17,25 24,14 33,24" fill="none" stroke={gray} strokeWidth="1.5" />
        </>
      )}
      {showMarkers && (
        <>
          <circle cx="9" cy="25" r="2.2" fill={blue} />
          <circle cx="17" cy="12" r="2.2" fill={blue} />
          <circle cx="24" cy="22" r="2.2" fill={blue} />
          <circle cx="33" cy="9" r="2.2" fill={blue} />
          <circle cx="13" cy="17" r="2.2" fill={gray} />
          <circle cx="27" cy="26" r="2.2" fill={gray} />
        </>
      )}
    </svg>
  );
}

export function RecommendedChartsGlyph({ size = 32 }: { readonly size?: number }) {
  return (
    <svg className="ribbon-glyph" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <rect x="3" y="4" width="5" height="16" fill="#a5aab4" stroke="#505662" strokeWidth="1" />
      <rect x="8" y="10" width="5" height="10" fill="#d8dce3" stroke="#505662" strokeWidth="1" />
      <text x="15" y="15" fill="#4ea1ff" fontSize="14" fontWeight="700">?</text>
    </svg>
  );
}
