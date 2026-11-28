import type { ReactNode } from 'react';

export function AddElementTriggerIcon({ size = 28 }: { readonly size?: number }) {
  return (
    <svg className="ribbon-glyph" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <rect x="3" y="10" width="5" height="9" fill="#69b3e7" stroke="#2d5f8b" strokeWidth="1" />
      <rect x="8" y="4" width="5" height="15" fill="#d4d8df" stroke="#5b6473" strokeWidth="1" />
      <rect x="13" y="7" width="5" height="12" fill="#9aa3b2" stroke="#5b6473" strokeWidth="1" />
      <path d="M18.5 11.5 V19.5 M14.5 15.5 H22.5" stroke="#5cc471" strokeWidth="2.4" strokeLinecap="round" />
    </svg>
  );
}

export function QuickLayoutTriggerIcon({ size = 28 }: { readonly size?: number }) {
  return (
    <svg className="ribbon-glyph" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <path d="M3 4 H21 M3 7 H21 M3 10 H21 M3 13 H21 M3 16 H21" stroke="#4ea1ff" strokeWidth="1.2" />
      <rect x="4" y="12" width="4" height="8" fill="#a5aab4" stroke="#454b56" strokeWidth="1" />
      <rect x="10" y="9" width="4" height="11" fill="#d4d8df" stroke="#454b56" strokeWidth="1" />
      <rect x="16" y="14" width="4" height="6" fill="#676d77" stroke="#454b56" strokeWidth="1" />
      <path d="M2 20 H22" stroke="#d8dee9" strokeWidth="1.2" />
    </svg>
  );
}

export function ChangeColorsTriggerIcon({ size = 28 }: { readonly size?: number }) {
  return (
    <svg className="ribbon-glyph" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M12 3 C6.8 3 3 6.8 3 11.8 C3 15.8 6.2 19 10.5 19 C11.9 19 12.6 18.1 12.4 16.9 C12.2 15.8 13.1 14.9 14.2 15 H16.5 C19.5 15 21.5 13 21.5 10.2 C21.5 5.8 17.4 3 12 3 Z"
        fill="#c8ced8"
        stroke="#5b6473"
        strokeWidth="1.2"
      />
      <circle cx="8" cy="9" r="1.6" fill="#4ea1ff" />
      <circle cx="12" cy="7" r="1.6" fill="#f38c4e" />
      <circle cx="16.5" cy="8.5" r="1.6" fill="#6ec071" />
      <circle cx="17" cy="12" r="1.5" fill="#e05353" />
    </svg>
  );
}

export function ElementMenuGlyph({ kind }: {
  readonly kind: 'axes' | 'axisTitles' | 'chartTitle' | 'dataLabels' | 'dataTable' | 'errorBars' | 'gridlines' | 'legend' | 'trendline';
}): ReactNode {
  switch (kind) {
    case 'axes':
      return (
        <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
          <path d="M4 3 V16 H17 M4 16 L2 16 M4 12 L2 12 M4 8 L2 8 M7 16 V18 M11 16 V18 M15 16 V18" stroke="#4ea1ff" strokeWidth="1.3" fill="none" />
          <rect x="6" y="9" width="3" height="6" fill="#a5aab4" />
          <rect x="11" y="6" width="3" height="9" fill="#d4d8df" />
        </svg>
      );
    case 'axisTitles':
      return (
        <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
          <rect x="2" y="4" width="2" height="10" fill="#4ea1ff" />
          <rect x="6" y="16" width="10" height="2" fill="#4ea1ff" />
          <rect x="7" y="8" width="3" height="6" fill="#a5aab4" />
          <rect x="12" y="5" width="3" height="9" fill="#d4d8df" />
        </svg>
      );
    case 'chartTitle':
      return (
        <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
          <rect x="5" y="2" width="10" height="2.2" fill="#4ea1ff" />
          <rect x="5" y="9" width="4" height="8" fill="#a5aab4" stroke="#5b6473" strokeWidth="0.8" />
          <rect x="11" y="6" width="4" height="11" fill="#d4d8df" stroke="#5b6473" strokeWidth="0.8" />
        </svg>
      );
    case 'dataLabels':
      return (
        <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
          <rect x="5" y="5" width="4" height="2" fill="#4ea1ff" />
          <rect x="11" y="8" width="4" height="2" fill="#4ea1ff" />
          <rect x="5" y="9" width="4" height="8" fill="#a5aab4" />
          <rect x="11" y="12" width="4" height="5" fill="#d4d8df" />
        </svg>
      );
    case 'dataTable':
      return (
        <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
          <rect x="5" y="5" width="4" height="7" fill="#a5aab4" />
          <rect x="11" y="3" width="4" height="9" fill="#d4d8df" />
          <rect x="3" y="13" width="14" height="5" fill="none" stroke="#4ea1ff" strokeWidth="1.2" />
          <path d="M9 13 V18 M3 15.5 H17" stroke="#4ea1ff" strokeWidth="1" />
        </svg>
      );
    case 'errorBars':
      return (
        <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
          <rect x="5" y="7" width="4" height="10" fill="#a5aab4" />
          <rect x="11" y="10" width="4" height="7" fill="#d4d8df" />
          <path d="M7 4 V10 M5.5 4 H8.5 M13 7 V13 M11.5 7 H14.5" stroke="#4ea1ff" strokeWidth="1.3" />
        </svg>
      );
    case 'gridlines':
      return (
        <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
          <path d="M3 5 H17 M3 10 H17 M3 15 H17 M6 3 V17 M11 3 V17 M16 3 V17" stroke="#4ea1ff" strokeWidth="1" />
          <rect x="7" y="8" width="3" height="8" fill="#a5aab4" />
          <rect x="12" y="5" width="3" height="11" fill="#d4d8df" />
        </svg>
      );
    case 'legend':
      return (
        <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
          <rect x="4" y="10" width="3" height="7" fill="#a5aab4" />
          <rect x="8.5" y="7" width="3" height="10" fill="#d4d8df" />
          <rect x="13" y="4" width="5" height="6" fill="none" stroke="#4ea1ff" strokeWidth="1.2" />
          <rect x="14.2" y="5.3" width="2.5" height="1.5" fill="#4ea1ff" />
        </svg>
      );
    case 'trendline':
      return (
        <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
          <circle cx="5" cy="14" r="1.4" fill="#d4d8df" />
          <circle cx="9" cy="11" r="1.4" fill="#d4d8df" />
          <circle cx="13" cy="9" r="1.4" fill="#d4d8df" />
          <circle cx="16" cy="5" r="1.4" fill="#d4d8df" />
          <path d="M3 16 Q11 11 17 4" fill="none" stroke="#4ea1ff" strokeWidth="1.4" />
        </svg>
      );
  }
}

export function BringForwardIcon({ size = 28 }: { readonly size?: number }) {
  return (
    <svg className="ribbon-glyph" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <rect x="10" y="9" width="10" height="10" fill="#d4d8df" stroke="#5b6473" strokeWidth="1.3" />
      <rect x="4" y="4" width="11" height="11" fill="#e6b85c" stroke="#8f6617" strokeWidth="1.3" />
    </svg>
  );
}

export function BringToFrontIcon({ size = 20 }: { readonly size?: number }) {
  return (
    <svg className="ribbon-glyph" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <rect x="3" y="4" width="8" height="8" fill="#d4d8df" stroke="#5b6473" strokeWidth="1.2" />
      <rect x="13" y="13" width="8" height="8" fill="#d4d8df" stroke="#5b6473" strokeWidth="1.2" />
      <rect x="7" y="8" width="10" height="10" fill="#e6b85c" stroke="#8f6617" strokeWidth="1.3" />
    </svg>
  );
}

export function SendBackwardIcon({ size = 28 }: { readonly size?: number }) {
  return (
    <svg className="ribbon-glyph" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <rect x="4" y="4" width="11" height="11" fill="#e6b85c" stroke="#8f6617" strokeWidth="1.3" />
      <rect x="10" y="9" width="10" height="10" fill="#d4d8df" stroke="#5b6473" strokeWidth="1.3" />
    </svg>
  );
}

export function SendToBackIcon({ size = 20 }: { readonly size?: number }) {
  return (
    <svg className="ribbon-glyph" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <rect x="7" y="8" width="10" height="10" fill="#e6b85c" stroke="#8f6617" strokeWidth="1.3" />
      <rect x="3" y="4" width="8" height="8" fill="#d4d8df" stroke="#5b6473" strokeWidth="1.2" />
      <rect x="13" y="13" width="8" height="8" fill="#d4d8df" stroke="#5b6473" strokeWidth="1.2" />
    </svg>
  );
}

export function ShapeFillIcon({ color }: { readonly color?: string | undefined }) {
  return (
    <svg className="ribbon-glyph" width={18} height={18} viewBox="0 0 20 20" aria-hidden="true">
      <path d="M9 3 L14.5 8.5 L8.5 14.5 L3 9 Z" fill="#c8ced8" stroke="#5b6473" strokeWidth="1.1" />
      <path d="M15.2 10.2 C15.2 11.7 14.2 13 15.2 13.8 C16.2 13 16.1 11.7 15.2 10.2 Z" fill="#4ea1ff" />
      <rect x="2" y="16" width="16" height="3" rx="0.5" fill={color && color !== 'none' ? color : '#f38c4e'} />
    </svg>
  );
}

export function ShapeOutlineIcon({ color }: { readonly color?: string | undefined }) {
  return (
    <svg className="ribbon-glyph" width={18} height={18} viewBox="0 0 20 20" aria-hidden="true">
      <rect x="3" y="4" width="10" height="10" fill="none" stroke="#9aa3b2" strokeWidth="1.2" />
      <path d="M8.5 12.5 L15.2 5.8 L16.8 7.4 L10.1 14.1 L7.8 14.7 Z" fill="#4ea1ff" stroke="#1f4e8c" strokeWidth="0.9" />
      <rect x="2" y="16" width="16" height="3" rx="0.5" fill={color && color !== 'none' ? color : '#4ea1ff'} />
    </svg>
  );
}

export function FormatSelectionIcon() {
  return (
    <svg className="ribbon-glyph" width={16} height={16} viewBox="0 0 18 18" aria-hidden="true">
      <rect x="2" y="3" width="14" height="12" rx="1.5" fill="none" stroke="#9aa3b2" strokeWidth="1.2" />
      <rect x="10" y="3" width="6" height="12" fill="#4ea1ff" fillOpacity="0.35" stroke="#4ea1ff" strokeWidth="1" />
      <path d="M4 7 H8 M4 10 H8" stroke="#d8dee9" strokeWidth="1.2" />
    </svg>
  );
}

export function ResetStyleIcon() {
  return (
    <svg className="ribbon-glyph" width={16} height={16} viewBox="0 0 18 18" aria-hidden="true">
      <path d="M13 5 A5 5 0 1 0 14 10" fill="none" stroke="#69b3e7" strokeWidth="1.5" />
      <path d="M13 2 V6 H9" fill="none" stroke="#69b3e7" strokeWidth="1.5" />
    </svg>
  );
}

export type SubmenuGlyphKind =
  | 'axesH'
  | 'axesV'
  | 'axisTitleH'
  | 'axisTitleV'
  | 'titleNone'
  | 'titleAbove'
  | 'titleOverlay'
  | 'labelsNone'
  | 'labelsCenter'
  | 'labelsInsideEnd'
  | 'labelsInsideBase'
  | 'labelsOutsideEnd'
  | 'labelsCallout'
  | 'tableNone'
  | 'tableLegend'
  | 'tableNoLegend'
  | 'ebNone'
  | 'ebStdErr'
  | 'ebPct'
  | 'ebStdDev'
  | 'gridMajorH'
  | 'gridMajorV'
  | 'gridMinorH'
  | 'gridMinorV'
  | 'legendNone'
  | 'legendRight'
  | 'legendTop'
  | 'legendLeft'
  | 'legendBottom'
  | 'tlNone'
  | 'tlLinear'
  | 'tlExp'
  | 'tlForecast'
  | 'tlMovingAvg';

function MiniCols({ xRed = false }: { readonly xRed?: boolean }) {
  return (
    <>
      <path d="M5 4 V19 H21" fill="none" stroke="#8c95a3" strokeWidth="1.2" />
      <rect x="7" y="11" width="3.5" height="7" fill="#a5aab4" stroke="#5b6473" strokeWidth="0.7" />
      <rect x="11" y="6" width="3.5" height="12" fill="#d4d8df" stroke="#5b6473" strokeWidth="0.7" />
      <rect x="15" y="13" width="3.5" height="5" fill="#9aa3b2" stroke="#5b6473" strokeWidth="0.7" />
      {xRed && (
        <path d="M16 3 L21 8 M21 3 L16 8" stroke="#d9534f" strokeWidth="1.7" strokeLinecap="round" />
      )}
    </>
  );
}

export function ElementSubmenuGlyph({ kind }: { readonly kind: SubmenuGlyphKind }): ReactNode {
  switch (kind) {
    case 'axesH':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <path d="M5 19 H21 M8 19 V21 M13 19 V21 M18 19 V21" stroke="#4ea1ff" strokeWidth="1.4" />
        </svg>
      );
    case 'axesV':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <path d="M5 4 V19 M3 7 H5 M3 11 H5 M3 15 H5" stroke="#4ea1ff" strokeWidth="1.4" />
        </svg>
      );
    case 'axisTitleH':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="7" y="20.5" width="11" height="2" fill="none" stroke="#4ea1ff" strokeWidth="1.2" />
        </svg>
      );
    case 'axisTitleV':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="1.5" y="6" width="2" height="11" fill="none" stroke="#4ea1ff" strokeWidth="1.2" />
        </svg>
      );
    case 'titleNone':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols xRed />
          <rect x="6" y="2" width="9" height="2" fill="none" stroke="#4ea1ff" strokeWidth="1.1" />
        </svg>
      );
    case 'titleAbove':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="7" y="2" width="11" height="2.4" fill="none" stroke="#4ea1ff" strokeWidth="1.3" />
        </svg>
      );
    case 'titleOverlay':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="7" y="5" width="11" height="2.4" fill="none" stroke="#4ea1ff" strokeWidth="1.3" />
        </svg>
      );
    case 'labelsNone':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols xRed />
        </svg>
      );
    case 'labelsCenter':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="7.5" y="13.5" width="2.5" height="2" fill="#4ea1ff" />
          <rect x="11.5" y="10.5" width="2.5" height="2" fill="#4ea1ff" />
          <rect x="15.5" y="14.5" width="2.5" height="2" fill="#4ea1ff" />
        </svg>
      );
    case 'labelsInsideEnd':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="7.5" y="11.5" width="2.5" height="2" fill="#4ea1ff" />
          <rect x="11.5" y="6.8" width="2.5" height="2" fill="#4ea1ff" />
          <rect x="15.5" y="13.5" width="2.5" height="2" fill="#4ea1ff" />
        </svg>
      );
    case 'labelsInsideBase':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="7.5" y="15.5" width="2.5" height="2" fill="#4ea1ff" />
          <rect x="11.5" y="15.5" width="2.5" height="2" fill="#4ea1ff" />
          <rect x="15.5" y="15.5" width="2.5" height="2" fill="#4ea1ff" />
        </svg>
      );
    case 'labelsOutsideEnd':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="7.5" y="8" width="2.5" height="2" fill="#4ea1ff" />
          <rect x="11.5" y="3" width="2.5" height="2" fill="#4ea1ff" />
          <rect x="15.5" y="10" width="2.5" height="2" fill="#4ea1ff" />
        </svg>
      );
    case 'labelsCallout':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <path d="M10 2 H16 V6 H13.5 L12.5 7.5 L12 6 H10 Z" fill="#4ea1ff" />
        </svg>
      );
    case 'tableNone':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols xRed />
          <rect x="5" y="18.5" width="14" height="4" fill="none" stroke="#4ea1ff" strokeWidth="1.1" />
        </svg>
      );
    case 'tableLegend':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="4" y="18.5" width="16" height="4.5" fill="#181b20" stroke="#4ea1ff" strokeWidth="1.1" />
          <path d="M10 18.5 V23 M15 18.5 V23 M4 20.8 H20" stroke="#4ea1ff" strokeWidth="1" />
          <rect x="5" y="19.5" width="1.5" height="1.5" fill="#4ea1ff" />
        </svg>
      );
    case 'tableNoLegend':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="4" y="18.5" width="16" height="4.5" fill="#181b20" stroke="#4ea1ff" strokeWidth="1.1" />
          <path d="M10 18.5 V23 M15 18.5 V23 M4 20.8 H20" stroke="#4ea1ff" strokeWidth="1" />
        </svg>
      );
    case 'ebNone':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 19 H20 M4 4 V19" stroke="#8c95a3" strokeWidth="1.2" />
          <path d="M7 5 V11 M5.5 5 H8.5 M5.5 11 H8.5 M12 9 V15 M10.5 9 H13.5 M10.5 15 H13.5" stroke="#4ea1ff" strokeWidth="1.4" />
          <path d="M15 3 L20 8 M20 3 L15 8" stroke="#d9534f" strokeWidth="1.7" strokeLinecap="round" />
        </svg>
      );
    case 'ebStdErr':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 19 H20 M4 4 V19" stroke="#8c95a3" strokeWidth="1.2" />
          <path d="M7 5 V11 M5.5 5 H8.5 M5.5 11 H8.5 M12 9 V15 M10.5 9 H13.5 M10.5 15 H13.5 M17 12 V18 M15.5 12 H18.5 M15.5 18 H18.5" stroke="#4ea1ff" strokeWidth="1.4" />
        </svg>
      );
    case 'ebPct':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 19 H20 M4 4 V19" stroke="#8c95a3" strokeWidth="1.2" />
          <path d="M7 8 V14 M5.5 8 H8.5 M5.5 14 H8.5 M12 11 V17 M10.5 11 H13.5 M10.5 17 H13.5" stroke="#4ea1ff" strokeWidth="1.4" />
          <text x="12" y="8" fill="#d8dee9" fontSize="6.5" fontWeight="700">5%</text>
        </svg>
      );
    case 'ebStdDev':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 19 H20 M4 4 V19" stroke="#8c95a3" strokeWidth="1.2" />
          <path d="M7 8 V14 M5.5 8 H8.5 M5.5 14 H8.5 M12 11 V17 M10.5 11 H13.5 M10.5 17 H13.5" stroke="#4ea1ff" strokeWidth="1.4" />
          <text x="14" y="8.5" fill="#d8dee9" fontSize="8" fontWeight="700">σ</text>
        </svg>
      );
    case 'gridMajorH':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 6 H20 M4 10 H20 M4 14 H20 M4 18 H20" stroke="#4ea1ff" strokeWidth="1" />
          <MiniCols />
        </svg>
      );
    case 'gridMajorV':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M6 4 V19 M10 4 V19 M14 4 V19 M18 4 V19" stroke="#4ea1ff" strokeWidth="1" />
          <MiniCols />
        </svg>
      );
    case 'gridMinorH':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 5 H20 M4 7 H20 M4 9 H20 M4 11 H20 M4 13 H20 M4 15 H20 M4 17 H20" stroke="#4ea1ff" strokeWidth="0.9" />
          <MiniCols />
        </svg>
      );
    case 'gridMinorV':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M6 4 V19 M8 4 V19 M10 4 V19 M12 4 V19 M14 4 V19 M16 4 V19 M18 4 V19" stroke="#4ea1ff" strokeWidth="0.9" />
          <MiniCols />
        </svg>
      );
    case 'legendNone':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols xRed />
          <rect x="14" y="2" width="6" height="5" fill="none" stroke="#4ea1ff" strokeWidth="1" />
        </svg>
      );
    case 'legendRight':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="17" y="7" width="5" height="8" fill="#181b20" stroke="#4ea1ff" strokeWidth="1.1" />
        </svg>
      );
    case 'legendTop':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="6" y="1.5" width="12" height="3.5" fill="#181b20" stroke="#4ea1ff" strokeWidth="1.1" />
        </svg>
      );
    case 'legendLeft':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="1.5" y="7" width="4.5" height="8" fill="#181b20" stroke="#4ea1ff" strokeWidth="1.1" />
        </svg>
      );
    case 'legendBottom':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <MiniCols />
          <rect x="6" y="19.5" width="12" height="3.5" fill="#181b20" stroke="#4ea1ff" strokeWidth="1.1" />
        </svg>
      );
    case 'tlNone':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 19 H20 M4 4 V19" stroke="#8c95a3" strokeWidth="1.2" />
          <circle cx="7" cy="14" r="1.2" fill="#d4d8df" /><circle cx="11" cy="11" r="1.2" fill="#d4d8df" />
          <circle cx="15" cy="9" r="1.2" fill="#d4d8df" /><circle cx="18" cy="6" r="1.2" fill="#d4d8df" />
          <path d="M5 17 L19 5" stroke="#4ea1ff" strokeWidth="1.3" />
          <path d="M15 15 L20 20 M20 15 L15 20" stroke="#d9534f" strokeWidth="1.6" />
        </svg>
      );
    case 'tlLinear':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 19 H20 M4 4 V19" stroke="#8c95a3" strokeWidth="1.2" />
          <circle cx="7" cy="14" r="1.2" fill="#d4d8df" /><circle cx="11" cy="11" r="1.2" fill="#d4d8df" />
          <circle cx="15" cy="9" r="1.2" fill="#d4d8df" /><circle cx="18" cy="6" r="1.2" fill="#d4d8df" />
          <path d="M5 17 L20 5" stroke="#4ea1ff" strokeWidth="1.5" />
        </svg>
      );
    case 'tlExp':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 19 H20 M4 4 V19" stroke="#8c95a3" strokeWidth="1.2" />
          <circle cx="7" cy="16" r="1.2" fill="#d4d8df" /><circle cx="11" cy="14" r="1.2" fill="#d4d8df" />
          <circle cx="15" cy="10" r="1.2" fill="#d4d8df" /><circle cx="18" cy="5" r="1.2" fill="#d4d8df" />
          <path d="M5 17 Q14 16 19 4" fill="none" stroke="#4ea1ff" strokeWidth="1.5" />
        </svg>
      );
    case 'tlForecast':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 19 H20 M4 4 V19" stroke="#8c95a3" strokeWidth="1.2" />
          <circle cx="7" cy="14" r="1.2" fill="#d4d8df" /><circle cx="11" cy="12" r="1.2" fill="#d4d8df" />
          <path d="M5 16 L15 9" stroke="#4ea1ff" strokeWidth="1.5" />
          <path d="M15 9 L21 5" stroke="#4ea1ff" strokeWidth="1.5" strokeDasharray="2 1.5" />
        </svg>
      );
    case 'tlMovingAvg':
      return (
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 19 H20 M4 4 V19" stroke="#8c95a3" strokeWidth="1.2" />
          <circle cx="7" cy="13" r="1.2" fill="#d4d8df" /><circle cx="11" cy="10" r="1.2" fill="#d4d8df" />
          <circle cx="15" cy="13" r="1.2" fill="#d4d8df" /><circle cx="18" cy="7" r="1.2" fill="#d4d8df" />
          <polyline points="5,15 10,9 15,12 19,6" fill="none" stroke="#4ea1ff" strokeWidth="1.5" />
        </svg>
      );
  }
}
