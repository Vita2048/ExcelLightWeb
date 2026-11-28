import { useEffect, useId, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import type { ChartDocument, ChartLineDash, ChartPlacement } from '../../core/chart';
import type { ChartReorderAction, ChartUpdatePatch } from '../../application/chart-commands';
import type { ResolvedChartData } from '../../charts/chart-data';
import { ColorPalette, ColorPickerDialog, ribbonColors, standards } from '../formatting/ColorControl';
import { positionRibbonPopup, RibbonFlyout } from '../shell/RibbonFlyout';
import {
  dashArrayFor,
  listChartElements,
  serializeSelectedElementKey,
  type ChartSelectedElement,
} from './chart-chrome';
import {
  BringForwardIcon,
  BringToFrontIcon,
  FormatSelectionIcon,
  ResetStyleIcon,
  SendBackwardIcon,
  SendToBackIcon,
  ShapeFillIcon,
  ShapeOutlineIcon,
} from './chart-ribbon-icons';

export interface ChartFormatRibbonProps {
  readonly chart: ChartDocument;
  readonly data: ResolvedChartData;
  readonly selectedElement: ChartSelectedElement;
  readonly onSelectElement: (element: ChartSelectedElement) => void;
  readonly onUpdateChart: (patch: ChartUpdatePatch) => void;
  readonly onMoveResizeChart: (placement: Partial<ChartPlacement>) => void;
  readonly onReorderChart: (action: ChartReorderAction) => void;
  readonly onOpenFormatPane: (element?: ChartSelectedElement) => void;
  readonly onResetToStyle: () => void;
}

const outlineWeights: readonly { readonly pt: number; readonly label: string }[] = [
  { pt: 0.25, label: '¼ pt' },
  { pt: 0.5, label: '½ pt' },
  { pt: 0.75, label: '¾ pt' },
  { pt: 1, label: '1 pt' },
  { pt: 1.5, label: '1½ pt' },
  { pt: 2.25, label: '2¼ pt' },
  { pt: 3, label: '3 pt' },
  { pt: 4.5, label: '4½ pt' },
  { pt: 6, label: '6 pt' },
];

const outlineDashes: readonly { readonly id: ChartLineDash; readonly label: string }[] = [
  { id: 'solid', label: 'Solid' },
  { id: 'roundDot', label: 'Round Dot' },
  { id: 'squareDot', label: 'Square Dot' },
  { id: 'dash', label: 'Dash' },
  { id: 'dashDot', label: 'Dash Dot' },
  { id: 'longDash', label: 'Long Dash' },
  { id: 'longDashDot', label: 'Long Dash Dot' },
  { id: 'longDashDotDot', label: 'Long Dash Dot Dot' },
];

export function ChartFormatRibbon({
  chart,
  data,
  selectedElement,
  onSelectElement,
  onUpdateChart,
  onMoveResizeChart,
  onReorderChart,
  onOpenFormatPane,
  onResetToStyle,
}: ChartFormatRibbonProps) {
  const choices = listChartElements(chart, data);
  const selectedKey = serializeSelectedElementKey(selectedElement);
  const [widthDraft, setWidthDraft] = useState(String(Math.round(chart.placement.width)));
  const [heightDraft, setHeightDraft] = useState(String(Math.round(chart.placement.height)));

  const [fillOpen, setFillOpen] = useState(false);
  const [fillCustomOpen, setFillCustomOpen] = useState(false);
  const [recentFillColors, setRecentFillColors] = useState<string[]>([]);
  const fillId = useId();
  const fillTrigger = useRef<HTMLButtonElement>(null);
  const fillMenu = useRef<HTMLDivElement>(null);

  const [outlineOpen, setOutlineOpen] = useState(false);
  const [outlineCustomOpen, setOutlineCustomOpen] = useState(false);
  const [outlineSub, setOutlineSub] = useState<'weight' | 'dashes' | undefined>(undefined);
  const [recentOutlineColors, setRecentOutlineColors] = useState<string[]>([]);
  const outlineId = useId();
  const outlineTrigger = useRef<HTMLButtonElement>(null);
  const outlineMenu = useRef<HTMLDivElement>(null);
  const outlineSubMenu = useRef<HTMLDivElement>(null);
  const weightTrigger = useRef<HTMLButtonElement>(null);
  const dashesTrigger = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    setWidthDraft(String(Math.round(chart.placement.width)));
    setHeightDraft(String(Math.round(chart.placement.height)));
  }, [chart.id, chart.placement.width, chart.placement.height]);

  useLayoutEffect(() => {
    if (!fillOpen) return;
    const btn = fillTrigger.current!;
    const panel = fillMenu.current!;
    const pos = () => positionRibbonPopup(btn, panel);
    const dismiss = (event: Event) => {
      const target = event.target as Node;
      if (!panel.contains(target) && !btn.contains(target)) setFillOpen(false);
    };
    const onKey = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape') {
        setFillOpen(false);
        btn.focus();
      }
    };
    pos();
    window.addEventListener('resize', pos);
    document.addEventListener('pointerdown', dismiss, true);
    document.addEventListener('keydown', onKey);
    return () => {
      window.removeEventListener('resize', pos);
      document.removeEventListener('pointerdown', dismiss, true);
      document.removeEventListener('keydown', onKey);
    };
  }, [fillOpen]);

  useLayoutEffect(() => {
    if (!outlineOpen) {
      setOutlineSub(undefined);
      return;
    }
    const btn = outlineTrigger.current!;
    const panel = outlineMenu.current!;
    const pos = () => positionRibbonPopup(btn, panel);
    const dismiss = (event: Event) => {
      const target = event.target as Node;
      if (!panel.contains(target) && !outlineSubMenu.current?.contains(target) && !btn.contains(target)) {
        setOutlineOpen(false);
      }
    };
    const onKey = (event: globalThis.KeyboardEvent) => {
      if (event.key === 'Escape') {
        setOutlineOpen(false);
        btn.focus();
      }
    };
    pos();
    window.addEventListener('resize', pos);
    document.addEventListener('pointerdown', dismiss, true);
    document.addEventListener('keydown', onKey);
    return () => {
      window.removeEventListener('resize', pos);
      document.removeEventListener('pointerdown', dismiss, true);
      document.removeEventListener('keydown', onKey);
    };
  }, [outlineOpen]);

  useLayoutEffect(() => {
    if (!outlineSub || !outlineOpen) return;
    const anchor = outlineSub === 'weight' ? weightTrigger.current : dashesTrigger.current;
    const subPanel = outlineSubMenu.current;
    if (anchor && subPanel) positionRibbonPopup(anchor, subPanel, true);
  }, [outlineSub, outlineOpen]);

  function applyShapeColor(mode: 'fill' | 'border', value: string | undefined) {
    if (selectedElement.kind === 'plotArea') {
      const current = chart.plotAreaStyle ?? {};
      if (mode === 'fill') {
        const { fillRgb: _prev, ...rest } = current;
        onUpdateChart({ plotAreaStyle: value !== undefined ? { ...rest, fillRgb: value } : rest });
      } else {
        const { borderRgb: _prev, ...rest } = current;
        onUpdateChart({ plotAreaStyle: value !== undefined ? { ...rest, borderRgb: value } : rest });
      }
      return;
    }
    if (selectedElement.kind === 'series') {
      onUpdateChart({
        series: chart.series.map(s => {
          if (s.id !== selectedElement.seriesId) return s;
          if (mode === 'fill') {
            const { fillRgb: _prev, ...rest } = s;
            return value && value !== 'none' ? { ...rest, fillRgb: value } : rest;
          }
          const { borderRgb: _prev, ...rest } = s;
          return value && value !== 'none' ? { ...rest, borderRgb: value } : rest;
        }),
      });
      return;
    }
    const current = chart.chartAreaStyle ?? {};
    if (mode === 'fill') {
      const { fillRgb: _prev, ...rest } = current;
      onUpdateChart({ chartAreaStyle: value !== undefined ? { ...rest, fillRgb: value } : rest });
    } else {
      const { borderRgb: _prev, ...rest } = current;
      onUpdateChart({ chartAreaStyle: value !== undefined ? { ...rest, borderRgb: value } : rest });
    }
  }

  function applyOutlineStroke(patch: { readonly width?: number; readonly dash?: ChartLineDash }) {
    if (selectedElement.kind === 'plotArea') {
      const current = chart.plotAreaStyle ?? {};
      onUpdateChart({
        plotAreaStyle: {
          ...current,
          ...(patch.width !== undefined && { borderWidth: patch.width }),
          ...(patch.dash !== undefined && { borderDash: patch.dash }),
        },
      });
      return;
    }
    if (selectedElement.kind === 'series') {
      onUpdateChart({
        series: chart.series.map(s => {
          if (s.id !== selectedElement.seriesId) return s;
          return {
            ...s,
            ...(patch.width !== undefined && { lineWidth: patch.width }),
            ...(patch.dash !== undefined && { lineDash: patch.dash }),
          };
        }),
      });
      return;
    }
    const current = chart.chartAreaStyle ?? {};
    onUpdateChart({
      chartAreaStyle: {
        ...current,
        ...(patch.width !== undefined && { borderWidth: patch.width }),
        ...(patch.dash !== undefined && { borderDash: patch.dash }),
      },
    });
  }

  function activeFillHex(): string | undefined {
    const raw = selectedElement.kind === 'plotArea'
      ? chart.plotAreaStyle?.fillRgb
      : selectedElement.kind === 'series'
        ? chart.series.find(s => s.id === selectedElement.seriesId)?.fillRgb
        : chart.chartAreaStyle?.fillRgb;
    return raw && raw !== 'none' ? raw.replace(/^#/, '') : undefined;
  }

  function activeOutlineInfo(): { readonly hex?: string; readonly width: number; readonly dash: ChartLineDash } {
    if (selectedElement.kind === 'plotArea') {
      const rgb = chart.plotAreaStyle?.borderRgb;
      return {
        ...(rgb && rgb !== 'none' ? { hex: rgb.replace(/^#/, '') } : {}),
        width: chart.plotAreaStyle?.borderWidth ?? 1,
        dash: chart.plotAreaStyle?.borderDash ?? 'solid',
      };
    }
    if (selectedElement.kind === 'series') {
      const ser = chart.series.find(s => s.id === selectedElement.seriesId);
      const rgb = ser?.borderRgb;
      return {
        ...(rgb && rgb !== 'none' ? { hex: rgb.replace(/^#/, '') } : {}),
        width: ser?.lineWidth ?? 2,
        dash: ser?.lineDash ?? 'solid',
      };
    }
    const rgb = chart.chartAreaStyle?.borderRgb;
    return {
      ...(rgb && rgb !== 'none' ? { hex: rgb.replace(/^#/, '') } : {}),
      width: chart.chartAreaStyle?.borderWidth ?? 1,
      dash: chart.chartAreaStyle?.borderDash ?? 'solid',
    };
  }

  function commitDimension(which: 'width' | 'height') {
    const raw = Number(which === 'width' ? widthDraft : heightDraft);
    if (!Number.isFinite(raw)) {
      setWidthDraft(String(Math.round(chart.placement.width)));
      setHeightDraft(String(Math.round(chart.placement.height)));
      return;
    }
    const clamped = Math.round(Math.max(which === 'width' ? 120 : 90, Math.min(4096, raw)));
    if (which === 'width') {
      setWidthDraft(String(clamped));
      if (clamped !== Math.round(chart.placement.width)) onMoveResizeChart({ width: clamped });
    } else {
      setHeightDraft(String(clamped));
      if (clamped !== Math.round(chart.placement.height)) onMoveResizeChart({ height: clamped });
    }
  }

  const fillHex = activeFillHex();
  const outlineInfo = activeOutlineInfo();

  return (
    <>
      <fieldset>
        <legend>Current Selection</legend>
        <div className="ribbon-stack chart-format-selection">
          <select
            className="chart-format-select"
            aria-label="Chart Elements"
            value={choices.some(c => c.key === selectedKey) ? selectedKey : 'chartArea'}
            onChange={event => {
              const found = choices.find(c => c.key === event.currentTarget.value);
              if (found) onSelectElement(found.element);
            }}
          >
            {choices.map(choice => (
              <option key={choice.key} value={choice.key}>{choice.label}</option>
            ))}
          </select>
          <button type="button" className="ribbon-command" onClick={() => onOpenFormatPane(selectedElement)}>
            <FormatSelectionIcon />
            Format Selection
          </button>
          <button type="button" className="ribbon-command" onClick={onResetToStyle}>
            <ResetStyleIcon />
            Reset to Match Style
          </button>
        </div>
      </fieldset>

      <fieldset>
        <legend>Shape Styles</legend>
        <div className="ribbon-stack">
          <button
            ref={fillTrigger}
            type="button"
            className="ribbon-command"
            aria-label="Shape Fill"
            aria-haspopup="menu"
            aria-expanded={fillOpen}
            aria-controls={fillOpen ? fillId : undefined}
            onClick={() => {
              setOutlineOpen(false);
              setFillOpen(o => !o);
            }}
          >
            <ShapeFillIcon color={fillHex ? '#' + fillHex : undefined} />
            Shape Fill
            <span className="ribbon-chevron" aria-hidden="true" />
          </button>

          {fillOpen && createPortal(
            <div
              ref={fillMenu}
              id={fillId}
              role="group"
              aria-label="Shape Fill"
              className="ribbon-palette chart-shape-palette"
            >
              <button
                type="button"
                className="ribbon-palette-action"
                onClick={() => {
                  applyShapeColor('fill', undefined);
                  setFillOpen(false);
                }}
              >
                Automatic
              </button>
              <fieldset className="ribbon-theme-colors">
                <legend>Theme Colors</legend>
                <ColorPalette
                  colors={ribbonColors}
                  value={fillHex}
                  onChange={hex => {
                    applyShapeColor('fill', '#' + hex);
                    setFillOpen(false);
                  }}
                />
              </fieldset>
              <fieldset>
                <legend>Standard Colors</legend>
                <ColorPalette
                  colors={standards}
                  value={fillHex}
                  onChange={hex => {
                    applyShapeColor('fill', '#' + hex);
                    setFillOpen(false);
                  }}
                />
              </fieldset>
              <button
                type="button"
                className="ribbon-palette-action"
                onClick={() => {
                  applyShapeColor('fill', 'none');
                  setFillOpen(false);
                }}
              >
                No Fill
              </button>
              <button
                type="button"
                className="ribbon-palette-action"
                onClick={() => {
                  setFillOpen(false);
                  setFillCustomOpen(true);
                }}
              >
                More Fill Colors…
              </button>
            </div>,
            document.body,
          )}

          {fillCustomOpen && (
            <ColorPickerDialog
              label="More Fill Colors"
              value={fillHex}
              automatic={false}
              ribbon
              customOnly
              recentColors={recentFillColors}
              onClose={() => setFillCustomOpen(false)}
              onChange={(hex, custom) => {
                if (hex && custom) setRecentFillColors(c => [hex, ...c.filter(x => x !== hex)].slice(0, 8));
                applyShapeColor('fill', hex ? '#' + hex : undefined);
                setFillCustomOpen(false);
              }}
            />
          )}

          <button
            ref={outlineTrigger}
            type="button"
            className="ribbon-command"
            aria-label="Shape Outline"
            aria-haspopup="menu"
            aria-expanded={outlineOpen}
            aria-controls={outlineOpen ? outlineId : undefined}
            onClick={() => {
              setFillOpen(false);
              setOutlineOpen(o => !o);
            }}
          >
            <ShapeOutlineIcon color={outlineInfo.hex ? '#' + outlineInfo.hex : undefined} />
            Shape Outline
            <span className="ribbon-chevron" aria-hidden="true" />
          </button>

          {outlineOpen && createPortal(
            <>
              <div
                ref={outlineMenu}
                id={outlineId}
                role="group"
                aria-label="Shape Outline"
                className="ribbon-palette chart-shape-palette"
              >
                <button
                  type="button"
                  className="ribbon-palette-action"
                  onClick={() => {
                    applyShapeColor('border', undefined);
                    setOutlineOpen(false);
                  }}
                >
                  Automatic
                </button>
                <fieldset className="ribbon-theme-colors">
                  <legend>Theme Colors</legend>
                  <ColorPalette
                    colors={ribbonColors}
                    value={outlineInfo.hex}
                    onChange={hex => {
                      applyShapeColor('border', '#' + hex);
                      setOutlineOpen(false);
                    }}
                  />
                </fieldset>
                <fieldset>
                  <legend>Standard Colors</legend>
                  <ColorPalette
                    colors={standards}
                    value={outlineInfo.hex}
                    onChange={hex => {
                      applyShapeColor('border', '#' + hex);
                      setOutlineOpen(false);
                    }}
                  />
                </fieldset>
                <button
                  type="button"
                  className="ribbon-palette-action"
                  onClick={() => {
                    applyShapeColor('border', 'none');
                    setOutlineOpen(false);
                  }}
                >
                  No Outline
                </button>
                <button
                  type="button"
                  className="ribbon-palette-action"
                  onClick={() => {
                    setOutlineOpen(false);
                    setOutlineCustomOpen(true);
                  }}
                >
                  More Outline Colors…
                </button>
                <div className="ribbon-flyout-divider">
                  <button
                    ref={weightTrigger}
                    type="button"
                    className="ribbon-palette-action chart-outline-subtrigger"
                    aria-haspopup="menu"
                    aria-expanded={outlineSub === 'weight'}
                    onMouseEnter={() => setOutlineSub('weight')}
                    onClick={() => setOutlineSub('weight')}
                  >
                    <span>Weight</span>
                    <span className="ribbon-submenu-chevron" aria-hidden="true" />
                  </button>
                  <button
                    ref={dashesTrigger}
                    type="button"
                    className="ribbon-palette-action chart-outline-subtrigger"
                    aria-haspopup="menu"
                    aria-expanded={outlineSub === 'dashes'}
                    onMouseEnter={() => setOutlineSub('dashes')}
                    onClick={() => setOutlineSub('dashes')}
                  >
                    <span>Dashes</span>
                    <span className="ribbon-submenu-chevron" aria-hidden="true" />
                  </button>
                </div>
              </div>

              {outlineSub && (
                <div
                  ref={outlineSubMenu}
                  role="menu"
                  aria-label={outlineSub === 'weight' ? 'Outline Weight' : 'Outline Dashes'}
                  className="ribbon-flyout chart-outline-submenu"
                >
                  {outlineSub === 'weight'
                    ? outlineWeights.map(w => (
                      <button
                        key={w.pt}
                        type="button"
                        role="menuitemradio"
                        aria-checked={Math.abs(outlineInfo.width - w.pt) < 0.05}
                        onClick={() => {
                          applyOutlineStroke({ width: w.pt });
                          setOutlineOpen(false);
                        }}
                      >
                        <span className="chart-weight-label">{w.label}</span>
                        <svg width="90" height="12" viewBox="0 0 90 12" aria-hidden="true">
                          <line x1="2" y1="6" x2="88" y2="6" stroke="currentColor" strokeWidth={Math.min(6, w.pt)} />
                        </svg>
                      </button>
                    ))
                    : outlineDashes.map(d => (
                      <button
                        key={d.id}
                        type="button"
                        role="menuitemradio"
                        aria-checked={outlineInfo.dash === d.id}
                        aria-label={d.label}
                        onClick={() => {
                          applyOutlineStroke({ dash: d.id });
                          setOutlineOpen(false);
                        }}
                      >
                        <svg width="110" height="12" viewBox="0 0 110 12" aria-hidden="true">
                          <line
                            x1="2"
                            y1="6"
                            x2="108"
                            y2="6"
                            stroke="currentColor"
                            strokeWidth="2"
                            strokeDasharray={dashArrayFor(d.id, 2)}
                          />
                        </svg>
                      </button>
                    ))}
                  <div className="ribbon-flyout-divider" role="none">
                    <button
                      type="button"
                      role="menuitem"
                      onClick={() => {
                        setOutlineOpen(false);
                        onOpenFormatPane(selectedElement);
                      }}
                    >
                      More Lines…
                    </button>
                  </div>
                </div>
              )}
            </>,
            document.body,
          )}

          {outlineCustomOpen && (
            <ColorPickerDialog
              label="More Outline Colors"
              value={outlineInfo.hex}
              automatic
              ribbon
              customOnly
              recentColors={recentOutlineColors}
              onClose={() => setOutlineCustomOpen(false)}
              onChange={(hex, custom) => {
                if (hex && custom) setRecentOutlineColors(c => [hex, ...c.filter(x => x !== hex)].slice(0, 8));
                applyShapeColor('border', hex ? '#' + hex : undefined);
                setOutlineCustomOpen(false);
              }}
            />
          )}
        </div>
      </fieldset>

      <fieldset>
        <legend>Arrange</legend>
        <RibbonFlyout
          label="Bring Forward"
          caption={<>Bring<br />Forward</>}
          iconNode={<BringForwardIcon />}
          large
          items={[
            { label: 'Bring Forward', iconNode: <BringForwardIcon size={20} />, select: () => onReorderChart('bringForward') },
            { label: 'Bring to Front', iconNode: <BringToFrontIcon size={20} />, select: () => onReorderChart('bringToFront') },
          ]}
        />
        <RibbonFlyout
          label="Send Backward"
          caption={<>Send<br />Backward</>}
          iconNode={<SendBackwardIcon />}
          large
          items={[
            { label: 'Send Backward', iconNode: <SendBackwardIcon size={20} />, select: () => onReorderChart('sendBackward') },
            { label: 'Send to Back', iconNode: <SendToBackIcon size={20} />, select: () => onReorderChart('sendToBack') },
          ]}
        />
      </fieldset>

      <fieldset>
        <legend>Size</legend>
        <div className="ribbon-stack chart-format-size">
          <label>
            <span>Height:</span>
            <input
              type="number"
              min={90}
              max={4096}
              aria-label="Height"
              value={heightDraft}
              onChange={event => setHeightDraft(event.currentTarget.value)}
              onBlur={() => commitDimension('height')}
              onKeyDown={event => { if (event.key === 'Enter') commitDimension('height'); }}
            />
            <span className="chart-size-unit">px</span>
          </label>
          <label>
            <span>Width:</span>
            <input
              type="number"
              min={120}
              max={4096}
              aria-label="Width"
              value={widthDraft}
              onChange={event => setWidthDraft(event.currentTarget.value)}
              onBlur={() => commitDimension('width')}
              onKeyDown={event => { if (event.key === 'Enter') commitDimension('width'); }}
            />
            <span className="chart-size-unit">px</span>
          </label>
        </div>
      </fieldset>
    </>
  );
}
