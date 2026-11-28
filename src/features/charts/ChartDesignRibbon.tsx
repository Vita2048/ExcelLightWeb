import { useId, useLayoutEffect, useRef, useState, type KeyboardEvent } from 'react';
import { createPortal } from 'react-dom';
import type { ChartDocument } from '../../core/chart';
import type { ChartUpdatePatch } from '../../application/chart-commands';
import type { ResolvedChartData } from '../../charts/chart-data';
import { chartPalettes, getSeriesColor } from '../../charts/palettes';
import { positionRibbonPopup, RibbonFlyout, type RibbonItem } from '../shell/RibbonFlyout';
import {
  applyChartStylePatch,
  applyQuickLayout,
  chartStylePresets,
  patchChartElements,
  quickLayoutPresets,
  setChartDataLabels,
  setChartLegend,
  supportedElementsForType,
  toggleChartAxis,
  toggleChartAxisTitle,
  toggleChartGridline,
  togglePrimaryErrorBars,
  togglePrimaryTrendline,
  type ChartSelectedElement,
} from './chart-chrome';
import {
  AddElementTriggerIcon,
  ChangeColorsTriggerIcon,
  ElementMenuGlyph,
  ElementSubmenuGlyph,
  QuickLayoutTriggerIcon,
} from './chart-ribbon-icons';

export interface ChartDesignRibbonProps {
  readonly chart: ChartDocument;
  readonly data: ResolvedChartData;
  readonly onUpdateChart: (patch: ChartUpdatePatch) => void;
  readonly onSwitchRowColumn: () => void;
  readonly onOpenSelectData: () => void;
  readonly onOpenChangeChartType: () => void;
  readonly onOpenMoveChart: () => void;
  readonly onOpenFormatPane: (element?: ChartSelectedElement) => void;
}

function ChangeColorsFlyout({
  chart,
  onUpdateChart,
}: {
  readonly chart: ChartDocument;
  readonly onUpdateChart: (patch: ChartUpdatePatch) => void;
}) {
  const id = useId();
  const [open, setOpen] = useState(false);
  const trigger = useRef<HTMLButtonElement>(null);
  const panel = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    if (!open) return;
    const button = trigger.current!;
    const menu = panel.current!;
    const position = () => positionRibbonPopup(button, menu);
    function dismiss(event: PointerEvent | FocusEvent) {
      const target = event.target as Node;
      if (!menu.contains(target) && !button.contains(target)) setOpen(false);
    }
    function onScroll(event: Event) {
      if (event.target instanceof Node && menu.contains(event.target)) return;
      position();
    }
    position();
    menu.querySelector('button')?.focus({ preventScroll: true });
    window.addEventListener('resize', position);
    document.addEventListener('scroll', onScroll, true);
    document.addEventListener('pointerdown', dismiss);
    document.addEventListener('focusin', dismiss);
    return () => {
      window.removeEventListener('resize', position);
      document.removeEventListener('scroll', onScroll, true);
      document.removeEventListener('pointerdown', dismiss);
      document.removeEventListener('focusin', dismiss);
    };
  }, [open]);

  function close() {
    setOpen(false);
    trigger.current?.focus();
  }

  function applyPalette(paletteId: string) {
    onUpdateChart({ paletteId });
    close();
  }

  function onMenuKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (event.key === 'Escape') {
      event.preventDefault();
      close();
      return;
    }
    const buttons = [...event.currentTarget.querySelectorAll('button')];
    const current = buttons.indexOf(document.activeElement as HTMLButtonElement);
    const next = event.key === 'ArrowDown' ? (current + 1) % buttons.length
      : event.key === 'ArrowUp' ? (current + buttons.length - 1) % buttons.length
      : event.key === 'Home' ? 0
      : event.key === 'End' ? buttons.length - 1
      : undefined;
    if (next !== undefined) {
      event.preventDefault();
      buttons[next]?.focus();
    }
    if (event.key === 'Tab') close();
  }

  const renderGroup = (prefix: 'colorful' | 'monochrome') =>
    chartPalettes.filter(p => p.id.startsWith(prefix)).map(pal => (
      <button
        key={pal.id}
        type="button"
        role="menuitemradio"
        aria-checked={chart.paletteId === pal.id}
        aria-label={pal.name}
        title={pal.name}
        className={'chart-palette-strip-btn' + (chart.paletteId === pal.id ? ' selected' : '')}
        onClick={() => applyPalette(pal.id)}
      >
        <span className="chart-palette-swatches" aria-hidden="true">
          {pal.colors.slice(0, 6).map(hex => (
            <span key={hex} className="chart-palette-swatch" style={{ background: hex }} />
          ))}
        </span>
      </button>
    ));

  return (
    <>
      <button
        ref={trigger}
        type="button"
        className="ribbon-command ribbon-large"
        aria-label="Change Colors"
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={open ? id : undefined}
        onClick={() => { if (open) close(); else setOpen(true); }}
        onKeyDown={event => {
          if (event.key === 'ArrowDown') { event.preventDefault(); setOpen(true); }
          if (event.key === 'Escape' && open) { event.preventDefault(); close(); }
        }}
      >
        <ChangeColorsTriggerIcon />
        Change<br />Colors
        <span className="ribbon-chevron" aria-hidden="true" />
      </button>
      {open && createPortal(
        <div
          ref={panel}
          id={id}
          role="menu"
          aria-label="Change Colors"
          className="ribbon-flyout chart-colors-gallery"
          onKeyDown={onMenuKeyDown}
        >
          <div className="ribbon-flyout-heading" aria-hidden="true">Colorful</div>
          {renderGroup('colorful')}
          <div className="ribbon-flyout-heading ribbon-flyout-divider" aria-hidden="true">Monochromatic</div>
          {renderGroup('monochrome')}
        </div>,
        document.body,
      )}
    </>
  );
}

export function ChartDesignRibbon({
  chart,
  onUpdateChart,
  onSwitchRowColumn,
  onOpenSelectData,
  onOpenChangeChartType,
  onOpenMoveChart,
  onOpenFormatPane,
}: ChartDesignRibbonProps) {
  const supported = supportedElementsForType(chart.chartType);

  const addElementItems: RibbonItem[] = [];
  if (supported.axes) {
    addElementItems.push({
      label: 'Axes',
      iconNode: <ElementMenuGlyph kind="axes" />,
      children: [
        {
          label: 'Primary Horizontal',
          iconNode: <ElementSubmenuGlyph kind="axesH" />,
          checked: chart.elements.axes?.category?.show !== false,
          select: () => onUpdateChart(toggleChartAxis(chart, 'category')),
        },
        {
          label: 'Primary Vertical',
          iconNode: <ElementSubmenuGlyph kind="axesV" />,
          checked: chart.elements.axes?.value?.show !== false,
          select: () => onUpdateChart(toggleChartAxis(chart, 'value')),
        },
        {
          label: 'More Axis Options...',
          divider: true,
          select: () => onOpenFormatPane({ kind: 'axis', axis: 'value' }),
        },
      ],
    });
  }
  if (supported.axisTitles) {
    addElementItems.push({
      label: 'Axis Titles',
      iconNode: <ElementMenuGlyph kind="axisTitles" />,
      children: [
        {
          label: 'Primary Horizontal',
          iconNode: <ElementSubmenuGlyph kind="axisTitleH" />,
          checked: !!chart.elements.axes?.category?.title?.show,
          select: () => onUpdateChart(toggleChartAxisTitle(chart, 'category')),
        },
        {
          label: 'Primary Vertical',
          iconNode: <ElementSubmenuGlyph kind="axisTitleV" />,
          checked: !!chart.elements.axes?.value?.title?.show,
          select: () => onUpdateChart(toggleChartAxisTitle(chart, 'value')),
        },
        {
          label: 'More Axis Title Options...',
          divider: true,
          select: () => onOpenFormatPane({ kind: 'axis', axis: 'value' }),
        },
      ],
    });
  }
  if (supported.chartTitle) {
    addElementItems.push({
      label: 'Chart Title',
      iconNode: <ElementMenuGlyph kind="chartTitle" />,
      children: [
        {
          label: 'None',
          iconNode: <ElementSubmenuGlyph kind="titleNone" />,
          checked: chart.elements.title?.show === false,
          select: () => onUpdateChart(patchChartElements(chart, { title: { ...chart.elements.title, show: false } })),
        },
        {
          label: 'Above Chart',
          iconNode: <ElementSubmenuGlyph kind="titleAbove" />,
          checked: chart.elements.title?.show !== false && !chart.elements.title?.overlay,
          select: () => onUpdateChart(patchChartElements(chart, {
            title: { ...chart.elements.title, show: true, overlay: false, text: chart.elements.title?.text || 'Chart Title' },
          })),
        },
        {
          label: 'Centered Overlay',
          iconNode: <ElementSubmenuGlyph kind="titleOverlay" />,
          checked: chart.elements.title?.show !== false && !!chart.elements.title?.overlay,
          select: () => onUpdateChart(patchChartElements(chart, {
            title: { ...chart.elements.title, show: true, overlay: true, text: chart.elements.title?.text || 'Chart Title' },
          })),
        },
        {
          label: 'More Title Options...',
          divider: true,
          select: () => onOpenFormatPane({ kind: 'title' }),
        },
      ],
    });
  }
  if (supported.dataLabels) {
    addElementItems.push({
      label: 'Data Labels',
      iconNode: <ElementMenuGlyph kind="dataLabels" />,
      children: [
        {
          label: 'None',
          iconNode: <ElementSubmenuGlyph kind="labelsNone" />,
          checked: !chart.elements.dataLabels?.show,
          select: () => onUpdateChart(setChartDataLabels(chart, false)),
        },
        {
          label: 'Center',
          iconNode: <ElementSubmenuGlyph kind="labelsCenter" />,
          checked: !!chart.elements.dataLabels?.show && chart.elements.dataLabels?.position === 'center',
          select: () => onUpdateChart(setChartDataLabels(chart, true, 'center')),
        },
        {
          label: 'Inside End',
          iconNode: <ElementSubmenuGlyph kind="labelsInsideEnd" />,
          checked: !!chart.elements.dataLabels?.show && chart.elements.dataLabels?.position === 'insideEnd',
          select: () => onUpdateChart(setChartDataLabels(chart, true, 'insideEnd')),
        },
        {
          label: 'Inside Base',
          iconNode: <ElementSubmenuGlyph kind="labelsInsideBase" />,
          checked: !!chart.elements.dataLabels?.show && chart.elements.dataLabels?.position === 'insideBase',
          select: () => onUpdateChart(setChartDataLabels(chart, true, 'insideBase')),
        },
        {
          label: 'Outside End',
          iconNode: <ElementSubmenuGlyph kind="labelsOutsideEnd" />,
          checked: !!chart.elements.dataLabels?.show && chart.elements.dataLabels?.position === 'outsideEnd',
          select: () => onUpdateChart(setChartDataLabels(chart, true, 'outsideEnd')),
        },
        {
          label: 'Data Callout',
          iconNode: <ElementSubmenuGlyph kind="labelsCallout" />,
          checked: !!chart.elements.dataLabels?.show && chart.elements.dataLabels?.position === 'callout',
          select: () => onUpdateChart(setChartDataLabels(chart, true, 'callout')),
        },
        {
          label: 'More Data Label Options...',
          divider: true,
          select: () => onOpenFormatPane({ kind: 'dataLabels' }),
        },
      ],
    });
  }
  if (supported.dataTable) {
    addElementItems.push({
      label: 'Data Table',
      iconNode: <ElementMenuGlyph kind="dataTable" />,
      children: [
        {
          label: 'None',
          iconNode: <ElementSubmenuGlyph kind="tableNone" />,
          checked: !chart.elements.dataTable?.show,
          select: () => onUpdateChart(patchChartElements(chart, { dataTable: { show: false } })),
        },
        {
          label: 'With Legend Keys',
          iconNode: <ElementSubmenuGlyph kind="tableLegend" />,
          checked: !!chart.elements.dataTable?.show && chart.elements.dataTable?.showLegendKeys !== false,
          select: () => onUpdateChart(patchChartElements(chart, { dataTable: { show: true, showLegendKeys: true } })),
        },
        {
          label: 'No Legend Keys',
          iconNode: <ElementSubmenuGlyph kind="tableNoLegend" />,
          checked: !!chart.elements.dataTable?.show && chart.elements.dataTable?.showLegendKeys === false,
          select: () => onUpdateChart(patchChartElements(chart, { dataTable: { show: true, showLegendKeys: false } })),
        },
        {
          label: 'More Data Table Options...',
          divider: true,
          select: () => onOpenFormatPane({ kind: 'chartArea' }),
        },
      ],
    });
  }
  if (supported.errorBars) {
    addElementItems.push({
      label: 'Error Bars',
      iconNode: <ElementMenuGlyph kind="errorBars" />,
      children: [
        {
          label: 'None',
          iconNode: <ElementSubmenuGlyph kind="ebNone" />,
          checked: !chart.series.some(s => s.errorBars !== undefined),
          select: () => onUpdateChart(togglePrimaryErrorBars(chart, undefined)),
        },
        {
          label: 'Standard Error',
          iconNode: <ElementSubmenuGlyph kind="ebStdErr" />,
          select: () => onUpdateChart(togglePrimaryErrorBars(chart, 'stdErr')),
        },
        {
          label: 'Percentage',
          iconNode: <ElementSubmenuGlyph kind="ebPct" />,
          select: () => onUpdateChart(togglePrimaryErrorBars(chart, 'percentage')),
        },
        {
          label: 'Standard Deviation',
          iconNode: <ElementSubmenuGlyph kind="ebStdDev" />,
          select: () => onUpdateChart(togglePrimaryErrorBars(chart, 'stdDev')),
        },
      ],
    });
  }
  if (supported.gridlines) {
    addElementItems.push({
      label: 'Gridlines',
      iconNode: <ElementMenuGlyph kind="gridlines" />,
      children: [
        {
          label: 'Primary Major Horizontal',
          iconNode: <ElementSubmenuGlyph kind="gridMajorH" />,
          checked: !!chart.elements.gridlines?.primaryMajorHorizontal,
          select: () => onUpdateChart(toggleChartGridline(chart, 'primaryMajorHorizontal')),
        },
        {
          label: 'Primary Major Vertical',
          iconNode: <ElementSubmenuGlyph kind="gridMajorV" />,
          checked: !!chart.elements.gridlines?.primaryMajorVertical,
          select: () => onUpdateChart(toggleChartGridline(chart, 'primaryMajorVertical')),
        },
        {
          label: 'Primary Minor Horizontal',
          iconNode: <ElementSubmenuGlyph kind="gridMinorH" />,
          checked: !!chart.elements.gridlines?.primaryMinorHorizontal,
          select: () => onUpdateChart(toggleChartGridline(chart, 'primaryMinorHorizontal')),
        },
        {
          label: 'Primary Minor Vertical',
          iconNode: <ElementSubmenuGlyph kind="gridMinorV" />,
          checked: !!chart.elements.gridlines?.primaryMinorVertical,
          select: () => onUpdateChart(toggleChartGridline(chart, 'primaryMinorVertical')),
        },
        {
          label: 'More Gridline Options...',
          divider: true,
          select: () => onOpenFormatPane({ kind: 'axis', axis: 'value' }),
        },
      ],
    });
  }
  if (supported.legend) {
    addElementItems.push({
      label: 'Legend',
      iconNode: <ElementMenuGlyph kind="legend" />,
      children: [
        {
          label: 'None',
          iconNode: <ElementSubmenuGlyph kind="legendNone" />,
          checked: chart.elements.legend?.show === false,
          select: () => onUpdateChart(setChartLegend(chart, false)),
        },
        {
          label: 'Right',
          iconNode: <ElementSubmenuGlyph kind="legendRight" />,
          checked: chart.elements.legend?.show !== false && chart.elements.legend?.position === 'right',
          select: () => onUpdateChart(setChartLegend(chart, true, 'right')),
        },
        {
          label: 'Top',
          iconNode: <ElementSubmenuGlyph kind="legendTop" />,
          checked: chart.elements.legend?.show !== false && chart.elements.legend?.position === 'top',
          select: () => onUpdateChart(setChartLegend(chart, true, 'top')),
        },
        {
          label: 'Left',
          iconNode: <ElementSubmenuGlyph kind="legendLeft" />,
          checked: chart.elements.legend?.show !== false && chart.elements.legend?.position === 'left',
          select: () => onUpdateChart(setChartLegend(chart, true, 'left')),
        },
        {
          label: 'Bottom',
          iconNode: <ElementSubmenuGlyph kind="legendBottom" />,
          checked: chart.elements.legend?.show !== false && chart.elements.legend?.position === 'bottom',
          select: () => onUpdateChart(setChartLegend(chart, true, 'bottom')),
        },
        {
          label: 'More Legend Options...',
          divider: true,
          select: () => onOpenFormatPane({ kind: 'legend' }),
        },
      ],
    });
  }
  if (supported.trendline) {
    addElementItems.push({
      label: 'Trendline',
      iconNode: <ElementMenuGlyph kind="trendline" />,
      children: [
        {
          label: 'None',
          iconNode: <ElementSubmenuGlyph kind="tlNone" />,
          checked: !chart.series.some(s => (s.trendlines?.length ?? 0) > 0),
          select: () => onUpdateChart(togglePrimaryTrendline(chart, undefined)),
        },
        {
          label: 'Linear',
          iconNode: <ElementSubmenuGlyph kind="tlLinear" />,
          select: () => onUpdateChart(togglePrimaryTrendline(chart, 'linear')),
        },
        {
          label: 'Exponential',
          iconNode: <ElementSubmenuGlyph kind="tlExp" />,
          select: () => onUpdateChart(togglePrimaryTrendline(chart, 'exponential')),
        },
        {
          label: 'Linear Forecast',
          iconNode: <ElementSubmenuGlyph kind="tlForecast" />,
          select: () => onUpdateChart(togglePrimaryTrendline(chart, 'linear', undefined, 2)),
        },
        {
          label: 'Two Period Moving Average',
          iconNode: <ElementSubmenuGlyph kind="tlMovingAvg" />,
          select: () => onUpdateChart(togglePrimaryTrendline(chart, 'movingAverage')),
        },
        {
          label: 'More Trendline Options...',
          divider: true,
          select: () => {
            const owner = chart.series.find(s => (s.trendlines?.length ?? 0) > 0);
            const tl = owner?.trendlines?.[0];
            if (owner && tl) onOpenFormatPane({ kind: 'trendline', seriesId: owner.id, trendlineId: tl.id });
          },
        },
      ],
    });
  }

  return (
    <>
      <fieldset>
        <legend>Chart Layouts</legend>
        <RibbonFlyout
          label="Add Chart Element"
          caption={<>Add Chart<br />Element</>}
          iconNode={<AddElementTriggerIcon />}
          large
          items={addElementItems}
        />
        <RibbonFlyout
          label="Quick Layout"
          caption={<>Quick<br />Layout</>}
          iconNode={<QuickLayoutTriggerIcon />}
          large
          items={quickLayoutPresets.map(q => ({
            label: q.label,
            iconNode: <QuickLayoutTriggerIcon size={22} />,
            detail: q.description,
            select: () => onUpdateChart(applyQuickLayout(chart, q.id)),
          }))}
        />
      </fieldset>

      <fieldset>
        <legend>Chart Styles</legend>
        <ChangeColorsFlyout
          chart={chart}
          onUpdateChart={onUpdateChart}
        />
        <div className="chart-design-styles-strip" role="listbox" aria-label="Chart Styles">
          {chartStylePresets.map(preset => (
            <button
              key={preset.id}
              type="button"
              role="option"
              aria-selected={chart.styleId === preset.id}
              aria-label={preset.name}
              title={preset.name}
              className={'chart-design-style-thumb' + (chart.styleId === preset.id ? ' selected' : '')}
              onClick={() => onUpdateChart(applyChartStylePatch(preset.id))}
            >
              <svg width="46" height="34" viewBox="0 0 46 34" aria-hidden="true">
                <rect
                  x="1"
                  y="1"
                  width="44"
                  height="32"
                  rx="2"
                  fill={preset.chartAreaBg}
                  stroke={preset.chartAreaBorder === 'none' ? '#2c323c' : preset.chartAreaBorder}
                />
                <rect
                  x="5"
                  y="5"
                  width="36"
                  height="22"
                  fill={preset.plotAreaBg === 'transparent' ? 'none' : preset.plotAreaBg}
                  stroke={preset.plotAreaBorder ?? 'none'}
                />
                {preset.showGridlines && (
                  <line x1="6" y1="14" x2="40" y2="14" stroke="#2c3340" strokeWidth="1" />
                )}
                <rect x="8" y="15" width="6" height="11" fill={getSeriesColor(chart.paletteId, 0)} />
                <rect x="16" y="10" width="6" height="16" fill={getSeriesColor(chart.paletteId, 1)} />
                <rect x="24" y="13" width="6" height="13" fill={getSeriesColor(chart.paletteId, 2)} />
                <rect x="32" y="8" width="6" height="18" fill={getSeriesColor(chart.paletteId, 0)} />
              </svg>
            </button>
          ))}
        </div>
      </fieldset>

      <fieldset>
        <legend>Data</legend>
        <button
          type="button"
          className="ribbon-command ribbon-large"
          aria-label="Switch Row/Column"
          disabled={chart.source.kind !== 'range'}
          title={chart.source.kind !== 'range' ? 'Switch Row/Column is available for range-based charts' : 'Switch Row/Column'}
          onClick={onSwitchRowColumn}
        >
          <svg className="ribbon-glyph" width={28} height={28} viewBox="0 0 24 24" aria-hidden="true">
            <rect x="4" y="5" width="7" height="12" fill="#4ea1ff" stroke="#1f4e8c" />
            <rect x="13" y="9" width="7" height="8" fill="#a5aab4" stroke="#454b56" />
            <path d="M5 20 H19 M16 18 L19 20 L16 22 M8 18 L5 20 L8 22" fill="none" stroke="#d8dee9" strokeWidth="1.3" />
          </svg>
          Switch Row/<br />Column
        </button>
        <button type="button" className="ribbon-command ribbon-large" aria-label="Select Data" onClick={onOpenSelectData}>
          <svg className="ribbon-glyph" width={28} height={28} viewBox="0 0 24 24" aria-hidden="true">
            <rect x="3" y="4" width="18" height="16" rx="1" fill="#1d222b" stroke="#9aa3b2" strokeWidth="1.2" />
            <path d="M3 9 H21 M9 4 V20 M15 4 V20" stroke="#5b6473" strokeWidth="1" />
            <rect x="10" y="10" width="4" height="5" fill="#4ea1ff" />
          </svg>
          Select<br />Data
        </button>
      </fieldset>

      <fieldset>
        <legend>Type</legend>
        <button type="button" className="ribbon-command ribbon-large" aria-label="Change Chart Type" onClick={onOpenChangeChartType}>
          <svg className="ribbon-glyph" width={28} height={28} viewBox="0 0 24 24" aria-hidden="true">
            <rect x="3" y="10" width="4" height="9" fill="#a5aab4" stroke="#454b56" />
            <rect x="9" y="5" width="4" height="14" fill="#4ea1ff" stroke="#1f4e8c" />
            <rect x="15" y="13" width="4" height="6" fill="#d4d8df" stroke="#454b56" />
            <path d="M13 4 L19 9 L11 17 L9 17.5 L9.5 15.5 Z" fill="#f7c331" stroke="#6b520c" strokeWidth="0.9" />
          </svg>
          Change<br />Chart Type
        </button>
      </fieldset>

      <fieldset>
        <legend>Location</legend>
        <button type="button" className="ribbon-command ribbon-large" aria-label="Move Chart" onClick={onOpenMoveChart}>
          <svg className="ribbon-glyph" width={28} height={28} viewBox="0 0 24 24" aria-hidden="true">
            <rect x="3" y="5" width="12" height="11" fill="none" stroke="#9aa3b2" strokeWidth="1.3" />
            <rect x="9" y="8" width="12" height="11" fill="#1d222b" stroke="#4ea1ff" strokeWidth="1.4" />
            <path d="M12 15 L17 11 M14 11 H17 V14" fill="none" stroke="#d8dee9" strokeWidth="1.4" />
          </svg>
          Move<br />Chart
        </button>
      </fieldset>

      {chart.source.kind === 'pivot' && (
        <fieldset>
          <legend>Show/Hide</legend>
          <button
            type="button"
            className="ribbon-command ribbon-large pivot-ribbon-toggle"
            aria-label="Field Buttons"
            aria-pressed={chart.showFieldButtons !== false}
            onClick={() => onUpdateChart({ showFieldButtons: chart.showFieldButtons === false })}
          >
            <svg className="ribbon-glyph" width={28} height={28} viewBox="0 0 24 24" aria-hidden="true">
              <rect x="3" y="4" width="18" height="6" rx="1" fill="#2a313d" stroke="#8aa8cc" />
              <path d="M15 7 H19 L17 9 Z" fill="#d8dee9" />
              <rect x="4" y="13" width="4" height="7" fill="#4ea1ff" />
              <rect x="10" y="11" width="4" height="9" fill="#f38c4e" />
              <rect x="16" y="14" width="4" height="6" fill="#a5aab4" />
            </svg>
            Field<br />Buttons
          </button>
        </fieldset>
      )}
    </>
  );
}
