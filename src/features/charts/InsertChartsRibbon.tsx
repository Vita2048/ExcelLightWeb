import { useId, useLayoutEffect, useRef, useState, type KeyboardEvent } from 'react';
import { createPortal } from 'react-dom';
import type { ChartType } from '../../core/chart';
import { positionRibbonPopup } from '../shell/RibbonFlyout';
import {
  ChartTypeGlyph,
  RecommendedChartsGlyph,
  coreChartFlyoutFamilies,
  type ChartFlyoutFamily,
  type CoreChartCategoryTab,
} from './chart-catalog';

function ChartCategoryIcon({ index }: { readonly index: number }) {
  if (index === 0) {
    return (
      <svg className="ribbon-glyph" width={20} height={20} viewBox="0 0 24 24" aria-hidden="true">
        <path d="M3 20 H21" stroke="#b0b6c0" strokeWidth="1.4" />
        <rect x="4" y="11" width="4" height="8" fill="#a5aab4" stroke="#454b56" />
        <rect x="8" y="5" width="4" height="14" fill="#d4d8df" stroke="#454b56" />
        <rect x="13" y="8" width="4" height="11" fill="#4ea1ff" stroke="#1f4e8c" />
      </svg>
    );
  }
  if (index === 1) {
    return (
      <svg className="ribbon-glyph" width={20} height={20} viewBox="0 0 24 24" aria-hidden="true">
        <path d="M3 20 H21 M4 4 V20" stroke="#b0b6c0" strokeWidth="1.2" />
        <polyline points="5,15 10,8 15,13 20,6" fill="none" stroke="#4ea1ff" strokeWidth="1.8" />
        <polyline points="5,18 10,13 15,16 20,11" fill="none" stroke="#a5aab4" strokeWidth="1.5" />
      </svg>
    );
  }
  if (index === 2) {
    return (
      <svg className="ribbon-glyph" width={20} height={20} viewBox="0 0 24 24" aria-hidden="true">
        <circle cx="12" cy="12" r="8" fill="#a5aab4" stroke="#454b56" strokeWidth="1.2" />
        <path d="M12 12 L12 4 A8 8 0 0 1 19 16 Z" fill="#4ea1ff" stroke="#1f4e8c" strokeWidth="1.2" />
      </svg>
    );
  }
  return (
    <svg className="ribbon-glyph" width={20} height={20} viewBox="0 0 24 24" aria-hidden="true">
      <path d="M3 20 H21 M4 4 V20" stroke="#b0b6c0" strokeWidth="1.2" />
      <circle cx="8" cy="15" r="2" fill="#4ea1ff" />
      <circle cx="13" cy="10" r="2" fill="#a5aab4" />
      <circle cx="18" cy="7" r="2" fill="#4ea1ff" />
    </svg>
  );
}

function ChartGalleryFlyout({
  family,
  iconIndex,
  onInsertChart,
  onMoreCharts,
}: {
  readonly family: ChartFlyoutFamily;
  readonly iconIndex: number;
  readonly onInsertChart: (chartType: ChartType) => void;
  readonly onMoreCharts: (category: CoreChartCategoryTab) => void;
}) {
  const id = useId();
  const [open, setOpen] = useState(false);
  const trigger = useRef<HTMLButtonElement>(null);
  const menu = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    if (!open) return;
    const button = trigger.current!;
    const panel = menu.current!;
    const position = () => positionRibbonPopup(button, panel);
    function dismiss(event: PointerEvent | FocusEvent) {
      const target = event.target as Node;
      if (!panel.contains(target) && !button.contains(target)) {
        setOpen(false);
      }
    }
    function onScroll(event: Event) {
      if (event.target instanceof Node && panel.contains(event.target)) return;
      position();
    }
    position();
    panel.querySelector('button')?.focus({ preventScroll: true });
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

  function handleMenuKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    if (event.key === 'Escape') {
      event.preventDefault();
      close();
      return;
    }
    const buttons = [...event.currentTarget.querySelectorAll('button')];
    const current = buttons.indexOf(document.activeElement as HTMLButtonElement);
    const next = event.key === 'Home' ? 0
      : event.key === 'End' ? buttons.length - 1
      : event.key === 'ArrowRight' || event.key === 'ArrowDown' ? (current + 1) % buttons.length
      : event.key === 'ArrowLeft' || event.key === 'ArrowUp' ? (current + buttons.length - 1) % buttons.length
      : undefined;
    if (next !== undefined) {
      event.preventDefault();
      buttons[next]?.focus();
    }
    if (event.key === 'Tab') close();
  }

  const groups: { readonly group: string; readonly options: readonly typeof family.items[number][] }[] = [];
  for (const item of family.items) {
    const existing = groups.find(g => g.group === item.group);
    if (existing) (existing.options as typeof family.items[number][]).push(item);
    else groups.push({ group: item.group, options: [item] });
  }

  return (
    <>
      <button
        ref={trigger}
        type="button"
        className="ribbon-command chart-ribbon-trigger"
        aria-label={family.triggerLabel}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={open ? id : undefined}
        title={family.triggerLabel}
        onClick={() => { if (open) close(); else setOpen(true); }}
        onKeyDown={event => {
          if (event.key === 'ArrowDown') { event.preventDefault(); setOpen(true); }
          if (event.key === 'Escape' && open) { event.preventDefault(); close(); }
        }}
      >
        <ChartCategoryIcon index={iconIndex} />
        <span className="ribbon-chevron" aria-hidden="true" />
      </button>
      {open && createPortal(
        <div
          ref={menu}
          id={id}
          role="menu"
          aria-label={family.triggerLabel}
          className="ribbon-flyout chart-gallery-flyout"
          onKeyDown={handleMenuKeyDown}
        >
          {groups.map((section, sectionIdx) => (
            <div key={section.group} role="none" className={sectionIdx > 0 ? 'ribbon-flyout-divider' : undefined}>
              <div className="ribbon-flyout-heading" aria-hidden="true">{section.group}</div>
              <div className="chart-gallery-grid" role="none">
                {section.options.map(option => (
                  <button
                    key={option.type}
                    type="button"
                    role="menuitem"
                    className="chart-gallery-tile"
                    title={option.label}
                    aria-label={option.label}
                    onClick={() => {
                      close();
                      onInsertChart(option.type);
                    }}
                  >
                    <span className="ribbon-flyout-glyph chart-thumb-glyph" aria-hidden="true">
                      <ChartTypeGlyph type={option.type} size={36} />
                    </span>
                    <span className="ribbon-flyout-label chart-gallery-sr">{option.label}</span>
                  </button>
                ))}
              </div>
            </div>
          ))}
          <div role="none" className="ribbon-flyout-divider">
            <button
              type="button"
              role="menuitem"
              className="chart-gallery-more"
              onClick={() => {
                close();
                onMoreCharts(family.dialogCategory);
              }}
            >
              <span className="ribbon-flyout-glyph" aria-hidden="true">
                <ChartCategoryIcon index={iconIndex} />
              </span>
              <span className="ribbon-flyout-label">{family.moreLabel}</span>
            </button>
          </div>
        </div>,
        document.body,
      )}
    </>
  );
}

export function InsertChartsRibbon({
  onRecommendedCharts,
  onAllCharts,
  onInsertChart,
}: {
  readonly onRecommendedCharts: () => void;
  readonly onAllCharts: (category?: CoreChartCategoryTab) => void;
  readonly onInsertChart: (chartType: ChartType) => void;
}) {
  return (
    <fieldset className="insert-charts-group">
      <legend>Charts</legend>
      <button
        type="button"
        className="ribbon-command ribbon-large"
        onClick={onRecommendedCharts}
      >
        <RecommendedChartsGlyph size={30} />
        Recommended Charts
      </button>
      <div className="chart-ribbon-grid" role="group" aria-label="Chart types">
        {coreChartFlyoutFamilies.map((family, idx) => (
          <ChartGalleryFlyout
            key={family.triggerLabel}
            family={family}
            iconIndex={idx}
            onInsertChart={onInsertChart}
            onMoreCharts={category => onAllCharts(category)}
          />
        ))}
      </div>
      <button
        type="button"
        className="ribbon-dialog-launcher"
        aria-label="See All Charts"
        title="See All Charts"
        onClick={() => onAllCharts('Column')}
      >
        ↘
      </button>
    </fieldset>
  );
}
