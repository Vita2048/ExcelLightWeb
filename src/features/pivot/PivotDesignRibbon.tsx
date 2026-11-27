import { pivotStyles, type PivotStyleId, type PivotTableDocument } from '../../core/pivot';
import { RibbonFlyout } from '../shell/RibbonFlyout';

export interface PivotDesignRibbonProps {
  readonly pivot: PivotTableDocument;
  readonly onUpdateLayout: (patch: Partial<Omit<PivotTableDocument, 'id' | 'extent'>>) => void;
}

const styleSwatchColors: Readonly<Record<PivotStyleId, { readonly header: string; readonly band: string; readonly total: string }>> = {
  Blue: { header: '#4f81bd', band: '#28394f', total: '#1e4471' },
  Green: { header: '#548235', band: '#243a2b', total: '#214e36' },
  Orange: { header: '#c65911', band: '#3d281a', total: '#693714' },
  Purple: { header: '#8064a2', band: '#322645', total: '#462d66' },
};

export function PivotDesignRibbon({ pivot, onUpdateLayout }: PivotDesignRibbonProps) {
  const { styleOptions, grandTotals } = pivot;

  return <>
    <fieldset>
      <legend>Layout</legend>
      <RibbonFlyout
        label="Subtotals"
        icon="Subtotals"
        large
        items={[
          {
            label: 'Do Not Show Subtotals',
            checked: pivot.subtotals === 'off',
            select: () => onUpdateLayout({ subtotals: 'off' }),
          },
          {
            label: 'Show all Subtotals at Bottom of Group',
            checked: pivot.subtotals === 'bottom',
            select: () => onUpdateLayout({ subtotals: 'bottom' }),
          },
          {
            label: 'Show all Subtotals at Top of Group',
            checked: pivot.subtotals === 'top',
            select: () => onUpdateLayout({ subtotals: 'top' }),
          },
        ]}
      />
      <RibbonFlyout
        label="Grand Totals"
        icon="Grand Totals"
        large
        items={[
          {
            label: 'Off for Rows and Columns',
            checked: !grandTotals.rows && !grandTotals.columns,
            select: () => onUpdateLayout({ grandTotals: { rows: false, columns: false } }),
          },
          {
            label: 'On for Rows and Columns',
            checked: grandTotals.rows && grandTotals.columns,
            select: () => onUpdateLayout({ grandTotals: { rows: true, columns: true } }),
          },
          {
            label: 'On for Rows Only',
            checked: grandTotals.rows && !grandTotals.columns,
            select: () => onUpdateLayout({ grandTotals: { rows: true, columns: false } }),
          },
          {
            label: 'On for Columns Only',
            checked: !grandTotals.rows && grandTotals.columns,
            select: () => onUpdateLayout({ grandTotals: { rows: false, columns: true } }),
          },
        ]}
      />
      <RibbonFlyout
        label="Report Layout"
        icon="Report Layout"
        large
        items={[
          {
            label: 'Show in Compact Form',
            checked: pivot.reportLayout === 'compact',
            select: () => onUpdateLayout({ reportLayout: 'compact' }),
          },
          {
            label: 'Show in Outline Form',
            checked: pivot.reportLayout === 'outline',
            select: () => onUpdateLayout({ reportLayout: 'outline' }),
          },
          {
            label: 'Show in Tabular Form',
            checked: pivot.reportLayout === 'tabular',
            select: () => onUpdateLayout({ reportLayout: 'tabular' }),
          },
          {
            label: 'Repeat All Item Labels',
            divider: true,
            checked: Boolean(pivot.repeatLabels),
            select: () => onUpdateLayout({ repeatLabels: true }),
          },
          {
            label: 'Do Not Repeat Item Labels',
            checked: !pivot.repeatLabels,
            select: () => onUpdateLayout({ repeatLabels: false }),
          },
        ]}
      />
      <RibbonFlyout
        label="Blank Rows"
        icon="Blank Rows"
        large
        items={[
          {
            label: 'Insert Blank Line after Each Item',
            checked: Boolean(pivot.blankRows),
            select: () => onUpdateLayout({ blankRows: true }),
          },
          {
            label: 'Remove Blank Line after Each Item',
            checked: !pivot.blankRows,
            select: () => onUpdateLayout({ blankRows: false }),
          },
        ]}
      />
    </fieldset>

    <fieldset>
      <legend>PivotTable Style Options</legend>
      <div className="pivot-style-options-grid">
        <label>
          <input
            type="checkbox"
            checked={styleOptions.rowHeaders}
            onChange={event => onUpdateLayout({ styleOptions: { ...styleOptions, rowHeaders: event.currentTarget.checked } })}
          />
          Row Headers
        </label>
        <label>
          <input
            type="checkbox"
            checked={styleOptions.columnHeaders}
            onChange={event => onUpdateLayout({ styleOptions: { ...styleOptions, columnHeaders: event.currentTarget.checked } })}
          />
          Column Headers
        </label>
        <label>
          <input
            type="checkbox"
            checked={styleOptions.bandedRows}
            onChange={event => onUpdateLayout({ styleOptions: { ...styleOptions, bandedRows: event.currentTarget.checked } })}
          />
          Banded Rows
        </label>
        <label>
          <input
            type="checkbox"
            checked={styleOptions.bandedColumns}
            onChange={event => onUpdateLayout({ styleOptions: { ...styleOptions, bandedColumns: event.currentTarget.checked } })}
          />
          Banded Columns
        </label>
      </div>
    </fieldset>

    <fieldset className="table-styles-group">
      <legend>PivotTable Styles</legend>
      <div className="table-styles-row" role="listbox" aria-label="PivotTable Styles">
        {pivotStyles.map(styleId => {
          const swatch = styleSwatchColors[styleId];
          const selected = pivot.style === styleId;
          return (
            <button
              key={styleId}
              type="button"
              role="option"
              aria-selected={selected}
              aria-label={`Pivot Style ${styleId}`}
              title={styleId}
              className={'table-style-preview-btn' + (selected ? ' selected' : '')}
              onClick={() => onUpdateLayout({ style: styleId })}
            >
              <div className="table-style-swatch">
                <div className="table-style-swatch-header" style={{ backgroundColor: swatch.header }} />
                <div className="table-style-swatch-row" style={{ backgroundColor: swatch.band }} />
                <div className="table-style-swatch-row" style={{ backgroundColor: swatch.total }} />
                <div className="table-style-swatch-row" style={{ backgroundColor: swatch.band }} />
                <div className="table-style-swatch-row" style={{ backgroundColor: swatch.total }} />
              </div>
            </button>
          );
        })}
      </div>
    </fieldset>
  </>;
}
