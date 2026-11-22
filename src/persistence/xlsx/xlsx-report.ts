/** Desktop exchange codes from XlsxExchange.ExchangeCodes plus the codes used by the style/gradient readers. */
export const exchangeCodes = Object.freeze({
  stableIds: 'stable-ids',
  unevaluatedFormulas: 'unevaluated-formulas',
  omittedObjects: 'omitted-objects',
  missingCache: 'missing-cache',
  newIds: 'new-ids',
  definedNames: 'defined-names',
  mergedCells: 'merged-cells',
  tables: 'tables',
  validations: 'validations',
  conditionalFormats: 'conditional-formats',
  pictures: 'pictures',
  pivots: 'pivots',
  macros: 'macros',
  charts: 'charts',
  drawings: 'drawings',
  comments: 'comments',
  externalLinks: 'external-links',
  hiddenSheet: 'hidden-sheet',
  arrayFormula: 'array-formula',
  constantError: 'constant-error',
  dateSerial: 'date-serial',
  dateSystem: 'date-system',
  unsupportedFormulas: 'unsupported-formulas',
  filterFreeze: 'filter-freeze',
  pythonSnapshots: 'python-snapshots',
  dimensionLimits: 'dimension-limits',
  columnWidth: 'column-width',
  cacheFreshness: 'cache-freshness',
  themeFill: 'theme-fill',
  themeFont: 'theme-font',
  borderColors: 'border-colors',
  borderStyles: 'border-styles',
  fontSize: 'font-size',
  fontEffects: 'font-effects',
  richText: 'rich-text',
  hyperlinks: 'hyperlinks',
  numberFormat: 'number-format',
  axisStyles: 'axis-styles',
  defaultDimensions: 'default-dimensions',
  freezePanes: 'freeze-panes',
  protection: 'protection',
  gradientFill: 'gradient-fill',
  patternFill: 'pattern-fill',
  sharedFormula: 'shared-formula',
} as const);

export type ExchangeCode = typeof exchangeCodes[keyof typeof exchangeCodes] | string;

export interface ExchangeIssue {
  readonly code: ExchangeCode;
  readonly detail: string;
}

/** Import/export loss report. Each code is reported once, as in the desktop ExchangeReport. */
export class XlsxReport {
  readonly #issues: ExchangeIssue[];

  constructor(issues: readonly ExchangeIssue[] = []) {
    this.#issues = [...issues];
  }

  get issues(): readonly ExchangeIssue[] { return this.#issues; }

  add(code: ExchangeCode, detail: string): void {
    if (this.#issues.some(issue => issue.code === code)) return;
    this.#issues.push({ code, detail });
  }

  has(code: ExchangeCode): boolean {
    return this.#issues.some(issue => issue.code === code);
  }

  get summary(): string {
    return this.#issues.length === 0 ? 'No unsupported content.' :
      this.#issues.map(issue => issue.code + ': ' + issue.detail).join('\n');
  }
}
