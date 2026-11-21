import { entryText, parseEntry } from './editing-commands';
import { formatCellValue } from '../core/formats/cell-format';
import { columnLetters } from '../core/formula-references';
import { documentLimits, findCell, type CellDocument, type SheetDocument, type WorkbookDocument } from '../core/workbook';
import { requireDocument } from '../core/workbook-validation';
import type { WorkbookSnapshot } from './workbook-protocol';
import type { CellEdit, SetCellsCommand } from './workbook-session';

// Find and Replace mirrors EverydayCommands.SearchOptions/Find/Replace: one query, a source or
// display view, ordinal case and whole-cell modes, and one Replace transaction for every target.
export type SearchLook = 'formulas' | 'values';
export type SearchScope = 'sheet' | 'workbook';
export type SearchOrder = 'rows' | 'columns';
export interface SearchOptions {
  readonly query: string;
  readonly lookIn: SearchLook;
  readonly within: SearchScope;
  readonly order: SearchOrder;
  readonly matchCase: boolean;
  readonly wholeCell: boolean;
  /**
   * Replace searches the stored source of every cell, as Excel's Formulas mode does, so constants
   * remain replaceable. Find leaves this off and Formulas matches only cells that hold a formula.
   */
  readonly includeConstants?: boolean;
}
export interface SearchMatch {
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
}
export interface SearchHistory {
  readonly find: readonly string[];
  readonly replace: readonly string[];
}
export const emptySearchHistory: SearchHistory = Object.freeze({ find: [], replace: [] });
export const searchHistoryLimit = 15;

/** Cell text searched by the dialog: worksheet source, or the text the worksheet displays. */
export function searchText(cell: CellDocument | undefined, lookIn: SearchLook): string {
  if (!cell) return '';
  return lookIn === 'formulas' ? entryText(cell) : formatCellValue(cell).text;
}

function matches(text: string, options: SearchOptions): boolean {
  if (options.wholeCell) {
    return options.matchCase ? text === options.query : text.toLowerCase() === options.query.toLowerCase();
  }
  return options.matchCase ? text.includes(options.query) :
    text.toLowerCase().includes(options.query.toLowerCase());
}

function scopeSheets(document: WorkbookDocument, scope: readonly string[]): readonly SheetDocument[] {
  const sheets = scope.map(id => document.sheets.find(sheet => sheet.id === id));
  for (const sheet of sheets) requireDocument(sheet, 'The worksheet no longer exists.');
  return sheets as readonly SheetDocument[];
}

/** Ordered matches, sheet by sheet, then by row or by column inside each sheet. */
export function findMatches(document: WorkbookDocument, scope: readonly string[], options: SearchOptions): SearchMatch[] {
  requireDocument(options.query.length > 0, 'Enter text to find.');
  const found: SearchMatch[] = [];
  for (const sheet of scopeSheets(document, scope)) {
    const cells = sheet.cells.filter(cell => cell.content.kind !== 'blank' &&
      (options.lookIn !== 'formulas' || options.includeConstants || cell.content.kind === 'formula') &&
      matches(searchText(cell, options.lookIn), options));
    if (options.order === 'columns') cells.sort((left, right) => left.column - right.column || left.row - right.row);
    else cells.sort((left, right) => left.row - right.row || left.column - right.column);
    for (const cell of cells) found.push({ sheetId: sheet.id, row: cell.row, column: cell.column });
  }
  return found;
}

const address = (match: SearchMatch) => columnLetters(match.column) + match.row;
const absolute = (match: SearchMatch) => '$' + columnLetters(match.column) + '$' + match.row;

export function matchLabel(document: WorkbookDocument, match: SearchMatch, scope: SearchScope): string {
  if (scope === 'sheet') return address(match);
  const sheet = document.sheets.find(item => item.id === match.sheetId);
  return (sheet?.name ?? '') + '!' + address(match);
}

export interface SearchResultRow {
  readonly book: string;
  readonly sheet: string;
  readonly name: string;
  readonly cell: string;
  readonly value: string;
  readonly formula: string;
  readonly match: SearchMatch;
}

/** Rows for the desktop Find All table: Book, Sheet, Name, Cell, Value and Formula. */
export function searchResultRows(basis: WorkbookSnapshot, book: string, matches: readonly SearchMatch[],
  options: SearchOptions): SearchResultRow[] {
  return matches.map(match => {
    const sheet = basis.document.sheets.find(item => item.id === match.sheetId);
    const cell = sheet && findCell(sheet, match);
    return {
      book,
      sheet: sheet?.name ?? '',
      name: '',
      cell: absolute(match),
      value: searchText(cell, 'values'),
      formula: cell?.content.kind === 'formula' ? cell.content.source : '',
      match,
    };
  });
}

function replaceText(text: string, options: SearchOptions, replacement: string): string {
  if (options.wholeCell) return replacement;
  const escaped = options.query.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&');
  return text.replace(new RegExp(escaped, options.matchCase ? 'gu' : 'giu'), () => replacement);
}

/**
 * One Replace transaction. Every target is re-checked against the captured workbook, so a stale
 * result list refuses instead of overwriting a cell that no longer matches.
 */
export function replaceMatches(basis: WorkbookSnapshot, targets: readonly SearchMatch[], options: SearchOptions,
  replacement: string): SetCellsCommand {
  requireDocument(options.query.length > 0, 'Enter text to find.');
  const edits: CellEdit[] = targets.map(target => {
    const sheet = basis.document.sheets.find(item => item.id === target.sheetId);
    const cell = sheet && findCell(sheet, target);
    requireDocument(cell, 'Search result is no longer available.');
    const text = searchText(cell, options.lookIn);
    requireDocument(matches(text, options), 'Search result changed. Find again.');
    const replaced = replaceText(text, options, replacement);
    const content = cell.content.kind === 'text' ? { kind: 'text' as const, value: replaced } : parseEntry(replaced);
    return { target: { sheetId: target.sheetId, row: target.row, column: target.column }, content };
  });
  return { kind: 'set-cells', edits, label: 'Replace cells' };
}

/** Most-recent-first history with the desktop's 15-entry bound and exact deduplication. */
export function rememberSearch(history: readonly string[], value: string, replacement = false): readonly string[] {
  if (value.length > documentLimits.textLength || value.includes('\0') || !replacement && value.length === 0) return history;
  return [value, ...history.filter(entry => entry !== value)].slice(0, searchHistoryLimit);
}

export { matches as searchMatchesText };
