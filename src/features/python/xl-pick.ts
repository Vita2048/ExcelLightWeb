import { pythonReferenceSlotAt } from '../../python/references';

export interface TextSpan {
  readonly start: number;
  readonly end: number;
}

export interface XlReferenceEdit {
  readonly text: string;
  /** Caret position just beyond the closing quote. */
  readonly caret: number;
}

/**
 * The caret is inside an empty or populated xl("...") reference string. This mirrors the
 * desktop editor test: picking is intentional only while the caret sits in that literal.
 */
export function xlReferenceSlot(source: string, offset: number): TextSpan | undefined {
  const slot = pythonReferenceSlotAt(source, offset);
  return slot && { start: slot.start, end: slot.end };
}

export interface XlReferenceInsertion extends XlReferenceEdit {
  /** Offsets replaced by the editor range; the closing quote is not included. */
  readonly start: number;
  readonly end: number;
  readonly inserted: string;
}

function escaped(reference: string, quote: string): string {
  return reference.replaceAll('\\', '\\\\').replaceAll(quote, '\\' + quote);
}

/** Replace only the reference inside the current xl literal for a Monaco executeEdits range. */
export function xlReferenceInsertion(source: string, offset: number, reference: string): XlReferenceInsertion | undefined {
  const slot = pythonReferenceSlotAt(source, offset);
  if (!slot) return undefined;
  // Raw and triple-quoted literals cannot always represent a picked sheet name with
  // ordinary backslash escaping. Replace that literal with a normal escaped string.
  const reformat = slot.raw || slot.triple;
  const start = reformat ? slot.literalStart : slot.start;
  const end = reformat ? slot.literalEnd : slot.end;
  const inserted = reformat ? JSON.stringify(reference) : escaped(reference, slot.quote);
  const caret = reformat ? start + inserted.length : start + inserted.length + 1;
  const text = source.slice(0, start) + inserted + source.slice(end);
  return { text, start, end, inserted, caret };
}

/**
 * Replace only the reference inside the current xl literal. Returns undefined when the
 * caret is not in a literal, so ordinary later cell selections never hijack the editor.
 */
export function insertXlReference(source: string, offset: number, reference: string): XlReferenceEdit | undefined {
  const insertion = xlReferenceInsertion(source, offset, reference);
  return insertion && { text: insertion.text, caret: insertion.caret };
}
