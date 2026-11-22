import { DOMParser, onErrorStopParsing, type Element } from '@xmldom/xmldom';
import { PackageException } from '../native';

export type { Element } from '@xmldom/xmldom';

export const spreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const relationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const parser = new DOMParser({ onError: onErrorStopParsing });
const decoder = new TextDecoder('utf-8', { fatal: true });

export function decodePart(bytes: Uint8Array, label: string): string {
  try {
    return decoder.decode(bytes);
  } catch {
    throw new PackageException(`${label} is not valid UTF-8. No document was adopted.`);
  }
}

export interface ParsedPart {
  readonly root: Element;
  readonly text: string;
}

/**
 * Parse an OOXML part. Document type declarations are refused so entity expansion cannot be used
 * to inflate a package, and a malformed part never becomes a partial workbook.
 */
export function parseXmlPart(bytes: Uint8Array, label: string): ParsedPart {
  const text = decodePart(bytes, label);
  if (/<!DOCTYPE/iu.test(text)) {
    throw new PackageException(`${label} contains a document type declaration, which is not supported. No document was adopted.`);
  }
  let root: Element | undefined;
  try {
    root = parser.parseFromString(text, 'application/xml').documentElement ?? undefined;
  } catch {
    root = undefined;
  }
  if (!root) throw new PackageException(`${label} is not valid XML. No document was adopted.`);
  return { root, text };
}

export function childElements(parent: Element | undefined | null, localName?: string): Element[] {
  const result: Element[] = [];
  if (!parent) return result;
  for (let node = parent.firstChild; node; node = node.nextSibling) {
    if (node.nodeType !== 1) continue;
    const element = node as Element;
    if (!localName || element.localName === localName) result.push(element);
  }
  return result;
}

export function firstChild(parent: Element | undefined | null, localName: string): Element | undefined {
  return childElements(parent, localName)[0];
}

export function descendantElements(root: Element, localName: string): Element[] {
  const result: Element[] = [];
  const all = root.getElementsByTagName('*');
  for (let index = 0; index < all.length; index++) {
    const element = all.item(index);
    if (element?.localName === localName) result.push(element);
  }
  return result;
}

export function attr(element: Element | undefined | null, name: string): string | undefined {
  const value = element ? element.getAttribute(name) : null;
  return value === null || value === undefined || value === '' ? undefined : value;
}

export function textOf(element: Element | undefined): string | undefined {
  return element ? element.textContent ?? '' : undefined;
}

export function parseInvariantNumber(text: string, message: string): number {
  if (!/^[\t\n\v\f\r ]*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?[\t\n\v\f\r ]*$/u.test(text)) {
    throw new PackageException(`${message} No document was adopted.`);
  }
  const value = Number(text);
  if (!Number.isFinite(value)) throw new PackageException(`${message} No document was adopted.`);
  return value;
}

export function isTruthyAttribute(element: Element, name: string): boolean {
  const value = attr(element, name);
  return value === '1' || value === 'true';
}

export function escapeText(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}

export function escapeAttribute(value: string): string {
  return escapeText(value).replaceAll('"', '&quot;').replaceAll("'", '&apos;');
}
