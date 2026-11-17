import { DOMParser, XMLSerializer, onErrorStopParsing, type Element } from '@xmldom/xmldom';
import { readProbeParts, writeProbeParts } from './bounded-zip';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const parser = new DOMParser({ onError: onErrorStopParsing });
const decoder = new TextDecoder('utf-8', { fatal: true });
const serializer = new XMLSerializer();
const edges = ['top', 'right', 'bottom', 'left', 'diagonal'] as const;
const styles = new Set(['hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin', 'mediumDashDotDot',
  'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double']);
type Pen = { style: string; color: string | null };
type Border = { up: boolean; down: boolean; pens: Partial<Record<typeof edges[number], Pen>> };
type Gradient = {
  stops: { position: number; argb: string }[];
  geometry: { type: 'linear'; degree: number } | { type: 'path'; offsets: [number, number, number, number] };
};
type Rotation = number | 'stacked';
const offsets = ['left', 'right', 'top', 'bottom'] as const;

function children(parent: Element, localName: string) {
  return Array.from(parent.childNodes).filter((node): node is Element =>
    node.nodeType === 1 && (node as Element).namespaceURI === ns && (node as Element).localName === localName);
}

function numberAttr(node: Element, key: string, fallback: number) {
  const text = node.getAttribute(key);
  const number = text === null ? fallback : text.trim() === '' ? NaN : Number(text);
  if (!Number.isFinite(number)) throw new Error(`Invalid XLSX style attribute: ${key}`);
  return number;
}

function boolAttr(node: Element, key: string) {
  const text = node.getAttribute(key);
  return text === '1' || text === 'true';
}

function readBorder(node: Element): Border {
  const pens: Border['pens'] = {};
  for (const edge of edges) {
    const record = children(node, edge)[0];
    const style = record?.getAttribute('style');
    if (record && style && style !== 'none') {
      if (!styles.has(style)) throw new Error(`Unsupported XLSX border: ${style}`);
      const color = children(record, 'color')[0];
      const argb = color?.getAttribute('rgb') ?? null;
      if (color && !argb && !boolAttr(color, 'auto') && color.getAttribute('indexed') !== '64') {
        throw new Error('Unsupported XLSX border color');
      }
      if (argb && !/^[0-9a-f]{8}$/iu.test(argb)) throw new Error('Invalid XLSX border color');
      pens[edge] = { style, color: argb?.toUpperCase() ?? null };
    }
  }
  return { up: boolAttr(node, 'diagonalUp'), down: boolAttr(node, 'diagonalDown'), pens };
}

function readGradient(node: Element): Gradient {
  const stops = children(node, 'stop').map(stop => {
    const argb = children(stop, 'color')[0]?.getAttribute('rgb');
    if (!argb || !/^[0-9a-f]{8}$/iu.test(argb)) throw new Error('Unsupported XLSX gradient color');
    return { position: numberAttr(stop, 'position', NaN), argb: argb.toUpperCase() };
  });
  if ((stops.length !== 2 && stops.length !== 3) || stops[0]?.position !== 0 ||
      stops.at(-1)?.position !== 1 || (stops.length === 3 &&
        (stops[1]?.position !== 0.5 || stops[2]?.argb !== stops[0]?.argb))) {
    throw new Error('Unsupported XLSX gradient stops');
  }
  const kind = node.getAttribute('type') ?? 'linear';
  if (kind === 'path') {
    const [left, right, top, bottom] = offsets.map(key => numberAttr(node, key, 0));
    if (left !== right || top !== bottom || !((left === 0.25 || left === 0.5) && left === top) &&
        !([0, 1].includes(left!) && [0, 1].includes(top!))) {
      throw new Error('Unsupported XLSX path gradient');
    }
    return { stops, geometry: { type: 'path', offsets: [left!, right!, top!, bottom!] } };
  }
  const degree = numberAttr(node, 'degree', 0);
  if (kind !== 'linear' || ![0, 45, 90, 135, 180, 225, 270, 315].includes(degree)) {
    throw new Error('Unsupported XLSX linear gradient');
  }
  return { stops, geometry: { type: 'linear', degree } };
}

function readRotation(node: Element): Rotation {
  const raw = numberAttr(node, 'textRotation', 0);
  if (raw === 255) return 'stacked';
  if (!Number.isInteger(raw) || raw < 0 || raw > 180) throw new Error('Invalid XLSX rotation');
  return raw > 90 ? 90 - raw : raw;
}

function readDocument(xml: string) {
  if (/<!DOCTYPE/iu.test(xml)) throw new Error('XLSX XML document types are outside the probe');
  const document = parser.parseFromString(xml, 'application/xml');
  const root = document.documentElement;
  if (!root || root.namespaceURI !== ns || root.localName !== 'styleSheet') throw new Error('Invalid XLSX styles part');
  return { document, root };
}

function styleRecords(root: Element) {
  const records = (container: string, name: string) => children(root, container).flatMap(parent => children(parent, name));
  return {
    borders: records('borders', 'border'),
    gradients: records('fills', 'fill').flatMap(fill => children(fill, 'gradientFill')),
    rotations: records('cellXfs', 'xf').flatMap(format => children(format, 'alignment'))
      .filter(alignment => alignment.hasAttribute('textRotation')),
  };
}

/** Decode and re-emit the selected records while keeping style indexes and neighboring fields. */
export function roundTripStyles(xml: string) {
  const { document, root } = readDocument(xml);
  const records = styleRecords(root);
  const borders = records.borders.map(readBorder);
  const gradients = records.gradients.map(readGradient);
  const rotations = records.rotations.map(readRotation);

  for (const [i, border] of borders.entries()) {
    const element = records.borders[i]!;
    for (const [key, enabled] of [['diagonalUp', border.up], ['diagonalDown', border.down]] as const) {
      if (enabled) element.setAttribute(key, '1');
      else element.removeAttribute(key);
    }
    for (const edge of edges) {
      const pen = border.pens[edge];
      if (!pen) continue;
      const target = children(element, edge)[0]!;
      target.setAttribute('style', pen.style);
      for (const color of children(target, 'color')) target.removeChild(color);
      const color = document.createElementNS(ns, 'color');
      color.setAttribute(pen.color ? 'rgb' : 'indexed', pen.color ?? '64');
      target.appendChild(color);
    }
  }
  for (const [i, gradient] of gradients.entries()) {
    const element = records.gradients[i]!;
    const geometry = gradient.geometry;
    if (geometry.type === 'path') {
      offsets.forEach((key, j) => element.setAttribute(key, String(geometry.offsets[j])));
    } else {
      element.setAttribute('degree', String(geometry.degree));
    }
    for (const stop of children(element, 'stop')) element.removeChild(stop);
    for (const stop of gradient.stops) {
      const node = document.createElementNS(ns, 'stop');
      node.setAttribute('position', String(stop.position));
      const color = document.createElementNS(ns, 'color');
      color.setAttribute('rgb', stop.argb);
      node.appendChild(color);
      element.appendChild(node);
    }
  }
  for (const [i, rotation] of rotations.entries()) {
    records.rotations[i]!.setAttribute('textRotation', String(rotation === 'stacked' ? 255 : rotation < 0 ? 90 - rotation : rotation));
  }
  return { xml: serializer.serializeToString(document), borders, gradients, rotations };
}

/** Known XLSX fixture only. S19 owns content-type checks, loss reports and workbook conversion. */
export async function rewriteXlsxProbe(blob: Blob) {
  const parts = await readProbeParts(blob);
  const styles = parts.get('xl/styles.xml');
  if (!styles) throw new Error('Missing XLSX styles part');
  const { xml, ...stylesRead } = roundTripStyles(decoder.decode(styles));
  parts.set('xl/styles.xml', new TextEncoder().encode(xml));
  return { ...stylesRead, blob: await writeProbeParts(parts) };
}
