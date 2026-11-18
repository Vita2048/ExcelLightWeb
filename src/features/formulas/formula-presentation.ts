export function localeArgumentSeparator(locale?: string): ',' | ';' {
  return new Intl.NumberFormat(locale).formatToParts(1.5).some(part => part.type === 'decimal' && part.value === ',') ? ';' : ',';
}

interface Frame { readonly name: string; argument: number }

export function scanFormula(source: string, end: number, separators: string, onSeparator?: (index: number) => void) {
  const frames: Frame[] = [];
  if (!source.startsWith('=')) return frames;
  let quote = '', brackets = 0, braces = 0;
  for (let i = 1; i < end; i++) {
    const character = source[i]!;
    if (quote) {
      if (character === quote) {
        if (i + 1 < end && source[i + 1] === quote) i++;
        else quote = '';
      }
      continue;
    }
    if (character === '"' || character === "'") { quote = character; continue; }
    if (character === '[') brackets++;
    if (character === ']') brackets = Math.max(0, brackets - 1);
    if (brackets) continue;
    if (character === '{') braces++;
    if (character === '}') braces = Math.max(0, braces - 1);
    if (braces) continue;
    if (character === '(') {
      const token = /[\p{L}\p{N}_.]+$/u.exec(source.slice(1, i).trimEnd())?.[0] ?? '';
      frames.push({ name: /^\p{L}/u.test(token) ? token.toUpperCase() : '', argument: 0 });
    } else if (character === ')') frames.pop();
    else if (separators.includes(character)) {
      const frame = frames.at(-1);
      if (!frame?.name) continue;
      frame.argument++;
      onSeparator?.(i);
    }
  }
  return frames;
}

export function presentFormula(source: string, from: ',' | ';', to: ',' | ';'): string {
  if (from === to) return source;
  let previous = 0;
  const parts: string[] = [];
  scanFormula(source, source.length, from, index => {
    parts.push(source.slice(previous, index), to);
    previous = index + 1;
  });
  return parts.join('') + source.slice(previous);
}
