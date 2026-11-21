import { PackageException } from './package-exception';

export function parseStrictJson(text: string): unknown {
  let pos = 0;
  const len = text.length;
  let depth = 0;

  function skipWhitespace(): void {
    while (pos < len) {
      const ch = text.charCodeAt(pos);
      if (ch === 0x20 || ch === 0x09 || ch === 0x0a || ch === 0x0d) {
        pos++;
      } else {
        break;
      }
    }
  }

  function error(msg = 'Malformed package JSON.'): never {
    throw new PackageException(msg);
  }

  function parseValue(): unknown {
    skipWhitespace();
    if (pos >= len) error();
    const ch = text.charCodeAt(pos);
    if (ch === 0x22) { // '"'
      return parseString();
    }
    if (ch === 0x7b) { // '{'
      return parseObject();
    }
    if (ch === 0x5b) { // '['
      return parseArray();
    }
    if (ch === 0x74) { // 't' -> true
      if (text.startsWith('true', pos)) {
        pos += 4;
        return true;
      }
      error();
    }
    if (ch === 0x66) { // 'f' -> false
      if (text.startsWith('false', pos)) {
        pos += 5;
        return false;
      }
      error();
    }
    if (ch === 0x6e) { // 'n' -> null
      if (text.startsWith('null', pos)) {
        pos += 4;
        return null;
      }
      error();
    }
    if (ch === 0x2d || (ch >= 0x30 && ch <= 0x39)) { // '-' or 0-9
      return parseNumber();
    }
    error();
  }

  function parseString(): string {
    if (text.charCodeAt(pos) !== 0x22) error();
    pos++; // skip opening quote
    let start = pos;
    let result = '';
    while (pos < len) {
      const ch = text.charCodeAt(pos);
      if (ch === 0x22) { // '"' closing
        result += text.slice(start, pos);
        pos++;
        return result;
      }
      if (ch < 0x20) { // unescaped control character
        error();
      }
      if (ch === 0x5c) { // '\' escape
        result += text.slice(start, pos);
        pos++;
        if (pos >= len) error();
        const esc = text.charCodeAt(pos++);
        switch (esc) {
          case 0x22: result += '"'; break;
          case 0x5c: result += '\\'; break;
          case 0x2f: result += '/'; break;
          case 0x62: result += '\b'; break;
          case 0x66: result += '\f'; break;
          case 0x6e: result += '\n'; break;
          case 0x72: result += '\r'; break;
          case 0x74: result += '\t'; break;
          case 0x75: { // '\uXXXX'
            if (pos + 4 > len) error();
            const hex = text.slice(pos, pos + 4);
            if (!/^[0-9a-fA-F]{4}$/.test(hex)) error();
            result += String.fromCharCode(parseInt(hex, 16));
            pos += 4;
            break;
          }
          default:
            error();
        }
        start = pos;
      } else {
        pos++;
      }
    }
    error();
  }

  function parseNumber(): number {
    const start = pos;
    if (text.charCodeAt(pos) === 0x2d) pos++; // '-'
    if (pos >= len) error();
    if (text.charCodeAt(pos) === 0x30) { // '0'
      pos++;
    } else if (text.charCodeAt(pos) >= 0x31 && text.charCodeAt(pos) <= 0x39) {
      pos++;
      while (pos < len && text.charCodeAt(pos) >= 0x30 && text.charCodeAt(pos) <= 0x39) pos++;
    } else {
      error();
    }
    if (pos < len && text.charCodeAt(pos) === 0x2e) { // '.'
      pos++;
      if (pos >= len || text.charCodeAt(pos) < 0x30 || text.charCodeAt(pos) > 0x39) error();
      while (pos < len && text.charCodeAt(pos) >= 0x30 && text.charCodeAt(pos) <= 0x39) pos++;
    }
    if (pos < len && (text.charCodeAt(pos) === 0x65 || text.charCodeAt(pos) === 0x45)) { // 'e' or 'E'
      pos++;
      if (pos < len && (text.charCodeAt(pos) === 0x2b || text.charCodeAt(pos) === 0x2d)) pos++; // '+' or '-'
      if (pos >= len || text.charCodeAt(pos) < 0x30 || text.charCodeAt(pos) > 0x39) error();
      while (pos < len && text.charCodeAt(pos) >= 0x30 && text.charCodeAt(pos) <= 0x39) pos++;
    }
    const numStr = text.slice(start, pos);
    const val = Number(numStr);
    return val;
  }

  function parseObject(): Record<string, unknown> {
    if (text.charCodeAt(pos) !== 0x7b) error();
    pos++; // skip '{'
    depth++;
    if (depth > 16) error();
    const obj: Record<string, unknown> = {};
    const seenKeys = new Set<string>();
    skipWhitespace();
    if (pos < len && text.charCodeAt(pos) === 0x7d) { // '}'
      pos++;
      depth--;
      return obj;
    }
    while (pos < len) {
      skipWhitespace();
      if (pos >= len || text.charCodeAt(pos) !== 0x22) error();
      const key = parseString();
      if (seenKeys.has(key)) {
        throw new PackageException('Duplicate JSON property.');
      }
      seenKeys.add(key);
      skipWhitespace();
      if (pos >= len || text.charCodeAt(pos) !== 0x3a) error(); // ':'
      pos++; // skip ':'
      const value = parseValue();
      obj[key] = value;
      skipWhitespace();
      if (pos >= len) error();
      if (text.charCodeAt(pos) === 0x7d) { // '}'
        pos++;
        depth--;
        return obj;
      }
      if (text.charCodeAt(pos) === 0x2c) { // ','
        pos++;
        skipWhitespace();
        if (pos < len && text.charCodeAt(pos) === 0x7d) { // trailing comma!
          error();
        }
      } else {
        error();
      }
    }
    error();
  }

  function parseArray(): unknown[] {
    if (text.charCodeAt(pos) !== 0x5b) error();
    pos++; // skip '['
    depth++;
    if (depth > 16) error();
    const arr: unknown[] = [];
    skipWhitespace();
    if (pos < len && text.charCodeAt(pos) === 0x5d) { // ']'
      pos++;
      depth--;
      return arr;
    }
    while (pos < len) {
      const val = parseValue();
      arr.push(val);
      skipWhitespace();
      if (pos >= len) error();
      if (text.charCodeAt(pos) === 0x5d) { // ']'
        pos++;
        depth--;
        return arr;
      }
      if (text.charCodeAt(pos) === 0x2c) { // ','
        pos++;
        skipWhitespace();
        if (pos < len && text.charCodeAt(pos) === 0x5d) { // trailing comma!
          error();
        }
      } else {
        error();
      }
    }
    error();
  }

  const result = parseValue();
  skipWhitespace();
  if (pos !== len) error();
  return result;
}
