const xlwsFunctions = new Set(['FILTER', 'SORT']);

const xlfnFunctions = new Set([
  'AGGREGATE', 'ARRAYTOTEXT', 'BYCOL', 'BYROW', 'CEILING.MATH', 'CHOOSECOLS', 'CHOOSEROWS',
  'CONCAT', 'COVARIANCE.S', 'DAYS', 'DROP', 'EXPAND', 'FLOOR.MATH', 'FORECAST.LINEAR',
  'FORMULATEXT', 'HSTACK', 'IFS', 'ISFORMULA', 'ISOMITTED', 'ISOWEEKNUM', 'LAMBDA', 'LET',
  'MAKEARRAY', 'MAP', 'MAXIFS', 'MINIFS', 'MODE.SNGL', 'NETWORKDAYS.INTL', 'NUMBERVALUE',
  'PERCENTILE.EXC', 'PERCENTILE.INC', 'QUARTILE.EXC', 'QUARTILE.INC', 'RANDARRAY', 'RANK.AVG',
  'RANK.EQ', 'REDUCE', 'SCAN', 'SEQUENCE', 'SORTBY', 'STDEV.P', 'STDEV.S', 'SWITCH',
  'TAKE', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTJOIN', 'TEXTSPLIT', 'TOCOL', 'TOROW',
  'UNICHAR', 'UNICODE', 'UNIQUE', 'VALUETOTEXT', 'VAR.P', 'VAR.S', 'VSTACK',
  'WORKDAY.INTL', 'WRAPCOLS', 'WRAPROWS', 'XLOOKUP', 'XMATCH', 'XOR',
]);

function transformFormulaIdentifiers(source: string, mapCallName: (prefix: string, name: string) => string): string {
  let out = '';
  let i = 0;
  while (i < source.length) {
    const ch = source[i]!;
    if (ch === '"') {
      const start = i++;
      while (i < source.length) {
        if (source[i] === '"' && source[i + 1] === '"') i += 2;
        else if (source[i] === '"') { i++; break; }
        else i++;
      }
      out += source.slice(start, i);
      continue;
    }
    if (ch === "'") {
      const start = i++;
      while (i < source.length) {
        if (source[i] === "'" && source[i + 1] === "'") i += 2;
        else if (source[i] === "'") { i++; break; }
        else i++;
      }
      out += source.slice(start, i);
      continue;
    }
    if (/[\p{L}_]/u.test(ch)) {
      const start = i++;
      while (i < source.length && /[\p{L}\p{N}_.]/u.test(source[i]!)) i++;
      const token = source.slice(start, i);
      let ws = i;
      while (ws < source.length && /\s/.test(source[ws]!)) ws++;
      if (source[ws] === '(') {
        const mapped = mapCallName('', token);
        out += mapped;
      } else {
        out += token;
      }
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

/** Strips OOXML `_xlfn._xlws.` and `_xlfn.` prefixes outside quoted strings and sheet identifiers. */
export function stripXlfnPrefixes(source: string): string {
  return transformFormulaIdentifiers(source, (_, token) =>
    token.replace(/^_xlfn\.(?:_xlws\.)?/iu, ''));
}

/** Re-applies OOXML `_xlfn.` / `_xlfn._xlws.` prefixes to post-2007 Excel functions on XLSX export. */
export function addXlfnPrefixes(source: string): string {
  return transformFormulaIdentifiers(source, (_, token) => {
    if (/^_xlfn\./iu.test(token)) return token;
    const upper = token.toUpperCase();
    if (xlwsFunctions.has(upper)) return `_xlfn._xlws.${token}`;
    if (xlfnFunctions.has(upper)) return `_xlfn.${token}`;
    return token;
  });
}
