import { error } from '../formula-values';
import type { FormulaExpression } from '../../core/formula-references';
import { datetimeFunctions } from './datetime';
import { dynamicArrayFunctions } from './dynamic-array';
import { financialFunctions } from './financial';
import { infoFunctions } from './info';
import { logicalFunctions } from './logical';
import { lookupFunctions } from './lookup';
import { mathFunctions } from './math';
import { shapingFunctions } from './shaping';
import { statisticalFunctions } from './statistical';
import { textFunctions } from './text';
import type { FunctionContext, FunctionSignature, FunctionSpec } from './types';

export const ribbonInsertFunctions = [
  'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'IF', 'AND', 'OR', 'NOT', 'IFERROR',
  'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'SUMIF', 'COUNTIF', 'LEFT', 'RIGHT', 'MID',
  'LEN', 'TRIM', 'UPPER', 'LOWER', 'DATE', 'YEAR', 'MONTH', 'DAY', 'TODAY', 'VLOOKUP', 'SUBTOTAL',
] as const;

export const nativeFunctions = [
  ...ribbonInsertFunctions,
  'LET', 'LAMBDA', 'ISOMITTED', 'MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY',
  'IFS', 'SWITCH', 'XOR', 'IFNA', 'TRUE', 'FALSE',
  'ISBLANK', 'ISERR', 'ISERROR', 'ISLOGICAL', 'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISTEXT',
  'ISFORMULA', 'ISREF', 'ISEVEN', 'ISODD', 'ERROR.TYPE', 'N', 'NA', 'TYPE',
  'XLOOKUP', 'XMATCH', 'INDEX', 'MATCH', 'GETPIVOTDATA', 'HLOOKUP', 'LOOKUP', 'CHOOSE',
  'OFFSET', 'INDIRECT', 'ROW', 'ROWS', 'COLUMN', 'COLUMNS', 'ADDRESS', 'AREAS', 'FORMULATEXT', 'HYPERLINK',
  'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'VSTACK', 'HSTACK',
  'TOROW', 'TOCOL', 'WRAPROWS', 'WRAPCOLS', 'TAKE', 'DROP', 'CHOOSEROWS', 'CHOOSECOLS', 'EXPAND', 'TEXTSPLIT',
  // Wave B Math & Aggregates (+42)
  'SUMIFS', 'SUMPRODUCT', 'AGGREGATE', 'PRODUCT', 'INT', 'TRUNC', 'MOD', 'POWER', 'SQRT', 'SIGN', 'RAND', 'RANDBETWEEN',
  'CEILING', 'CEILING.MATH', 'FLOOR', 'FLOOR.MATH', 'MROUND', 'EXP', 'LN', 'LOG', 'LOG10', 'PI',
  'SIN', 'COS', 'TAN', 'ASIN', 'ACOS', 'ATAN', 'ATAN2', 'DEGREES', 'RADIANS',
  'FACT', 'COMBIN', 'PERMUT', 'GCD', 'LCM', 'QUOTIENT', 'EVEN', 'ODD', 'SUMSQ', 'MMULT', 'TRANSPOSE',
  // Wave B Text (+25)
  'CONCAT', 'CONCATENATE', 'TEXTJOIN', 'TEXT', 'VALUE', 'NUMBERVALUE', 'FIND', 'SEARCH',
  'SUBSTITUTE', 'REPLACE', 'REPT', 'PROPER', 'EXACT', 'CLEAN', 'CHAR', 'CODE', 'UNICODE', 'UNICHAR',
  'T', 'TEXTBEFORE', 'TEXTAFTER', 'VALUETOTEXT', 'ARRAYTOTEXT', 'FIXED', 'DOLLAR',
  // Wave B Date & Time (+20)
  'NOW', 'TIME', 'HOUR', 'MINUTE', 'SECOND', 'DATEVALUE', 'TIMEVALUE', 'EDATE', 'EOMONTH',
  'DATEDIF', 'DAYS', 'DAYS360', 'WEEKDAY', 'WEEKNUM', 'ISOWEEKNUM',
  'WORKDAY', 'WORKDAY.INTL', 'NETWORKDAYS', 'NETWORKDAYS.INTL', 'YEARFRAC',
  // Wave B Statistical (+27)
  'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS', 'MAXIFS', 'MINIFS', 'COUNTBLANK', 'MEDIAN', 'MODE', 'MODE.SNGL',
  'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'LARGE', 'SMALL', 'RANK', 'RANK.EQ', 'RANK.AVG',
  'PERCENTILE.INC', 'PERCENTILE.EXC', 'QUARTILE.INC', 'QUARTILE.EXC',
  'CORREL', 'COVARIANCE.S', 'SLOPE', 'INTERCEPT', 'FORECAST.LINEAR',
  // Wave B Financial (+14)
  'PMT', 'PV', 'FV', 'NPV', 'XNPV', 'IRR', 'XIRR', 'RATE', 'NPER', 'IPMT', 'PPMT', 'SLN', 'DB', 'DDB',
] as const;

export type NativeFunction = typeof nativeFunctions[number];

const allSpecs: readonly FunctionSpec[] = [
  ...logicalFunctions,
  ...mathFunctions,
  ...statisticalFunctions,
  ...textFunctions,
  ...datetimeFunctions,
  ...lookupFunctions,
  ...financialFunctions,
  ...dynamicArrayFunctions,
  ...infoFunctions,
  ...shapingFunctions,
];

export const functionRegistry: ReadonlyMap<string, FunctionSpec> = new Map(
  allSpecs.map(spec => [spec.name, spec]),
);

export const supportedFunctions: ReadonlySet<string> = new Set(nativeFunctions);

export function getFunctionSpec(name: string): FunctionSpec | undefined {
  return functionRegistry.get(name.toUpperCase());
}

export function isVolatileFunction(name: string): boolean {
  return getFunctionSpec(name)?.volatile === true;
}

export function functionSignature(name: string): FunctionSignature | undefined {
  return getFunctionSpec(name)?.signature;
}

export function evaluateFunction(name: string, args: readonly FormulaExpression[], ctx: FunctionContext) {
  const spec = supportedFunctions.has(name) ? functionRegistry.get(name) : undefined;
  return spec ? spec.evaluate(args, ctx) : error('#NAME?');
}
