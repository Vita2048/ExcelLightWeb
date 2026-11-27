import { error } from '../formula-values';
import type { FormulaExpression } from '../../core/formula-references';
import { datetimeFunctions } from './datetime';
import { dynamicArrayFunctions } from './dynamic-array';
import { financialFunctions } from './financial';
import { infoFunctions } from './info';
import { logicalFunctions } from './logical';
import { lookupFunctions } from './lookup';
import { mathFunctions } from './math';
import { statisticalFunctions } from './statistical';
import { textFunctions } from './text';
import type { FunctionContext, FunctionSignature, FunctionSpec } from './types';

export const nativeFunctions = [
  'SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'IF', 'AND', 'OR', 'NOT', 'IFERROR',
  'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'SUMIF', 'COUNTIF', 'LEFT', 'RIGHT', 'MID',
  'LEN', 'TRIM', 'UPPER', 'LOWER', 'DATE', 'YEAR', 'MONTH', 'DAY', 'TODAY', 'VLOOKUP', 'SUBTOTAL',
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
