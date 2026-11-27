import type { FormulaExpression } from '../../core/formula-references';
import type { CellDocument } from '../../core/workbook';
import type { FormulaArray, FormulaLambda, FormulaValue, Range, Scalar } from '../formula-values';

export interface FunctionContext {
  readonly evaluate: (expression: FormulaExpression, includeRanges?: boolean) => FormulaValue;
  readonly scalar: (value: FormulaValue) => Scalar;
  readonly read: (range: Range, row: number, column: number) => Scalar;
  readonly values: (range: Range, subtotal?: { skipHidden: boolean }) => Iterable<Scalar>;
  readonly entries: (range: Range) => Iterable<{ row: number; column: number; value: Scalar }>;
  readonly number: (value: Scalar) => number;
  readonly text: (value: Exclude<Scalar, FormulaErrorLike>) => string;
  readonly locale: string;
  readonly today: number;
  readonly now: number;
  readonly target: { readonly sheetId: string; readonly row: number; readonly column: number };
  readonly cellAt: (sheetId: string, row: number, column: number) => CellDocument | undefined;
  readonly resolveRefText: (refText: string, a1?: boolean) => readonly Range[] | undefined;
  readonly random: () => number;
  readonly env: ReadonlyMap<string, FormulaValue>;
  readonly evaluateInEnv: (expression: FormulaExpression, env: ReadonlyMap<string, FormulaValue>, includeRanges?: boolean) => FormulaValue;
  readonly invokeLambda: (lambda: FormulaLambda, argValues: readonly FormulaValue[]) => FormulaValue;
  readonly arrayOf: (value: FormulaValue) => FormulaArray | undefined;
}

type FormulaErrorLike = Extract<Scalar, { kind: 'error' }>;

export type FunctionSignature = readonly [parameters: readonly string[], variadic?: true];

export type FunctionCategory =
  | 'logical'
  | 'math'
  | 'text'
  | 'datetime'
  | 'lookup'
  | 'statistical'
  | 'financial'
  | 'dynamic-array'
  | 'info';

export interface FunctionSpec {
  readonly name: string;
  readonly category: FunctionCategory;
  readonly minArgs: number;
  readonly maxArgs: number;
  readonly volatile?: boolean;
  readonly signature: FunctionSignature;
  readonly evaluate: (args: readonly FormulaExpression[], ctx: FunctionContext) => FormulaValue;
}
