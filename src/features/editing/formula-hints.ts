import { functionSignature } from '../../calculation/formula-functions';
import { scanFormula } from '../formulas/formula-presentation';

export function formulaHint(source: string, caret: number, separator = ',') {
  const frames = scanFormula(source, Math.min(caret, source.length), ',' + separator);
  for (const { name, argument } of frames.reverse()) {
    const signature = functionSignature(name);
    if (signature) {
      const [parameters, variadic = false] = signature;
      return { name, parameters, variadic, active: Math.min(argument, parameters.length - 1) };
    }
  }
}
