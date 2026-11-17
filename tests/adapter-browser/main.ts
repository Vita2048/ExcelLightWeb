import { evaluateProbe } from '../../src/calculation/formula-probe';
import { openNativeProbe } from '../../src/persistence/native-probe';
import { rewriteXlsxProbe } from '../../src/persistence/xlsx-style-probe';

const adapterProbe = {
  async native(base64: string) {
    const blob = new Blob([Uint8Array.from(atob(base64), char => char.charCodeAt(0))]);
    const document = await openNativeProbe(blob);
    const saved = await openNativeProbe(await document.save());
    return saved.sheets.map(({ id, name, freeze }) => ({ id, name, freeze }));
  },
  async xlsx(base64: string) {
    const blob = new Blob([Uint8Array.from(atob(base64), char => char.charCodeAt(0))]);
    const written = await rewriteXlsxProbe(blob);
    const reread = await rewriteXlsxProbe(written.blob);
    return { rotations: reread.rotations, gradients: reread.gradients, borders: reread.borders };
  },
  evaluate: evaluateProbe,
};

declare global {
  interface Window { adapterProbe: typeof adapterProbe }
}
window.adapterProbe = adapterProbe;
