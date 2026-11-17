import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';

const fixture = (name: string) => readFileSync(new URL(`../fixtures/interchange/${name}`, import.meta.url), 'base64');

test('round-trips desktop fixtures through the built browser adapters', async ({ page }) => {
  await page.goto('/ExcelLightWeb/');
  const native = await page.evaluate(bytes => window.adapterProbe.native(bytes), fixture('DesktopStyles.exlt'));
  expect(native).toEqual([{ id: 'b71420c7-ad56-4000-8000-000000000001', name: 'Styles', freeze: [1, 0] }]);
  const xlsx = await page.evaluate(bytes => window.adapterProbe.xlsx(bytes), fixture('DesktopStyles.xlsx'));
  expect(xlsx.rotations).toContain(-45);
  expect(xlsx.rotations).toContain('stacked');
  expect(xlsx.gradients.map(gradient => gradient.geometry.type)).toEqual(['linear', 'path']);
  expect(xlsx.borders.some(border => border.up && border.down && border.pens.diagonal?.color === 'FF778899')).toBe(true);
  expect(await page.evaluate(() => window.adapterProbe.evaluate('=IFERROR(ROUND(-2.5,0),1/0)'))).toBe(-3);
});
