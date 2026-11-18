import { expect, test } from '@playwright/test';

test('Chromium native clipboard and IME reach the production editing surface', async ({ page, context, browserName }) => {
  test.skip(browserName !== 'chromium', 'Playwright clipboard permissions and CDP IME are verified through Chromium.');
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await page.goto('/ExcelLightWeb/');
  await page.evaluate(() => navigator.clipboard.writeText('7\t= A1*3'));
  const grid = page.getByRole('grid');
  await grid.focus();
  await page.keyboard.press('Control+v');
  await expect(page.getByRole('gridcell', { name: 'A1, 7' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B1, 21' })).toBeAttached();
  await page.keyboard.press('Control+c');
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('7');

  await grid.press('ArrowDown');
  await grid.press('F2');
  const editor = page.getByRole('textbox', { name: 'Edit A2' });
  const cdp = await context.newCDPSession(page);
  await cdp.send('Input.imeSetComposition', { text: 'かな', selectionStart: 2, selectionEnd: 2 });
  await expect(editor).toHaveValue('かな');
  await expect(page.getByRole('gridcell', { name: 'A2, blank' })).toBeAttached();
  await cdp.send('Input.insertText', { text: '仮名' });
  await editor.press('Enter');
  await expect(page.getByRole('gridcell', { name: 'A2, 仮名' })).toBeAttached();
});
