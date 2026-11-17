import { expect, test } from '@playwright/test';

for (const path of ['/', '/ExcelLightWeb/']) {
  test(`production worksheet loads at ${path} with dark top bars under OS light mode`, async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('response', response => { if (response.status() >= 400) errors.push(response.url()); });
    await page.emulateMedia({ colorScheme: 'light' });
    await page.goto(path);
    await expect(page.getByRole('grid', { name: 'Sheet1' })).toHaveAttribute('aria-rowcount', '1000');
    await expect(page.getByRole('grid')).toHaveAttribute('aria-colcount', '26');
    await expect(page.locator('html')).toHaveCSS('color-scheme', 'dark');
    for (const selector of ['.title-bar', '.ribbon', '.formula-bar']) {
      await expect(page.locator(selector)).toHaveCSS('background-color', 'rgb(0, 0, 0)');
    }
    await expect(page.getByRole('tab', { name: 'Home' })).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveAttribute('readonly', '');
    await expect(page.getByRole('button', { name: 'Undo' })).toBeDisabled();
    expect(errors).toEqual([]);
  });
}
