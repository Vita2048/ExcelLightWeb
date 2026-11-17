import { expect, test } from '@playwright/test';

for (const path of ['/', '/ExcelLightWeb/']) {
  test(`production scaffold loads at ${path}`, async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', error => errors.push(error.message));
    page.on('response', response => { if (response.status() >= 400) errors.push(response.url()); });
    await page.emulateMedia({ colorScheme: 'light' });
    await page.goto(path);
    await expect(page.getByRole('heading', { name: 'Your spreadsheets. On your device.' })).toBeVisible();
    await expect(page.locator('html')).toHaveCSS('color-scheme', 'dark');
    await expect(page.getByRole('heading', { name: 'Application scaffold' })).toBeVisible();
    expect(errors).toEqual([]);
  });
}
