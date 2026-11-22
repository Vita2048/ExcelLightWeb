import { expect, test, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const appPath = '/ExcelLightWeb/';
const xlsxFixture = fileURLToPath(new URL('../fixtures/sales/showcase_data.xlsx', import.meta.url));

test.beforeEach(async ({ page }) => {
  page.on('dialog', dialog => { void dialog.accept(); });
});

test.afterEach(async ({ page }) => {
  if (!page.isClosed()) await page.close({ runBeforeUnload: false });
});

async function disableFileSystemPickers(page: Page) {
  await page.addInitScript(() => {
    Object.defineProperty(window, 'showOpenFilePicker', { configurable: true, value: undefined });
    Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: undefined });
  });
}

async function chooseXlsx(page: Page) {
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('dialog', { name: 'File Home' }).getByRole('button', { name: 'Open', exact: true }).click();
  await (await chooser).setFiles(xlsxFixture);
}

test('imports an XLSX workbook through File → Open after reviewing the loss report', async ({ page }) => {
  await disableFileSystemPickers(page);
  await page.goto(appPath);
  await expect(page.getByRole('grid')).toBeVisible();
  await page.getByRole('tab', { name: 'File' }).click();

  await chooseXlsx(page);
  const importReport = page.getByRole('dialog', { name: 'XLSX import report' });
  await expect(importReport).toBeVisible();
  await expect(importReport).toContainText('column-width');
  await expect(importReport).toContainText('default-dimensions');
  await importReport.getByRole('button', { name: 'Cancel' }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Book1');

  await chooseXlsx(page);
  await page.getByRole('dialog', { name: 'XLSX import report' }).getByRole('button', { name: 'Continue' }).click();
  await expect(page.locator('.workbook-title')).toHaveText('showcase_data');
  await expect(page.getByRole('gridcell', { name: 'A1, Customer Data' })).toHaveAttribute('aria-label', 'A1, Customer Data');
  await expect(page.getByRole('gridcell', { name: 'B3, Customer Name' })).toHaveAttribute('aria-label', 'B3, Customer Name');
  await expect(page.getByRole('gridcell', { name: 'B4, Acme Corporation' })).toHaveAttribute('aria-label', 'B4, Acme Corporation');
  const worksheets = page.getByRole('navigation', { name: 'Worksheets' });
  await expect(worksheets.getByRole('button', { name: 'Customers' })).toBeVisible();
  await expect(worksheets.getByRole('button', { name: 'Sales' })).toBeVisible();
});

test('exports an imported XLSX with a reviewed report without clearing its unsaved edits', async ({ page }) => {
  await disableFileSystemPickers(page);
  await page.goto(appPath);
  await expect(page.getByRole('grid')).toBeVisible();
  await page.getByRole('tab', { name: 'File' }).click();
  await chooseXlsx(page);
  await page.getByRole('dialog', { name: 'XLSX import report' }).getByRole('button', { name: 'Continue' }).click();
  await expect(page.locator('.workbook-title')).toHaveText('showcase_data');

  const nameBox = page.getByRole('textbox', { name: 'Name box' });
  await nameBox.fill('A1');
  await nameBox.press('Enter');
  await page.keyboard.type('Edited title');
  await page.keyboard.press('Enter');
  await expect(page.locator('.workbook-title')).toHaveText('showcase_data *');

  await page.getByRole('tab', { name: 'File' }).click();
  const download = page.waitForEvent('download');
  await page.getByRole('dialog', { name: 'File Home' }).getByRole('button', { name: 'Export', exact: true }).click();
  const exportReport = page.getByRole('dialog', { name: 'XLSX export report' });
  await expect(exportReport).toBeVisible();
  await expect(exportReport).toContainText('cache-freshness');
  await exportReport.getByRole('button', { name: 'Continue' }).click();
  expect((await download).suggestedFilename()).toMatch(/showcase_data\.xlsx$/i);
  await expect(page.locator('.workbook-title')).toHaveText('showcase_data *');
  await expect(page.locator('.status-bar .edit-status')).toContainText('XLSX export download started');
});
