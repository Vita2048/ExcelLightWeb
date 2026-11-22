import { expect, test, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const appPath = '/ExcelLightWeb/';
const salesFixture = fileURLToPath(new URL('../fixtures/sales/Showcase.exlt', import.meta.url));
const slaFixture = fileURLToPath(new URL('../fixtures/sla/Support-SLA.exlt', import.meta.url));

test.afterEach(async ({ page }) => {
  if (!page.isClosed()) await page.close({ runBeforeUnload: false });
});

async function openFixture(page: Page, fixture: string, title: string) {
  await page.getByRole('tab', { name: 'File' }).click();
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Open', exact: true }).click();
  await (await chooser).setFiles(fixture);
  await expect(page.locator('.workbook-title')).toHaveText(title);
}

async function disableFileSystemPickers(page: Page) {
  await page.addInitScript(() => {
    Object.defineProperty(window, 'showOpenFilePicker', { configurable: true, value: undefined });
    Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: undefined });
  });
}

async function enter(page: Page, address: string, value: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(address);
  await name.press('Enter');
  const formula = page.getByRole('textbox', { name: 'Formula bar' });
  await formula.fill(value);
  await formula.press('Enter');
}

async function runOnce(page: Page, anchor: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(anchor);
  await name.press('Enter');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await page.getByRole('dialog', { name: 'Run Python in this workbook?' })
    .getByRole('button', { name: 'Run once' }).click();
  const status = page.locator('.python-run-controls [role="status"]');
  await expect.poll(async () => await status.textContent(), { timeout: 420_000 })
    .toMatch(/(?:completed|failed|stopped|timed out)\./u);
  const message = await status.textContent();
  if (message !== 'Python run completed.') {
    const error = await page.locator('.python-error').textContent();
    throw new Error(`The fixture Python run did not complete: ${error ?? message ?? 'no error message'}`);
  }
}

async function cellLabel(page: Page, address: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(address);
  await name.press('Enter');
  const cell = page.locator(`[role="gridcell"][id$="-${address}"]`);
  await cell.waitFor({ state: 'attached', timeout: 15_000 });
  return await cell.getAttribute('aria-label');
}

test('runs the edited Sales showcase spill and refreshes its native KPI', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'S22 showcase chains execute the real Pyodide worker in Chromium.');
  test.setTimeout(480_000);
  const runtimeRequests: string[] = [];
  page.on('request', request => {
    if (request.url().includes('/runtime/pyodide/')) runtimeRequests.push(request.url());
  });

  await disableFileSystemPickers(page);
  await page.goto(appPath);
  await page.getByRole('grid').waitFor();
  await openFixture(page, salesFixture, 'Showcase');
  expect(runtimeRequests).toEqual([]);

  await page.getByRole('button', { name: 'Dashboard', exact: true }).click();
  const initialRevenue = await cellLabel(page, 'A4');
  await page.getByRole('button', { name: 'Sales', exact: true }).click();
  await enter(page, 'D4', '6');
  expect(runtimeRequests).toEqual([]);

  await page.getByRole('button', { name: 'Sales Analysis', exact: true }).click();
  await runOnce(page, 'A3');
  await expect(page.locator('.python-result-table')).toContainText('8999.94');
  await page.getByRole('button', { name: 'Dashboard', exact: true }).click();
  const updatedRevenue = await page.getByRole('gridcell', { name: /^A4,/u }).getAttribute('aria-label');
  expect(updatedRevenue).not.toBe(initialRevenue);
  expect(runtimeRequests.length).toBeGreaterThan(0);
});

test('runs the edited SLA showcase spill and refreshes its native breach KPI', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'S22 showcase chains execute the real Pyodide worker in Chromium.');
  test.setTimeout(180_000);
  const runtimeRequests: string[] = [];
  page.on('request', request => {
    if (request.url().includes('/runtime/pyodide/')) runtimeRequests.push(request.url());
  });

  await disableFileSystemPickers(page);
  await page.goto(appPath);
  await page.getByRole('grid').waitFor();
  await openFixture(page, slaFixture, 'Support-SLA');
  expect(runtimeRequests).toEqual([]);

  await page.getByRole('button', { name: 'Dashboard', exact: true }).click();
  await enter(page, 'P9', '2');
  expect(runtimeRequests).toEqual([]);
  const initialRate = await cellLabel(page, 'E6');
  await page.getByRole('button', { name: 'Analysis', exact: true }).click();
  await runOnce(page, 'B6');
  await expect(page.locator('.python-result-table')).toContainText('0.3207236842105263');
  await page.getByRole('button', { name: 'Dashboard', exact: true }).click();
  const updatedRate = await page.getByRole('gridcell', { name: /^E6,/u }).getAttribute('aria-label');
  expect(updatedRate).not.toBe(initialRate);
  expect(updatedRate).toMatch(/32\.1%/u);
  expect(runtimeRequests.length).toBeGreaterThan(0);
});
