import { expect, test, type Page } from '@playwright/test';

const appPath = '/ExcelLightWeb/';

test.afterEach(async ({ page }) => {
  if (!page.isClosed()) await page.close({ runBeforeUnload: false });
});

async function go(page: Page, address: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(address);
  await name.press('Enter');
}

async function enter(page: Page, address: string, value: string) {
  await go(page, address);
  const editor = page.getByRole('textbox', { name: 'Formula bar' });
  await editor.fill(value);
  await editor.press('Enter');
}

async function replacePythonSource(page: Page, source: string) {
  await page.locator('.python-editor .view-line').first().click();
  await page.keyboard.press('Control+A');
  await page.keyboard.insertText(source);
}

test('runs a mixed dependency chain atomically and schedules only after local Automatic consent', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'The mixed Pyodide runtime journey runs in Chromium.');
  test.setTimeout(900_000);
  await page.goto(appPath);
  await page.getByRole('grid').waitFor();
  await enter(page, 'A1', '4');
  await enter(page, 'B1', '=A1*2');
  await enter(page, 'C1', '=PY');
  await go(page, 'C1');
  await expect(page.locator('.python-panel')).toBeVisible();
  await replacePythonSource(page, 'xl("B1") + 1');
  await enter(page, 'D1', '=C1*2');
  await go(page, 'C1');

  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await page.getByRole('dialog', { name: 'Run Python in this workbook?' })
    .getByRole('button', { name: 'Run once' }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run completed.', { timeout: 420_000 });
  await expect(page.locator('.python-scalar-result')).toHaveText('9');
  await expect(page.getByRole('gridcell', { name: 'D1, 18', exact: true })).toBeAttached();

  // One Undo removes the mixed Python/formula result generation but keeps authored source.
  await page.getByRole('group', { name: 'Quick Access' }).getByRole('button', { name: 'Undo' }).click();
  await expect(page.locator('.python-empty-result')).toHaveText('Run Python to see its result.');
  await expect(page.locator('.python-editor .view-lines')).toHaveText('xl("B1") + 1');
  await page.getByRole('group', { name: 'Quick Access' }).getByRole('button', { name: 'Redo' }).click();
  await expect(page.locator('.python-scalar-result')).toHaveText('9');

  // The browser-local Automatic choice is consent; the next edit runs the mixed chain.
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Options', exact: true }).click();
  const options = page.getByRole('dialog', { name: 'Excel-Lite Options' });
  await options.locator('input[name="python-calculation"][value="automatic"]').check();
  await options.getByRole('button', { name: 'OK' }).click();
  await page.getByRole('button', { name: /Back/u }).click();
  await enter(page, 'A1', '5');
  await go(page, 'C1');
  await expect(page.locator('.python-scalar-result')).toHaveText('11', { timeout: 120_000 });
  await expect(page.getByRole('gridcell', { name: 'D1, 22', exact: true })).toBeAttached();

  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Options', exact: true }).click();
  const manualOptions = page.getByRole('dialog', { name: 'Excel-Lite Options' });
  await manualOptions.locator('input[name="python-calculation"][value="manual"]').check();
  await manualOptions.getByRole('button', { name: 'OK' }).click();
  await page.getByRole('button', { name: /Back/u }).click();
  await enter(page, 'A1', '6');
  await go(page, 'C1');
  await expect(page.getByRole('gridcell', { name: 'B1, 12', exact: true })).toBeAttached();
  await expect(page.locator('.python-result-header')).toContainText('Previous successful result · stale');
  await page.getByRole('group', { name: 'Quick Access' }).getByRole('button', { name: 'Save', exact: true }).click();
  const save = page.getByRole('dialog', { name: 'Save to this browser' });
  await save.getByRole('textbox', { name: 'Workbook name' }).fill('Mixed Automatic Stale');
  await save.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.locator('.status-bar [role="status"]'))
    .toContainText('Saved to this browser as Mixed Automatic Stale.');

  await page.getByRole('tab', { name: 'File' }).click();
  const newWorkbook = page.getByRole('button', { name: 'New', exact: true });
  await expect(newWorkbook).toBeEnabled();
  await newWorkbook.click();
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Options', exact: true }).click();
  const automaticOptions = page.getByRole('dialog', { name: 'Excel-Lite Options' });
  await automaticOptions.locator('input[name="python-calculation"][value="automatic"]').check();
  await automaticOptions.getByRole('button', { name: 'OK' }).click();
  await page.getByRole('button', { name: 'Open recent Mixed Automatic Stale' }).click();
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Opened Mixed Automatic Stale.');
  await expect(page.locator('.workbook-title')).toHaveText('Mixed Automatic Stale');
  await go(page, 'C1');
  await expect(page.getByRole('grid')).toHaveAttribute('aria-activedescendant', /-C1$/u);
  await expect(page.locator('.python-scalar-result')).toHaveText('13', { timeout: 120_000 });
  await expect(page.getByRole('gridcell', { name: 'D1, 26', exact: true })).toBeAttached();
  await expect(page.getByRole('dialog', { name: 'Run Python in this workbook?' })).toHaveCount(0);
});

