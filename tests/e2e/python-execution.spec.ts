import { expect, test, type Page } from '@playwright/test';

const path = '/ExcelLightWeb/';

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

async function replacePythonSource(page: Page, text: string) {
  const firstLine = page.locator('.python-editor .view-line').first();
  await firstLine.click();
  await page.keyboard.press('Control+A');
  await page.keyboard.insertText(text);
}

test('runs local Python only on consent, rejects occupied spills and inserts a successful table as values', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'The full real-Pyodide journey runs in Chromium via npm run test:python-execution.');
  test.setTimeout(480_000);
  const origin = new URL(test.info().project.use.baseURL as string).origin;
  const runtimeRequests: string[] = [];
  const externalRequests: string[] = [];
  page.on('request', request => {
    const url = request.url();
    if (url.includes('/runtime/pyodide/')) runtimeRequests.push(url);
    if (!url.startsWith(origin) && !url.startsWith('blob:') && !url.startsWith('data:')) externalRequests.push(url);
  });

  await page.goto(path);
  await page.getByRole('grid').waitFor();
  await enter(page, 'B1', '5');
  await enter(page, 'B2', '8');
  await enter(page, 'D2', '99');
  await enter(page, 'D1', '=PY');
  await go(page, 'D1');
  await expect(page.locator('.python-panel')).toBeVisible();
  await expect(page.getByRole('button', { name: 'Run', exact: true })).toBeEnabled();
  await expect.poll(() => runtimeRequests).toHaveLength(0);

  await page.getByRole('group', { name: 'Quick Access' }).getByRole('button', { name: 'Save', exact: true }).click();
  const nameDialog = page.getByRole('dialog', { name: 'Save to this browser' });
  await expect(nameDialog).toBeVisible();
  await nameDialog.getByRole('textbox', { name: 'Workbook name' }).fill('Python execution');
  await nameDialog.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Saved to this browser as Python execution.');
  await page.reload();
  await page.getByRole('grid').waitFor();
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Open recent Python execution' }).click();
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Opened Python execution.');
  await go(page, 'D1');
  await expect(page.locator('.python-editor .view-lines')).toHaveText('42');
  expect(runtimeRequests).toHaveLength(0);

  await page.getByRole('button', { name: 'Run', exact: true }).click();
  const prompt = page.getByRole('dialog', { name: 'Run Python in this workbook?' });
  await expect(prompt).toBeVisible();
  await prompt.getByRole('button', { name: 'Cancel' }).click();
  expect(runtimeRequests).toHaveLength(0);

  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await page.getByRole('dialog', { name: 'Run Python in this workbook?' })
    .getByRole('button', { name: 'Trust this session and Run' }).click();
  await expect(page.getByRole('status').filter({ hasText: 'Python run completed.' })).toBeVisible({ timeout: 420_000 });
  await expect(page.locator('.python-scalar-result')).toHaveText('42');
  expect(runtimeRequests.length).toBeGreaterThan(0);

  await replacePythonSource(page, 'xl("B1:B2")');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.getByRole('status').filter({ hasText: 'Python run failed.' })).toBeVisible({ timeout: 120_000 });
  await expect(page.locator('.python-error')).toContainText('#SPILL!');
  await expect(page.locator('.python-scalar-result')).toHaveText('42');

  await enter(page, 'D2', '');
  await go(page, 'D1');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  const result = page.getByRole('region', { name: 'Python result' });
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run completed.', { timeout: 120_000 });
  await expect(result.getByRole('table')).toContainText('5');
  await expect(result.getByRole('table')).toContainText('8');

  await enter(page, 'B1', '9');
  await go(page, 'D1');
  await expect(result.getByText('Previous successful result · stale')).toBeVisible();
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run completed.', { timeout: 120_000 });
  await expect(result.getByRole('table')).toContainText('9');

  await replacePythonSource(page, 'import matplotlib.pyplot as plt\nplt.figure(figsize=(1, 1))\nplt.plot([1, 2], [3, 4])\nplt.gcf()');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run completed.', { timeout: 120_000 });
  await expect(result.locator('.python-figure')).toHaveAttribute('width', '100');
  await expect(result.locator('.python-figure')).toHaveAttribute('height', '100');

  await replacePythonSource(page, '__import__("pandas").Series([1, 2], name="Count")');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run completed.', { timeout: 120_000 });
  await expect(result.getByRole('table')).toContainText('Count');
  await expect(result.getByRole('table')).toContainText('2');

  await replacePythonSource(page, '__import__("numpy").array([[1, 2], [3, 4]])');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run completed.', { timeout: 120_000 });
  await expect(result.getByRole('table')).toContainText('4');

  await replacePythonSource(page, 'float("inf")');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run failed.', { timeout: 120_000 });
  await expect(result.locator('.python-error')).toContainText('Unsupported result value');
  await expect(result.getByRole('table')).toContainText('4');

  await replacePythonSource(page, 'xl("B1:B2")');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run completed.', { timeout: 120_000 });
  await expect(result.getByRole('table')).toContainText('8');
  await expect(page.getByRole('gridcell', { name: 'E1, blank', exact: true })).toBeAttached();

  await replacePythonSource(page, 'while True: pass');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Stop', exact: true })).toBeEnabled();
  await page.getByRole('button', { name: 'Stop', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run stopped.');
  await expect(result.getByRole('table')).toContainText('8');

  await replacePythonSource(page, 'while True: pass');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run timed out.', { timeout: 60_000 });
  await expect(result.getByRole('table')).toContainText('8');

  await replacePythonSource(page, 'xl("B1:B2")');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run completed.', { timeout: 420_000 });
  await expect(result.getByRole('table')).toContainText('8');

  await result.getByRole('button', { name: 'Insert as Values' }).click();
  await expect(page.locator('.python-panel')).toHaveCount(0);
  await expect(page.getByRole('gridcell', { name: 'D1, 0', exact: true })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'D2, 9', exact: true })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'D3, 8', exact: true })).toBeAttached();
  expect(externalRequests).toEqual([]);
});
