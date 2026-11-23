import { expect, test, type Page } from '@playwright/test';

const appPath = '/ExcelLightWeb/';

async function enter(page: Page, address: string, value: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(address);
  await name.press('Enter');
  const formula = page.getByRole('textbox', { name: 'Formula bar' });
  await formula.fill(value);
  await formula.press('Enter');
}

for (const path of ['/', appPath]) {
  test(`service worker scope and CSP are correct at ${path}`, async ({ page }) => {
    const external: string[] = [];
    const cspViolations: string[] = [];
    const origin = new URL(test.info().project.use.baseURL as string).origin;
    page.on('request', request => {
      if (request.url().startsWith('http') && new URL(request.url()).origin !== origin) external.push(request.url());
    });
    page.on('console', message => {
      if (/content security policy|refused to (?:load|execute)/i.test(message.text())) cspViolations.push(message.text());
    });

    await page.goto(path);
    await expect(page.getByRole('grid')).toBeVisible();
    const policy = await page.locator('meta[http-equiv="Content-Security-Policy"]').getAttribute('content');
    expect(policy).toContain("'wasm-unsafe-eval'");
    const first = await page.evaluate(async () => {
      const registration = await navigator.serviceWorker.ready;
      return { scope: new URL(registration.scope).pathname, controlled: !!navigator.serviceWorker.controller };
    });
    expect(first.scope).toBe(path);

    await page.reload();
    await expect(page.getByRole('grid')).toBeVisible();
    await expect.poll(() => page.evaluate(() => !!navigator.serviceWorker.controller)).toBe(true);
    expect(external).toEqual([]);
    expect(cspViolations).toEqual([]);
  });

  test(`scoped shell cache reopens offline at ${path}`, async ({ page, browserName }) => {
    test.skip(browserName === 'webkit', 'Playwright WebKit offline service-worker navigation is unsupported (microsoft/playwright#42775).');
  const external: string[] = [];
  const cspViolations: string[] = [];
  const origin = new URL(test.info().project.use.baseURL as string).origin;
  page.on('request', request => {
    if (request.url().startsWith('http') && new URL(request.url()).origin !== origin) external.push(request.url());
  });
  page.on('console', message => {
    if (/content security policy|refused to (?:load|execute)/i.test(message.text())) cspViolations.push(message.text());
  });

  await page.goto(path);
  await expect(page.getByRole('grid')).toBeVisible();
  const policy = await page.locator('meta[http-equiv="Content-Security-Policy"]').getAttribute('content');
  expect(policy).toContain("'wasm-unsafe-eval'");
  const first = await page.evaluate(async () => {
    const registration = await navigator.serviceWorker.ready;
    return { scope: new URL(registration.scope).pathname, controlled: !!navigator.serviceWorker.controller };
  });
  expect(first.scope).toBe(path);

  await page.reload();
  await expect(page.getByRole('grid')).toBeVisible();
  await expect.poll(() => page.evaluate(() => !!navigator.serviceWorker.controller)).toBe(true);
  await page.context().setOffline(true);
  await page.reload();
  await expect(page.getByRole('grid')).toBeVisible();
  await expect(page.getByRole('button', { name: 'Offline' })).toBeVisible();

  await enter(page, 'B1', '<img src=x onerror=alert(1)>');
  await expect(page.locator('img[src="x"], img[onerror]')).toHaveCount(0);
  await enter(page, 'A1', '=PY');
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill('A1');
  await name.press('Enter');
  await expect(page.locator('.python-editor .monaco-editor')).toBeVisible();
  await expect(page.locator('.python-editor .view-lines')).toHaveText('');
  await page.context().setOffline(false);

  expect(external).toEqual([]);
  expect(cspViolations).toEqual([]);
  });
}

test('a checksum-complete runtime executes real Pyodide after reload with the browser offline', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'The offline Pyodide runtime journey uses Chromium, the production WASM worker, and real Cache Storage.');
  test.setTimeout(480_000);
  const origin = new URL(test.info().project.use.baseURL as string).origin;
  const external: string[] = [];
  page.on('request', request => {
    const url = request.url();
    if (url.startsWith('http') && new URL(url).origin !== origin) external.push(url);
  });

  await page.goto(appPath);
  await expect(page.getByRole('grid')).toBeVisible();
  await page.evaluate(async () => { await navigator.serviceWorker.ready; });
  await page.reload();
  await expect(page.getByRole('grid')).toBeVisible();
  await expect(page.getByRole('button', { name: 'Offline setup' })).toBeVisible();
  await page.getByRole('button', { name: 'Offline setup' }).click();
  await page.getByRole('button', { name: 'Prepare Python for offline use' }).click();
  await expect(page.getByRole('button', { name: 'Python ready offline' })).toBeVisible({ timeout: 240_000 });

  await page.context().setOffline(true);
  await page.reload();
  await expect(page.getByRole('grid')).toBeVisible();
  const offlineButton = page.getByRole('button', { name: 'Offline' });
  await expect(offlineButton).toBeVisible();
  await offlineButton.click();
  await expect(page.getByText('The app shell is cached. Offline Python is ready.')).toBeVisible();
  await enter(page, 'A1', '=PY');
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill('A1');
  await name.press('Enter');
  await expect(page.locator('.python-editor .monaco-editor')).toBeVisible();
  await expect(page.locator('.python-editor .view-lines')).toHaveText('');
  const sourceLine = page.locator('.python-editor .view-line').first();
  await sourceLine.click();
  await page.keyboard.insertText('42');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await page.getByRole('dialog', { name: 'Run Python in this workbook?' })
    .getByRole('button', { name: 'Trust this session and Run' }).click();
  await expect(page.getByRole('status').filter({ hasText: 'Python run completed.' })).toBeVisible({ timeout: 180_000 });
  await expect(page.locator('.python-scalar-result')).toHaveText('42');
  expect(external).toEqual([]);
});
