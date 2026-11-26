import { expect, test, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import { observePythonRuntime, readPythonRuntimeTrace, waitForPythonRuntime } from './python-runtime-observer';

const appPath = '/ExcelLightWeb/';
const salesFixture = fileURLToPath(new URL('../../docs/showcase/Showcase.exlt', import.meta.url));
const slaFixture = fileURLToPath(new URL('../../docs/showcase/Support-SLA.exlt', import.meta.url));

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
    .toMatch(/(?:completed|unavailable|failed|stopped|timed out)\./u);
  const message = await status.textContent();
  if (message !== 'Python run completed.') {
    const error = await page.locator('.python-error').textContent();
    throw new Error(`The fixture Python run did not complete: ${error ?? message ?? 'no error message'}`);
  }
}

async function waitForAutomaticPython(page: Page, sheet: string, anchor: string) {
  await page.getByRole('button', { name: sheet, exact: true }).click();
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(anchor);
  await name.press('Enter');
  await expect(page.locator('.python-run-controls [role="status"]'))
    .toHaveText('Automatic · Python results are current.', { timeout: 420_000 });
}

async function cellLabel(page: Page, address: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(address);
  await name.press('Enter');
  const cell = page.locator(`[role="gridcell"][id$="-${address}"]`);
  await cell.waitFor({ state: 'attached', timeout: 15_000 });
  return await cell.getAttribute('aria-label');
}

async function placeOnDashboard(page: Page, sourceLabel: string) {
  await expect(page.locator('.python-figure')).toBeVisible();
  await page.getByRole('button', { name: 'Place on Sheet', exact: true }).click();
  await page.getByRole('button', { name: 'Dashboard', exact: true }).click();
  const capture = page.getByLabel('Place Python Figure on worksheet');
  const bounds = await capture.boundingBox();
  if (!bounds) throw new Error('The Dashboard Figure placement preview is not available.');
  await page.mouse.move(bounds.x + 60, bounds.y + 45);
  await page.mouse.down();
  await page.mouse.move(bounds.x + 310, bounds.y + 205, { steps: 5 });
  await page.mouse.up();
  const figure = page.getByRole('group', { name: `Figure from ${sourceLabel}` }).last();
  await expect(figure).toBeVisible();
  await expect(figure.getByRole('button', { name: 'Resize figure bottom right' })).toBeVisible();
  await expect(figure.getByRole('button', { name: 'Resize figure top left', exact: true })).toHaveCSS('cursor', 'nwse-resize');
  await expect(figure.getByRole('button', { name: 'Resize figure top', exact: true })).toHaveCSS('cursor', 'ns-resize');
  await expect(figure.getByRole('button', { name: 'Resize figure right', exact: true })).toHaveCSS('cursor', 'ew-resize');
  await expect(figure.getByRole('button', { name: 'Resize figure bottom right', exact: true })).toHaveCSS('cursor', 'nwse-resize');
  await expect(figure.getByRole('button', { name: 'Resize figure bottom', exact: true })).toHaveCSS('cursor', 'ns-resize');
  await expect(figure.getByRole('button', { name: 'Resize figure bottom left', exact: true })).toHaveCSS('cursor', 'nesw-resize');
  await expect(figure.getByRole('button', { name: 'Resize figure left', exact: true })).toHaveCSS('cursor', 'ew-resize');
  await expect(figure.getByRole('button', { name: 'Resize figure top right', exact: true })).toHaveCSS('cursor', 'nesw-resize');
  const initialWidth = await figure.locator('img').evaluate(image => (image as HTMLImageElement).naturalWidth);
  const handle = figure.getByRole('button', { name: 'Resize figure bottom right' });
  const handleBounds = await handle.boundingBox();
  if (!handleBounds) throw new Error('The Dashboard Figure resize handle is not visible.');
  await page.mouse.move(handleBounds.x + handleBounds.width / 2, handleBounds.y + handleBounds.height / 2);
  await page.mouse.down();
  await page.mouse.move(handleBounds.x + handleBounds.width / 2 + 25,
    handleBounds.y + handleBounds.height / 2 + 20, { steps: 3 });
  await page.mouse.up();
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Figure placement updated.');
  await expect.poll(() => figure.locator('img').evaluate(image => (image as HTMLImageElement).naturalWidth))
    .toBeGreaterThan(initialWidth);
}

test('runs the edited Sales showcase spill and refreshes its native KPI', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'S22 showcase chains execute the real Pyodide worker in Chromium.');
  test.setTimeout(480_000);
  await observePythonRuntime(page);

  await disableFileSystemPickers(page);
  await page.goto(appPath);
  await page.getByRole('grid').waitFor();
  await waitForPythonRuntime(page);
  await openFixture(page, salesFixture, 'Showcase');
  await waitForAutomaticPython(page, 'Sales Analysis', 'A3');
  const initialRuns = (await readPythonRuntimeTrace(page)).runRequests;
  expect(initialRuns).toBeGreaterThan(0);

  await page.getByRole('button', { name: 'Dashboard', exact: true }).click();
  const initialRevenue = await cellLabel(page, 'A4');
  await page.getByRole('button', { name: 'Sales', exact: true }).click();
  await enter(page, 'D4', '6');
  await expect.poll(async () => (await readPythonRuntimeTrace(page)).runRequests, { timeout: 120_000 })
    .toBeGreaterThan(initialRuns);
  await waitForAutomaticPython(page, 'Sales Analysis', 'A3');

  await page.getByRole('button', { name: 'Sales Analysis', exact: true }).click();
  await runOnce(page, 'A3');
  await expect(page.locator('.python-result-table')).toContainText('8999.94');
  await page.getByRole('button', { name: 'Dashboard', exact: true }).click();
  const updatedRevenue = await page.getByRole('gridcell', { name: /^A4,/u }).getAttribute('aria-label');
  expect(updatedRevenue).not.toBe(initialRevenue);

  await page.getByRole('button', { name: 'Charts', exact: true }).click();
  await runOnce(page, 'A3');
  await expect(page.locator('.python-result-table')).toBeVisible();
  await runOnce(page, 'E22');
  expect(await cellLabel(page, 'E22')).toContain('[plot]');
  await placeOnDashboard(page, 'Charts!E22');
  expect((await readPythonRuntimeTrace(page)).runRequests).toBeGreaterThan(initialRuns);
});

test('both showcases calculate on open and keep edits across a browser save and reload', async ({ page }) => {
  test.setTimeout(600_000);
  await observePythonRuntime(page);
  await disableFileSystemPickers(page);

  const showcases = [
    [salesFixture, 'Showcase', 'Dashboard', 'A4'],
    [slaFixture, 'Support-SLA', 'Dashboard', 'E6'],
  ] as const;
  const savedKpis = new Map<string, string | null>();

  await page.goto(appPath);
  await page.getByRole('grid').waitFor();
  await waitForPythonRuntime(page);
  for (const [fixture, title, sheet, kpi] of showcases) {
    await openFixture(page, fixture, title);
    await waitForAutomaticPython(page, title === 'Showcase' ? 'Sales Analysis' : 'Analysis',
      title === 'Showcase' ? 'A3' : 'B6');
    await page.getByRole('button', { name: sheet, exact: true }).click();
    const savedKpi = await cellLabel(page, kpi);
    savedKpis.set(title, savedKpi);
    await enter(page, 'Z60', 'Reviewed in the browser');
    await expect(page.getByRole('gridcell', { name: 'Z60, Reviewed in the browser' })).toHaveAttribute('aria-label',
      'Z60, Reviewed in the browser');

    await page.getByRole('tab', { name: 'File' }).click();
    await page.getByRole('dialog', { name: 'File Home' }).getByRole('button', { name: 'Save', exact: true }).click();
    await expect(page.locator('.status-bar [role="status"]')).toContainText('Saved to this browser');
    const stored = await page.evaluate(async name => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open('excellight-workspace', 2);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const transaction = db.transaction('workspaces', 'readonly');
      const loaded = await new Promise<boolean>((resolve, reject) => {
        const request = transaction.objectStore('workspaces').getAll();
        request.onsuccess = () => resolve((request.result as { name?: string }[]).some(record => record.name === name));
        request.onerror = () => reject(request.error);
      });
      db.close();
      return loaded;
    }, title);
    expect(stored).toBe(true);
  }

  await page.reload();
  await page.getByRole('grid').waitFor();
  await waitForPythonRuntime(page);
  let reopened = 0;
  for (const [, title, sheet, kpi] of showcases) {
    if (reopened > 0) {
      await page.getByRole('group', { name: 'Quick Access' }).getByRole('button', { name: 'Save', exact: true }).click();
      await expect(page.locator('.status-bar [role="status"]')).toContainText('Saved to this browser');
    }
    await page.getByRole('tab', { name: 'File' }).click();
    const runsBeforeOpen = (await readPythonRuntimeTrace(page)).runRequests;
    await page.getByRole('button', { name: 'Open recent ' + title, exact: true }).click();
    await expect(page.locator('.workbook-title')).toHaveText(title);
    await waitForAutomaticPython(page, title === 'Showcase' ? 'Sales Analysis' : 'Analysis',
      title === 'Showcase' ? 'A3' : 'B6');
    await page.getByRole('button', { name: sheet, exact: true }).click();
    // Navigate to the edited cell so the bounded accessible projection mounts it.
    expect(await cellLabel(page, 'Z60')).toBe('Z60, Reviewed in the browser');
    // The saved workbook is reopened in Automatic mode and its Python code recalculates.
    expect(await cellLabel(page, kpi)).toBe(savedKpis.get(title));
    await expect.poll(async () => (await readPythonRuntimeTrace(page)).runRequests, { timeout: 120_000 })
      .toBeGreaterThan(runsBeforeOpen);
    reopened++;
  }
});

test('runs the edited SLA showcase spill and refreshes its native breach KPI', async ({ page, browserName }) => {
  test.skip(browserName !== 'chromium', 'S22 showcase chains execute the real Pyodide worker in Chromium.');
  test.setTimeout(480_000);
  await observePythonRuntime(page);

  await disableFileSystemPickers(page);
  await page.goto(appPath);
  await page.getByRole('grid').waitFor();
  const runtime = await waitForPythonRuntime(page);
  expect(runtime.progressMessages.some(message => message.includes('Loading NumPy (1 of 3)'))).toBe(true);
  expect(runtime.progressMessages.some(message => message.includes('Loading pandas (2 of 3)'))).toBe(true);
  expect(runtime.progressMessages.some(message => message.includes('Loading Matplotlib (3 of 3)'))).toBe(true);
  await openFixture(page, slaFixture, 'Support-SLA');
  await waitForAutomaticPython(page, 'Analysis', 'B6');

  await page.getByRole('button', { name: 'Dashboard', exact: true }).click();
  const runsBeforeEdit = (await readPythonRuntimeTrace(page)).runRequests;
  await enter(page, 'P9', '2');
  await expect.poll(async () => (await readPythonRuntimeTrace(page)).runRequests, { timeout: 120_000 })
    .toBeGreaterThan(runsBeforeEdit);
  await waitForAutomaticPython(page, 'Analysis', 'B6');
  const initialRate = await cellLabel(page, 'E6');
  await page.getByRole('button', { name: 'Analysis', exact: true }).click();
  await page.evaluate(() => {
    const probe = { messages: [] as string[], sawCalculatingCell: false };
    (window as Window & { calculationProgressProbe?: typeof probe }).calculationProgressProbe = probe;
    const capture = () => {
      const status = document.querySelector('.status-bar [role="status"]');
      if (status) probe.messages.push(status.textContent ?? '');
      const cell = document.querySelector('[role="gridcell"][id$="-B6"]');
      if (cell?.getAttribute('aria-description') === 'Calculating') probe.sawCalculatingCell = true;
    };
    const observer = new MutationObserver(capture);
    observer.observe(document.documentElement, { childList: true, subtree: true, attributes: true,
      attributeFilter: ['aria-description'], characterData: true });
    (window as Window & { calculationProgressObserver?: MutationObserver }).calculationProgressObserver = observer;
    capture();
  });
  await runOnce(page, 'B6');
  const progress = await page.evaluate(() => (window as Window & { calculationProgressProbe?: {
    messages: string[]; sawCalculatingCell: boolean;
  } }).calculationProgressProbe);
  expect(progress?.messages.some(message => message.includes('Calculating Analysis!B6'))).toBe(true);
  expect(progress?.sawCalculatingCell).toBe(true);
  await page.evaluate(() => (window as Window & { calculationProgressObserver?: MutationObserver })
    .calculationProgressObserver?.disconnect());
  await expect(page.locator('.python-result-table')).toContainText('0.3207236842105263');
  await page.getByRole('button', { name: 'Dashboard', exact: true }).click();
  const updatedRate = await page.getByRole('gridcell', { name: /^E6,/u }).getAttribute('aria-label');
  expect(updatedRate).not.toBe(initialRate);
  expect(updatedRate).toMatch(/32\.1%/u);

  await page.getByRole('button', { name: 'Analysis', exact: true }).click();
  await runOnce(page, 'B52');
  expect(await cellLabel(page, 'B52')).toContain('[plot]');
  await placeOnDashboard(page, 'Analysis!B52');
  expect((await readPythonRuntimeTrace(page)).runRequests).toBeGreaterThan(runsBeforeEdit);
});
