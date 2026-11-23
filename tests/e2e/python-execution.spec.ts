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
  test.setTimeout(900_000);
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
  await expect(page.locator('.python-editor .view-lines')).toHaveText('');
  expect(runtimeRequests).toHaveLength(0);
  await replacePythonSource(page, '42');

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

  const imageDownload = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save Image', exact: true }).click();
  expect((await imageDownload).suggestedFilename()).toBe('Python Figure.png');

  await page.getByRole('button', { name: 'Place on Sheet', exact: true }).click();
  const placementCapture = page.getByLabel('Place Python Figure on worksheet');
  const captureBounds = await placementCapture.boundingBox();
  if (!captureBounds) throw new Error('Figure placement capture is not visible.');
  await page.mouse.move(captureBounds.x + 40, captureBounds.y + 35);
  await page.mouse.down();
  await page.mouse.move(captureBounds.x + 300, captureBounds.y + 185, { steps: 5 });
  await page.mouse.up();
  const firstPlacement = page.getByRole('group', { name: 'Figure from Sheet1!D1' });
  await expect(firstPlacement).toBeVisible();
  await expect(firstPlacement.getByRole('button', { name: 'Resize figure bottom right' })).toBeVisible();
  await expect(page.getByRole('gridcell', { name: 'B1, 9', exact: true })).toBeAttached();
  const initialRaster = await firstPlacement.locator('img').evaluate(image => ({ width: (image as HTMLImageElement).naturalWidth,
    height: (image as HTMLImageElement).naturalHeight }));
  const initialPlacementPng = await firstPlacement.locator('img').getAttribute('src');

  const resizeHandle = firstPlacement.getByRole('button', { name: 'Resize figure bottom right' });
  const handleBounds = await resizeHandle.boundingBox();
  if (!handleBounds) throw new Error('Figure resize handle is not visible.');
  await page.mouse.move(handleBounds.x + handleBounds.width / 2, handleBounds.y + handleBounds.height / 2);
  await page.mouse.down();
  await page.mouse.move(handleBounds.x + handleBounds.width / 2 + 45, handleBounds.y + handleBounds.height / 2 + 30, { steps: 4 });
  await page.mouse.up();
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Figure placement updated.');
  await expect.poll(() => firstPlacement.locator('img').evaluate(image => (image as HTMLImageElement).naturalWidth))
    .toBeGreaterThan(initialRaster.width);

  const resizedBounds = await firstPlacement.getByRole('button', { name: 'Resize figure bottom right' }).boundingBox();
  if (!resizedBounds) throw new Error('Figure resize handle disappeared after redraw.');
  await page.mouse.move(resizedBounds.x + resizedBounds.width / 2, resizedBounds.y + resizedBounds.height / 2);
  await page.mouse.down();
  await page.mouse.move(resizedBounds.x + resizedBounds.width / 2 - 45,
    resizedBounds.y + resizedBounds.height / 2 - 30, { steps: 4 });
  await page.mouse.up();
  await expect.poll(() => firstPlacement.locator('img').getAttribute('src'), { timeout: 120_000 }).toBe(initialPlacementPng);

  await page.getByRole('button', { name: 'Place on Sheet', exact: true }).click();
  const secondCaptureBounds = await placementCapture.boundingBox();
  if (!secondCaptureBounds) throw new Error('Figure placement capture is not visible for a second placement.');
  await page.mouse.move(secondCaptureBounds.x + 330, secondCaptureBounds.y + 35);
  await page.mouse.down();
  await page.mouse.move(secondCaptureBounds.x + 515, secondCaptureBounds.y + 150, { steps: 4 });
  await page.mouse.up();
  const placements = page.getByRole('group', { name: 'Figure from Sheet1!D1' });
  await expect(placements).toHaveCount(2);
  const rasterSizes = await placements.locator('img').evaluateAll(images => images.map(image => [
    (image as HTMLImageElement).naturalWidth, (image as HTMLImageElement).naturalHeight]));
  expect(new Set(rasterSizes.map(size => `${size[0]}x${size[1]}`)).size).toBe(2);
  await placements.last().getByRole('button', { name: 'Delete figure placement' }).click();
  await expect(placements).toHaveCount(1);

  await replacePythonSource(page, 'while True: pass');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Stop', exact: true })).toBeEnabled();
  await page.getByRole('button', { name: 'Stop', exact: true }).click();
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run stopped.');
  const retainedPlacement = page.getByRole('group', { name: 'Figure from Sheet1!D1' });
  await replacePythonSource(page, 'import matplotlib.pyplot as plt\nplt.figure(figsize=(1, 1))\nplt.plot([1, 2], [3, 4])\nplt.gcf()');
  await retainedPlacement.click();
  const beforeMove = await retainedPlacement.boundingBox();
  if (!beforeMove) throw new Error('The saved Figure placement disappeared after Stop.');
  const fallbackPng = await retainedPlacement.locator('img').getAttribute('src');
  await page.mouse.move(beforeMove.x + beforeMove.width / 2, beforeMove.y + beforeMove.height / 2);
  await page.mouse.down();
  await page.mouse.move(beforeMove.x + beforeMove.width / 2 + 45, beforeMove.y + beforeMove.height / 2 + 25, { steps: 3 });
  await page.mouse.up();
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Figure placement updated.');
  const afterMove = await retainedPlacement.boundingBox();
  expect(afterMove?.x).toBeGreaterThan(beforeMove.x);
  expect(afterMove?.y).toBeGreaterThan(beforeMove.y);
  await expect(retainedPlacement.locator('img')).toHaveAttribute('src', fallbackPng!);

  await page.getByRole('button', { name: 'Place on Sheet', exact: true }).click();
  const cachedCaptureBounds = await placementCapture.boundingBox();
  if (!cachedCaptureBounds) throw new Error('Cached Figure placement capture is not visible after the worker stopped.');
  await page.mouse.move(cachedCaptureBounds.x + 330, cachedCaptureBounds.y + 35);
  await page.mouse.down();
  await page.mouse.move(cachedCaptureBounds.x + 515, cachedCaptureBounds.y + 150, { steps: 4 });
  await page.mouse.up();
  const afterStopPlacements = page.getByRole('group', { name: 'Figure from Sheet1!D1' });
  await expect(afterStopPlacements).toHaveCount(2);
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Figure placed from its cached image.');
  await expect(page.locator('.python-figure-status')).toContainText('Retained the cached Figure image.');
  await expect(page.getByRole('button', { name: 'Rebuild Figure', exact: true })).toBeVisible();
  const cachedPlacement = afterStopPlacements.last();
  const cachedPng = await cachedPlacement.locator('img').getAttribute('src');
  const cachedBeforeMove = await cachedPlacement.boundingBox();
  if (!cachedBeforeMove) throw new Error('The cached Figure placement is not visible after worker loss.');
  await page.mouse.move(cachedBeforeMove.x + cachedBeforeMove.width / 2, cachedBeforeMove.y + cachedBeforeMove.height / 2);
  await page.mouse.down();
  await page.mouse.move(cachedBeforeMove.x + cachedBeforeMove.width / 2 + 45,
    cachedBeforeMove.y + cachedBeforeMove.height / 2 + 25, { steps: 3 });
  await page.mouse.up();
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Figure placement updated.');
  const cachedAfterMove = await cachedPlacement.boundingBox();
  expect(cachedAfterMove?.x).toBeGreaterThan(cachedBeforeMove.x);
  expect(cachedAfterMove?.y).toBeGreaterThan(cachedBeforeMove.y);
  await expect(cachedPlacement.locator('img')).toHaveAttribute('src', cachedPng!);

  const cachedResize = cachedPlacement.getByRole('button', { name: 'Resize figure bottom right' });
  const cachedResizeBounds = await cachedResize.boundingBox();
  if (!cachedResizeBounds) throw new Error('The cached Figure resize handle is not visible after worker loss.');
  const beforeCachedResize = await cachedPlacement.boundingBox();
  if (!beforeCachedResize) throw new Error('The cached Figure placement disappeared before resize.');
  const cachedRasterWidth = await cachedPlacement.locator('img').evaluate(image => (image as HTMLImageElement).naturalWidth);
  await page.mouse.move(cachedResizeBounds.x + cachedResizeBounds.width / 2, cachedResizeBounds.y + cachedResizeBounds.height / 2);
  await page.mouse.down();
  await page.mouse.move(cachedResizeBounds.x + cachedResizeBounds.width / 2 + 35,
    cachedResizeBounds.y + cachedResizeBounds.height / 2 + 25, { steps: 3 });
  await page.mouse.up();
  await expect(page.locator('.python-run-controls [role="status"]'))
    .toContainText(/^(?:Starting local Python runtime|Automatic · calculating mixed dependencies|Loading (?:Python runtime|NumPy|pandas|Matplotlib)(?: \(\d of 3\))?|Calculating .+|Updating .+)…$/iu);
  expect((await cachedPlacement.boundingBox())?.width).toBeGreaterThan(beforeCachedResize.width);
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run completed.', { timeout: 420_000 });
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Figure placement updated.');
  await expect.poll(() => cachedPlacement.locator('img').evaluate(image => (image as HTMLImageElement).naturalWidth))
    .toBeGreaterThan(cachedRasterWidth);
  await expect(cachedPlacement.locator('img')).not.toHaveAttribute('src', cachedPng!);

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
  await expect(page.locator('.python-run-controls [role="status"]')).toHaveText('Python run timed out.', { timeout: 180_000 });
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
