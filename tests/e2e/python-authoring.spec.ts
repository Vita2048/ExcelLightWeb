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

async function point(page: Page, row: number, column: number) {
  const node = page.getByRole('grid');
  const grid = (await node.boundingBox())!;
  const header = Number(await node.getAttribute('data-row-header'));
  return { x: grid.x + header + (column - 0.5) * 96, y: grid.y + 26 + (row - 0.5) * 26 };
}

function editor(page: Page) {
  return page.locator('.python-editor .view-lines');
}

async function replaceLine(page: Page, text: string) {
  await page.locator('.python-editor .view-line').first().click({ clickCount: 3 });
  await page.keyboard.insertText(text);
}

function undoKey(project: string): string {
  return project === 'webkit' ? 'Meta+z' : 'Control+z';
}

async function ready(page: Page) {
  await page.getByRole('grid').waitFor();
  // Let the browser-local workspace lease and initial calculation settle before typing.
  await enter(page, 'B2', 'seed');
  await page.waitForTimeout(250);
}

async function createAnchor(page: Page, source = '=PY') {
  await enter(page, 'A1', source);
  await page.waitForTimeout(250);
  await go(page, 'A1');
  await expect(page.locator('.python-panel')).toBeVisible();
  await expect(page.locator('.python-editor .monaco-editor')).toBeVisible();
  await expect(editor(page)).toHaveText('');
}

test('=PY opens a lazy local Monaco panel, commits through a selection barrier and clears on ordinary cells', async ({ page }) => {
  const external: string[] = [];
  let origin = '';
  page.on('request', request => {
    const url = request.url();
    if (origin && !url.startsWith(origin) && !url.startsWith('blob:') && !url.startsWith('data:')) external.push(url);
  });
  await page.goto(path);
  origin = new URL(page.url()).origin;
  await ready(page);
  await createAnchor(page);
  await expect(editor(page)).toHaveText('');
  await expect(page.getByRole('gridcell', { name: 'A1, blank', exact: true })).toBeAttached();
  await page.getByRole('tab', { name: 'Formulas' }).click();
  const showFormulas = page.getByRole('button', { name: 'Show Formulas' });
  await showFormulas.click();
  await expect(page.getByRole('gridcell', { name: 'A1, =PY()', exact: true })).toBeAttached();
  await showFormulas.click();
  await expect(page.getByRole('gridcell', { name: 'A1, blank', exact: true })).toBeAttached();

  await replaceLine(page, 'xl("B1")');
  await expect(editor(page)).toHaveText('xl("B1")');

  await go(page, 'B2');
  await expect(page.locator('.python-panel')).toBeVisible();
  await expect(editor(page)).toHaveText('');
  await go(page, 'A1');
  await expect(editor(page)).toHaveText('xl("B1")');

  await page.getByRole('button', { name: 'Hide Python workspace', exact: true }).click();
  await expect(page.locator('.python-panel')).toHaveCount(0);
  await page.waitForTimeout(300);
  await expect(page.locator('.python-panel')).toHaveCount(0);
  await page.getByRole('tab', { name: 'Formulas' }).click();
  await expect(page.locator('.ribbon-panel:not([hidden]) fieldset > legend'))
    .toHaveText(['Function Library', 'Formula Auditing', 'Defined Names', 'Calculation', 'Python']);
  const pythonGroup = page.getByRole('group', { name: 'Python', exact: true });
  await expect(pythonGroup).toBeVisible();
  const pythonToggle = pythonGroup.getByRole('button', { name: 'Show or hide Python workspace' });
  await expect(pythonToggle).toHaveAttribute('aria-pressed', 'false');
  await expect(pythonToggle.locator('svg.ribbon-glyph')).toBeVisible();
  await expect(pythonToggle.locator('svg.ribbon-glyph path')).toHaveCount(2);
  await pythonToggle.click();
  await expect(pythonToggle).toHaveAttribute('aria-pressed', 'true');
  await expect(page.locator('.python-panel')).toBeVisible();
  await expect(editor(page)).toHaveText('xl("B1")');
  expect(external).toEqual([]);
});

test('xl range picking works across sheets, returns to the anchor, and keeps local undo separate', async ({ page }, testInfo) => {
  await page.goto(path);
  await ready(page);
  await page.getByRole('button', { name: 'New sheet' }).click();
  await page.getByRole('button', { name: 'Sheet2', exact: true }).dblclick();
  const rename = page.getByRole('textbox', { name: 'Sheet name' });
  await rename.fill("O'Brien");
  await rename.press('Enter');
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click();
  await createAnchor(page);

  const lines = page.locator('.python-editor .view-lines');
  await replaceLine(page, 'xl("")');
  await page.keyboard.press('End');
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('ArrowLeft');
  await page.waitForTimeout(150);

  const b1 = await point(page, 1, 2);
  const b2 = await point(page, 2, 2);
  await page.mouse.move(b1.x, b1.y);
  await page.mouse.down();
  await page.mouse.move(b2.x, b2.y, { steps: 4 });
  await page.mouse.up();
  await expect(lines).toHaveText('xl("B1:B2")');

  await lines.click();
  await page.keyboard.press(undoKey(testInfo.project.name));
  await expect(lines).toHaveText('xl("")');

  await page.getByRole('button', { name: "O'Brien", exact: true }).click();
  await expect(page.getByRole('button', { name: "O'Brien", exact: true })).toHaveAttribute('aria-current', 'true');
  await expect(lines).toHaveText('xl("")');
  const remote = await point(page, 1, 1);
  await page.mouse.move(remote.x, remote.y);
  await page.mouse.down();
  await page.mouse.up();
  await expect(lines).toHaveText('xl("\'O\'\'Brien\'!A1")');
  await expect(page.getByRole('button', { name: 'Sheet1', exact: true })).toHaveAttribute('aria-current', 'true');

  await lines.click();
  await page.keyboard.press(undoKey(testInfo.project.name));
  await expect(lines).toHaveText('xl("")');
});

test('saving flushes the pending source before the workbook copy is written', async ({ page }) => {
  await page.goto(path);
  await ready(page);
  await createAnchor(page);

  await replaceLine(page, 'pending = 1');
  await page.getByRole('group', { name: 'Quick Access' }).getByRole('button', { name: 'Save', exact: true }).click();
  const nameDialog = page.getByRole('dialog', { name: 'Save to this browser' });
  await expect(nameDialog).toBeVisible();
  await nameDialog.getByRole('textbox', { name: 'Workbook name' }).fill('Python draft');
  await nameDialog.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Saved to this browser');

  await page.reload();
  await page.getByRole('grid').waitFor();
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Open recent Python draft' }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Python draft');
  await go(page, 'A1');
  await expect(page.locator('.python-panel')).toBeVisible();
  await expect(editor(page)).toHaveText('pending = 1');
});
