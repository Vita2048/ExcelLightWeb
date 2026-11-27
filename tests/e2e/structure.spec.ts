import { expect, test, type Page } from '@playwright/test';

async function go(page: Page, address: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(address);
  await name.press('Enter');
}

async function author(page: Page, address: string, value: string) {
  await go(page, address);
  const formula = page.getByRole('textbox', { name: 'Formula bar' });
  await formula.fill(value);
  await formula.press('Enter');
}

test('sheet context commands target the clicked sheet and restore cross-sheet references on undo', async ({ page }) => {
  await page.goto('/ExcelLightWeb/');
  await author(page, 'A1', '12');
  await page.getByRole('button', { name: 'New sheet' }).click();
  await expect(page.getByRole('grid', { name: 'Sheet2' })).toBeFocused();
  await author(page, 'A1', '=Sheet1!$A$1');
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click({ button: 'right' });
  const menu = page.getByRole('group', { name: 'Structure commands' });
  await menu.getByRole('button', { name: 'Rename Sheet' }).click();
  const name = page.getByRole('textbox', { name: 'Sheet name' });
  await name.fill("O'Brien");
  await name.press('Enter');
  await page.getByRole('button', { name: 'Sheet2', exact: true }).click();
  await go(page, 'A1');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue("='O''Brien'!$A$1");

  const source = page.getByRole('button', { name: "O'Brien", exact: true });
  await source.click({ button: 'right' });
  await menu.getByRole('button', { name: 'Move or Copy Sheet…', exact: true }).click();
  const place = page.getByRole('dialog', { name: 'Move or Copy Sheet', exact: true });
  await place.getByLabel('Before worksheet').selectOption({ label: 'Move to end' });
  await place.getByLabel('Create a copy').check();
  await place.getByRole('button', { name: 'OK' }).click();
  await expect(page.getByRole('grid', { name: 'Sheet1' })).toBeFocused();
  await expect(page.getByRole('gridcell', { name: 'A1, 12', exact: true })).toBeAttached();

  await source.click({ button: 'right' });
  await menu.getByRole('button', { name: 'Delete Sheet' }).click();
  const confirm = page.getByRole('dialog', { name: 'Delete Sheet' });
  await expect(confirm.getByRole('button')).toHaveText(['Delete', 'Cancel']);
  await expect(confirm.getByRole('button', { name: 'Cancel' })).toBeFocused();
  await confirm.getByRole('button', { name: 'Cancel' }).click();
  await expect(source).toBeVisible();
  await source.click({ button: 'right' });
  await menu.getByRole('button', { name: 'Delete Sheet' }).click();
  await confirm.getByRole('button', { name: 'Delete', exact: true }).click();
  await expect(source).toHaveCount(0);
  await page.getByRole('button', { name: 'Sheet2', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'A1, #REF!', exact: true })).toBeAttached();
  await page.getByRole('button', { name: 'Undo' }).click();
  await expect(source).toBeVisible();
  await expect(page.getByRole('gridcell', { name: 'A1, 12', exact: true })).toBeAttached();
});

test('cell and axis commands shift populated cells and update absolute dependents', async ({ page }) => {
  await page.goto('/ExcelLightWeb/');
  await author(page, 'B2', '27');
  await author(page, 'D1', '=$B$2');
  await go(page, 'B2');
  const grid = page.getByRole('grid');
  const header = Number(await grid.getAttribute('data-row-header'));
  const bounds = (await grid.boundingBox())!;
  await page.mouse.click(bounds.x + header + 96 + 40, bounds.y + 60, { button: 'right' });
  await page.getByRole('group', { name: 'Structure commands' }).getByRole('button', { name: 'Insert Cells…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Insert', exact: true });
  await expect(dialog.getByRole('radio', { name: 'Shift cells right' })).toBeChecked();
  await dialog.getByRole('radio', { name: 'Shift cells down' }).check();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(page.getByRole('gridcell', { name: 'B3, 27', exact: true })).toBeAttached();
  await go(page, 'D1');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('=$B$3');
  await go(page, 'B3');
  await page.getByRole('button', { name: 'Delete', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Delete Sheet Rows', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'D1, #REF!', exact: true })).toBeAttached();
  await grid.press('Control+z');
  await expect(page.getByRole('gridcell', { name: 'D1, 27', exact: true })).toBeAttached();
});

test('headers resize selected columns together, fit content, and navigate over hidden columns', async ({ page }) => {
  await page.goto('/ExcelLightWeb/');
  const grid = page.getByRole('grid');
  const header = Number(await grid.getAttribute('data-row-header'));
  const { x, y } = (await grid.boundingBox())!;
  const columnEdge = (column: number) => header + column * 96;
  await page.mouse.move(x + header + 40, y + 10);
  await expect(grid).toHaveCSS('cursor', /^url\(.+\) 8 14, default$/u);
  await page.mouse.move(x + 20, y + 38);
  await expect(grid).toHaveCSS('cursor', /^url\(.+\) 14 8, default$/u);
  await page.mouse.move(x + 20, y + 51);
  await expect(grid).toHaveCSS('cursor', 'row-resize');
  await page.mouse.move(x + header + 40, y + 38);
  await expect(grid).toHaveCSS('cursor', 'auto');
  await page.mouse.click(x + header + 40, y + 10);
  await page.keyboard.down('Shift');
  await page.mouse.click(x + header + 96 + 40, y + 10);
  await page.keyboard.up('Shift');
  await page.mouse.move(x + columnEdge(1) - 1, y + 10);
  await expect(grid).toHaveCSS('cursor', 'col-resize');
  await page.mouse.down();
  await page.mouse.move(x + columnEdge(1) + 29, y + 10);
  await page.mouse.up();
  const spacer = grid.locator('.worksheet-spacer');
  const baseWidth = header + 16_384 * 96;
  const spacerWidth = () => spacer.evaluate(element => Number.parseFloat((element as HTMLElement).style.width));
  // Chromium serializes a multi-million pixel length to six significant digits.
  await expect.poll(spacerWidth).toBeGreaterThan(baseWidth + 40);
  await expect.poll(spacerWidth).toBeLessThan(baseWidth + 80);
  await page.getByRole('button', { name: 'Undo' }).click();
  await expect.poll(async () => Math.abs(await spacerWidth() - baseWidth)).toBeLessThan(20);

  await author(page, 'B1', 'An entry long enough to make its column substantially wider');
  await page.mouse.dblclick(x + columnEdge(2) - 1, y + 10);
  await expect.poll(() => spacer.evaluate(element => element.getBoundingClientRect().width)).toBeGreaterThan(baseWidth);
  await page.getByRole('button', { name: 'Undo' }).click();
  await page.mouse.click(x + header + 96 + 40, y + 10, { button: 'right' });
  await page.getByRole('group', { name: 'Structure commands' }).getByRole('button', { name: 'Hide Columns', exact: true }).click();
  await go(page, 'A1');
  await grid.press('ArrowRight');
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('C1');
  await page.mouse.click(x + header + 96 + 40, y + 10, { button: 'right' });
  await page.getByRole('group', { name: 'Structure commands' }).getByRole('button', { name: 'Unhide Columns' }).click();
  await go(page, 'B1');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue(/An entry long enough/);
  await page.getByRole('button', { name: 'Format', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Row Height…', exact: true }).click();
  const sizeDialog = page.getByRole('dialog', { name: 'Row Height', exact: true });
  await expect(sizeDialog.getByRole('button')).toHaveText(['OK', 'Cancel']);
  const size = sizeDialog.getByLabel('Size in pixels');
  await expect(size).toHaveCSS('appearance', 'textfield');
  await size.fill('44');
  await sizeDialog.getByRole('button', { name: 'OK' }).click();
  await expect.poll(() => spacer.evaluate(element => Number.parseFloat((element as HTMLElement).style.height))).toBeGreaterThan(16_777_000);
  await expect.poll(() => spacer.evaluate(element => Number.parseFloat((element as HTMLElement).style.height))).toBeLessThan(16_777_216);
});
