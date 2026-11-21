import { expect, test, type Page } from '@playwright/test';

const path = '/ExcelLightWeb/';
test.describe.configure({ timeout: 60_000 });

async function go(page: Page, address: string, value?: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(address);
  await name.press('Enter');
  if (value !== undefined) {
    const formula = page.getByRole('textbox', { name: 'Formula bar' });
    await formula.fill(value);
    await formula.press('Enter');
  }
}

async function openFind(page: Page, replace = false) {
  await page.keyboard.press(replace ? 'Control+h' : 'Control+f');
  return page.getByRole('dialog', { name: 'Find and Replace' });
}

test('finds sheet and workbook cells in values and formulas with case, whole-cell and column order options', async ({ page }) => {
  await page.goto(path);
  await go(page, 'A1', 'apple');
  await go(page, 'A2', '=A1&" pie"');
  await go(page, 'A3', 'Apple');
  await page.getByRole('button', { name: 'New sheet' }).click();
  await go(page, 'A1', 'Apple');
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click();
  const grid = page.getByRole('grid');
  await go(page, 'A1');

  const menu = page.getByRole('button', { name: 'Find & Select', exact: true });
  await expect(menu).toHaveClass(/ribbon-large/);
  await menu.click();
  await expect(page.getByRole('menuitem')).toHaveText(['Find...', 'Replace...']);
  await expect(page.getByRole('menuitem', { name: 'Replace...', exact: true }).locator('svg path').nth(1))
    .toHaveAttribute('d', /M5 7 C1 8 1 16 6 17 H10 M7 14 L10 17 L7 20/);
  await page.getByRole('menuitem', { name: 'Find...', exact: true }).click();
  const find = page.getByRole('dialog', { name: 'Find and Replace' });
  await expect(find.getByRole('tab', { name: 'Find', exact: true })).toHaveAttribute('aria-selected', 'true');
  await expect(find.getByRole('button', { name: 'Replace All' })).toHaveCount(0);
  await expect(find.getByRole('button', { name: 'Find Next' })).toHaveClass(/find-primary/);
  const query = find.getByRole('textbox', { name: 'Find what' });
  await find.getByRole('button', { name: 'Options >>' }).click();
  await expect(find.getByLabel('Look in')).toHaveValue('values');
  await expect(find.getByRole('option', { name: 'Values' })).toHaveCount(1);
  const withinBox = (await find.getByLabel('Within').boundingBox())!;
  const searchBox = (await find.getByLabel('Search', { exact: true }).boundingBox())!;
  const lookBox = (await find.getByLabel('Look in').boundingBox())!;
  expect(searchBox.y).toBeGreaterThan(withinBox.y);
  expect(lookBox.y).toBeGreaterThan(withinBox.y + withinBox.height - 1);
  expect(lookBox.y).toBeGreaterThan(searchBox.y + searchBox.height - 1);
  await query.fill('A1');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('0 cell(s) found.');
  await find.getByLabel('Look in').selectOption('formulas');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('1 cell(s) found.');
  const results = find.getByRole('table', { name: 'Find All results' });
  await expect(results.getByRole('columnheader')).toHaveText(['Book', 'Sheet', 'Name', 'Cell', 'Value', 'Formula']);
  await expect(results.getByRole('row')).toHaveCount(2);
  await expect(results.getByRole('row').nth(1)).toContainText('Book1');
  await expect(results.getByRole('row').nth(1)).toContainText('Sheet1');
  await expect(results.getByRole('row').nth(1)).toContainText('$A$2');
  await expect(results.getByRole('row').nth(1)).toContainText('apple pie');
  await expect(results.getByRole('row').nth(1)).toContainText('=A1&" pie"');
  await query.fill('apple');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('0 cell(s) found.');

  await find.getByLabel('Look in').selectOption('values');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('3 cell(s) found.');
  await query.fill('apple pie');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('1 cell(s) found.');
  await expect(results.getByRole('row').nth(1)).toContainText('$A$2');
  await find.getByLabel('Look in').selectOption('formulas');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('0 cell(s) found.');

  await find.getByLabel('Look in').selectOption('values');
  await query.fill('apple');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('3 cell(s) found.');
  await expect(results.getByRole('row')).toHaveCount(4);
  await expect(results.getByRole('row').nth(2)).toContainText('apple pie');
  await expect(results.getByRole('row').nth(2)).toContainText('=A1&" pie"');
  const bookHeader = results.getByRole('columnheader', { name: 'Book' });
  const before = (await bookHeader.boundingBox())!.width;
  await results.getByRole('separator', { name: 'Resize Book column' }).focus();
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowRight');
  await expect.poll(async () => (await bookHeader.boundingBox())!.width).toBeGreaterThan(before + 8);

  await find.getByRole('checkbox', { name: 'Match entire cell contents' }).check();
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('2 cell(s) found.');
  await expect(results.getByRole('row')).toHaveCount(3);
  await find.getByRole('checkbox', { name: 'Match entire cell contents' }).uncheck();
  await find.getByRole('checkbox', { name: 'Match case' }).check();
  await query.fill('Apple');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('1 cell(s) found.');
  await expect(results.getByRole('row').nth(1)).toContainText('$A$3');

  await find.getByLabel('Search').selectOption('columns');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('1 cell(s) found.');
  await find.getByLabel('Within').selectOption('workbook');
  await find.getByRole('button', { name: 'Find All' }).click();
  await expect(find.getByRole('status')).toHaveText('2 cell(s) found.');
  await expect(results.getByRole('row')).toHaveCount(3);
  await expect(results.getByRole('row').nth(2)).toContainText('Sheet2');
  await expect(results.getByRole('row').nth(2)).toContainText('$A$1');

  await page.keyboard.press('Escape');
  await expect(find).toBeHidden();
  await expect(grid).toBeFocused();
  await go(page, 'A1');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('apple');
});

test('navigates matches with Find Next, wraps, and replaces one or all matches in one undo step', async ({ page }) => {
  await page.goto(path);
  await go(page, 'A1', '10');
  await go(page, 'A3', '10');
  await go(page, 'A5', '10');
  const name = page.getByRole('textbox', { name: 'Name box' });
  await go(page, 'A1');

  const find = await openFind(page);
  await find.getByRole('textbox', { name: 'Find what' }).fill('10');
  await find.getByRole('button', { name: 'Find Next' }).click();
  await expect(find.getByRole('status')).toHaveText('3 match(es). A3');
  await expect(name).toHaveValue('A3');
  await find.getByRole('button', { name: 'Find Next' }).click();
  await expect(find.getByRole('status')).toHaveText('3 match(es). A5');
  await find.getByRole('button', { name: 'Find Next' }).click();
  await expect(find.getByRole('status')).toHaveText('3 match(es). A1');
  await expect(name).toHaveValue('A1');
  await page.keyboard.press('Escape');

  const replace = await openFind(page, true);
  await expect(replace.getByRole('tab', { name: 'Replace', exact: true })).toHaveAttribute('aria-selected', 'true');
  const query = replace.getByRole('textbox', { name: 'Find what' });
  await query.fill('10');
  await replace.getByRole('textbox', { name: 'Replace with' }).fill('7');
  await replace.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect(replace.getByRole('status')).toHaveText('Replaced 1 cell(s).');
  await expect(name).toHaveValue('A3');
  await replace.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect(name).toHaveValue('A5');
  await replace.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect(replace.getByRole('status')).toHaveText('Replaced 1 cell(s).');
  await expect(name).toHaveValue('A5');

  await replace.getByRole('button', { name: 'Options >>' }).click();
  const look = replace.getByLabel('Look in');
  await expect(look).toBeDisabled();
  await expect(look.getByRole('option')).toHaveText(['Formulas']);
  await replace.getByRole('button', { name: 'Options <<' }).click();
  await page.keyboard.press('Escape');
  await expect(replace).toBeHidden();

  for (const address of ['A1', 'A3', 'A5']) {
    await go(page, address);
    await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('7');
  }
  for (let step = 0; step < 3; step++) await page.getByRole('button', { name: 'Undo' }).click();
  for (const address of ['A1', 'A3', 'A5']) {
    await go(page, address);
    await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('10');
  }

  const all = await openFind(page, true);
  await all.getByRole('textbox', { name: 'Find what' }).fill('10');
  await all.getByRole('textbox', { name: 'Replace with' }).fill('1000');
  await all.getByRole('button', { name: 'Replace All' }).click();
  await expect(all.getByRole('status')).toHaveText('Replaced 3 cell(s).');
  await page.keyboard.press('Escape');
  for (const address of ['A1', 'A3', 'A5']) {
    await go(page, address);
    await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('1000');
  }
  await page.getByRole('button', { name: 'Undo' }).click();
  for (const address of ['A1', 'A3', 'A5']) {
    await go(page, address);
    await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('10');
  }
});

test('applies the desktop Zoom dialog presets, custom percentages and Fit selection, and keeps hidden columns hidden', async ({ page }) => {
  await page.goto(path);
  await go(page, 'A1', 'Wide enough label for a column');
  await go(page, 'B1', 'B');
  const grid = page.getByRole('grid');
  const zoom = page.getByRole('slider', { name: 'Zoom' });
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  const dialogButton = page.getByRole('button', { name: 'Zoom...', exact: true });
  await expect(dialogButton).toHaveClass(/ribbon-large/);
  await expect(page.getByRole('button', { name: '100%', exact: true })).toHaveClass(/ribbon-large/);
  await expect(page.getByRole('button', { name: 'Zoom to Selection', exact: true })).toHaveClass(/ribbon-large/);

  const dialog = page.getByRole('dialog', { name: 'Zoom', exact: true });
  await dialogButton.click();
  await expect(dialog.getByRole('heading', { name: 'Magnification' })).toBeVisible();
  await expect(dialog.getByRole('radio')).toHaveCount(7);
  await expect(dialog.getByRole('radio', { name: '100%' })).toBeChecked();
  await dialog.getByRole('radio', { name: '200%' }).check();
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect(zoom).toHaveAttribute('aria-valuetext', '200%');
  await expect(page.getByRole('button', { name: /^Zoom percentage/ })).toHaveText('200%');

  await page.getByRole('button', { name: '100%', exact: true }).click();
  await expect(zoom).toHaveAttribute('aria-valuetext', '100%');
  await page.getByRole('button', { name: 'Zoom to Selection', exact: true }).click();
  await expect(zoom).toHaveAttribute('aria-valuetext', '400%');

  await page.getByRole('button', { name: /^Zoom percentage/ }).click();
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('radio', { name: 'Custom' })).toBeChecked();
  const custom = dialog.getByRole('spinbutton', { name: 'Custom zoom percentage, 10 to 400' });
  await custom.fill('55');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(zoom).toHaveAttribute('aria-valuetext', '55%');

  await page.getByRole('button', { name: /^Zoom percentage/ }).click();
  await custom.fill('500');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(zoom).toHaveAttribute('aria-valuetext', '400%');
  await page.getByRole('button', { name: /^Zoom percentage/ }).click();
  await custom.fill('5');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(zoom).toHaveAttribute('aria-valuetext', '10%');

  await page.getByRole('button', { name: '100%', exact: true }).click();
  await expect(zoom).toHaveAttribute('aria-valuetext', '100%');
  await go(page, 'A1');
  const bounds = (await grid.boundingBox())!;
  await page.mouse.click(bounds.x + 175, bounds.y + 10, { button: 'right' });
  await page.getByRole('group', { name: 'Structure commands' }).getByRole('button', { name: 'Hide Columns', exact: true }).click();
  await page.mouse.click(bounds.x + 20, bounds.y + 65, { button: 'right' });
  await page.getByRole('group', { name: 'Structure commands' }).getByRole('button', { name: 'Hide Rows', exact: true }).click();
  await expect(grid.locator('[role="row"][aria-rowindex="2"]')).toHaveCount(0);
  await page.getByRole('button', { name: /^Zoom percentage/ }).click();
  await dialog.getByRole('spinbutton', { name: 'Custom zoom percentage, 10 to 400' }).fill('150');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(zoom).toHaveAttribute('aria-valuetext', '150%');
  await expect(grid.locator('[role="gridcell"][aria-colindex="2"]')).toHaveCount(0);
  await expect(grid.locator('[role="gridcell"][aria-colindex="1"]').first()).toBeAttached();
  await go(page, 'A1');
  await grid.press('ArrowRight');
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('C1');
});

test('freezes and unfreezes panes per worksheet from the View tab and keeps frozen bands pinned', async ({ page }) => {
  await page.goto(path);
  await go(page, 'C1', 'Pinned');
  await go(page, 'A1');
  await page.getByRole('button', { name: 'Fill Color', exact: true }).click();
  await page.getByRole('group', { name: 'Fill Color', exact: true }).getByRole('button', { name: 'Color #156082', exact: true }).click();
  const grid = page.getByRole('grid');
  const canvas = page.locator('.worksheet-canvas');
  const pixel = (x: number, y: number) => canvas.evaluate((element: HTMLCanvasElement, point: { x: number; y: number }) => Array.from(
    element.getContext('2d')!.getImageData(point.x * devicePixelRatio, point.y * devicePixelRatio, 1, 1).data), { x, y });
  const sample = () => pixel(52, 34);
  await expect.poll(sample).toEqual([21, 96, 130, 255]);

  await page.getByRole('tab', { name: 'View', exact: true }).click();
  const freeze = page.getByRole('button', { name: 'Freeze Panes', exact: true });
  await expect(freeze).toHaveClass(/ribbon-large/);
  await freeze.click();
  await page.getByRole('menuitem', { name: 'Freeze Top Row', exact: true }).click();
  await expect(page.locator('.edit-status')).toHaveText('Froze the top row.');
  // The desktop paints a thicker frozen divider over the grid at the pinned band's edge.
  await expect.poll(() => pixel(100, 51)).toEqual([160, 160, 160, 255]);
  await grid.evaluate(node => { node.scrollTop = 400; });
  await expect.poll(sample).toEqual([21, 96, 130, 255]);

  await freeze.click();
  await expect(page.getByRole('menuitem', { name: 'Freeze Panes', exact: true })).toHaveCount(0);
  await page.getByRole('menuitem', { name: 'Unfreeze Panes', exact: true }).click();
  await expect(page.locator('.edit-status')).toHaveText('Unfroze panes.');
  await grid.evaluate(node => { node.scrollTop = 400; });
  await expect.poll(sample).not.toEqual([21, 96, 130, 255]);
  await expect.poll(() => pixel(100, 51)).not.toEqual([160, 160, 160, 255]);

  await page.getByRole('button', { name: 'New sheet' }).click();
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await freeze.click();
  await page.getByRole('menuitem', { name: 'Freeze First Column', exact: true }).click();
  await expect(page.locator('.edit-status')).toHaveText('Froze the first column.');
  await expect.poll(() => pixel(139, 100)).toEqual([160, 160, 160, 255]);
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click();
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await freeze.click();
  await expect(page.getByRole('menuitem', { name: 'Freeze Panes', exact: true })).toBeVisible();
  await expect(page.getByRole('menuitem', { name: 'Unfreeze Panes', exact: true })).toHaveCount(0);
  await page.keyboard.press('Escape');

  await page.setViewportSize({ width: 420, height: 320 });
  await go(page, 'B2');
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await freeze.click();
  await page.getByRole('menuitem', { name: 'Freeze Panes', exact: true }).click();
  await expect(page.locator('.edit-status')).toHaveText('Froze panes.');
  await grid.evaluate(node => { node.scrollTop = 300; node.scrollLeft = 300; });
  await expect.poll(sample).toEqual([21, 96, 130, 255]);
});
