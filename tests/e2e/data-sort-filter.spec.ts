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

async function enter(page: Page, rows: readonly (readonly [string, string])[]) {
  for (const [address, value] of rows) await go(page, address, value);
}

async function extend(page: Page, address: string, down: number, right: number) {
  await go(page, address);
  for (let index = 0; index < right; index++) await page.getByRole('grid').press('Shift+ArrowRight');
  for (let index = 0; index < down; index++) await page.getByRole('grid').press('Shift+ArrowDown');
}

test('filters records from the Data tab funnels, updates statistics and SUBTOTAL, then clears and undoes', async ({ page }) => {
  await page.goto(path);
  await enter(page, [
    ['A1', 'Name'], ['B1', 'Amount'], ['C1', 'Dept'],
    ['A2', 'aaa'], ['B2', '10'], ['C2', 'Ops'],
    ['A3', 'bbb'], ['B3', '20'], ['C3', 'Sales'],
    ['A4', 'ccc'], ['B4', '30'], ['C4', 'Ops'],
    ['A6', 'Total'], ['B6', '=SUBTOTAL(109,B2:B4)'],
  ]);
  await expect(page.getByRole('gridcell', { name: 'B6, 60' })).toBeAttached();
  await go(page, 'A1');

  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await expect(page.locator('.ribbon-panel:not([hidden]) legend')).toHaveText(['Sort & Filter']);
  await expect(page.getByRole('button', { name: 'A to Z', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Z to A', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Sort', exact: true })).toHaveClass(/ribbon-large/);
  const toggle = page.getByRole('button', { name: 'Filter', exact: true });
  await expect(toggle).toHaveClass(/ribbon-large/);
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');
  await expect(page.getByRole('button', { name: 'Clear filter' })).toBeDisabled();
  await expect(page.getByRole('button', { name: 'Reapply filter' })).toBeDisabled();
  await expect(page.getByRole('button', { name: 'Filter Name' })).toHaveCount(0);

  await toggle.click();
  await expect(toggle).toHaveAttribute('aria-pressed', 'true');
  await expect(page.locator('.edit-status')).toHaveText('Filter applied to A1:C4. Choose values from a header funnel.');
  for (const header of ['Name', 'Amount', 'Dept']) {
    await expect(page.getByRole('button', { name: 'Filter ' + header, exact: true })).toBeVisible();
  }

  await page.getByRole('button', { name: 'Filter Dept', exact: true }).click();
  const dropdown = page.getByRole('dialog', { name: 'Filter Dept', exact: true });
  await expect(dropdown.getByRole('button', { name: 'Sort A to Z' })).toBeVisible();
  await expect(dropdown.getByRole('button', { name: 'Sort Z to A' })).toBeVisible();
  await expect(dropdown.getByRole('button', { name: 'Clear Filter From "Dept"' })).toBeDisabled();
  await expect(dropdown.getByRole('searchbox', { name: 'Search filter values' })).toBeFocused();
  await expect(dropdown.getByRole('checkbox', { name: '(Select All)' })).toBeChecked();
  await expect(dropdown.getByRole('checkbox', { name: 'Ops' })).toBeChecked();
  await expect(dropdown.getByRole('checkbox', { name: 'Sales' })).toBeChecked();
  await dropdown.getByRole('checkbox', { name: 'Sales' }).click();
  await dropdown.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dropdown).toBeHidden();

  await expect(page.locator('.edit-status')).toHaveText('Filtered Dept. 2 of 3 records');
  await expect(page.getByRole('gridcell', { name: 'A2, aaa' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'A3, bbb' })).toHaveCount(0);
  await expect(page.getByRole('gridcell', { name: 'A4, ccc' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B6, 40' })).toBeAttached();
  await expect(page.getByRole('button', { name: 'Clear filter' })).toBeEnabled();
  await expect(page.getByRole('button', { name: 'Reapply filter' })).toBeEnabled();
  await expect(page.getByRole('button', { name: 'Filter Dept', exact: true })).toHaveClass(/filter-active/);

  await extend(page, 'B2', 2, 0);
  const stats = page.getByLabel('Selection statistics');
  await expect(stats).toContainText('Sum: 40');
  await expect(stats).toContainText('Average: 20');
  await expect(stats).toContainText('Count: 2');

  await page.getByRole('button', { name: 'Reapply filter' }).click();
  await expect(page.locator('.edit-status')).toHaveText('Filter reapplied. 2 of 3 records');

  await page.getByRole('button', { name: 'Clear filter' }).click();
  await expect(page.locator('.edit-status')).toHaveText('Filter cleared. Manually hidden rows stay hidden.');
  await expect(page.getByRole('gridcell', { name: 'A3, bbb' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B6, 60' })).toBeAttached();

  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'A3, bbb' })).toHaveCount(0);
  await expect(page.getByRole('gridcell', { name: 'B6, 40' })).toBeAttached();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'A3, bbb' })).toBeAttached();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');
});

test('uses the desktop Sort dialog, its options, the sort warning and the Home Sort & Filter menu', async ({ page }) => {
  await page.goto(path);
  await enter(page, [
    ['A1', 'Name'], ['B1', 'Score'],
    ['A2', 'a'], ['B2', '1'],
    ['A3', 'b'], ['B3', '2'],
  ]);
  await extend(page, 'A1', 2, 1);
  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  const menu = page.getByRole('button', { name: 'Sort & Filter', exact: true });
  await expect(menu).toHaveClass(/ribbon-large/);
  await menu.click();
  const flyout = page.getByRole('menu', { name: 'Sort & Filter', exact: true });
  await expect(flyout.getByRole('menuitem')).toHaveText(['Sort A to Z', 'Sort Z to A', 'Custom Sort...', 'Filter', 'Clear', 'Reapply']);
  await expect(flyout.getByRole('menuitem', { name: 'Clear', exact: true })).toBeDisabled();
  await expect(flyout.getByRole('menuitem', { name: 'Reapply', exact: true })).toBeDisabled();
  await flyout.getByRole('menuitem', { name: 'Custom Sort...' }).click();

  const sort = page.getByRole('dialog', { name: 'Sort', exact: true });
  await expect(sort.getByText('Column', { exact: true })).toBeVisible();
  await expect(sort.getByText('Sort On', { exact: true })).toBeVisible();
  await expect(sort.getByText('Order', { exact: true })).toBeVisible();
  await expect(sort.getByRole('button', { name: 'Options...' })).toBeVisible();
  await expect(sort.getByRole('checkbox', { name: 'My data has headers' })).toBeChecked();
  await expect(sort.getByRole('combobox', { name: 'Sort column' })).toHaveValue('1');
  await expect(sort.getByRole('combobox', { name: 'Sort column' }).locator('option')).toHaveText(['Name', 'Score']);
  await expect(sort.getByRole('combobox', { name: 'Sort on' })).toHaveValue('values');
  await expect(sort.getByRole('combobox', { name: 'Sort order' })).toHaveValue('asc');
  await expect(sort.getByRole('combobox', { name: 'Sort order' }).locator('option')).toHaveText(['A to Z', 'Z to A']);

  await sort.getByRole('button', { name: 'Options...' }).click();
  await expect(sort).toBeHidden();
  const options = page.getByRole('dialog', { name: 'Sort Options', exact: true });
  await expect(options.getByText('Orientation', { exact: true })).toBeVisible();
  await expect(options.getByRole('radio', { name: 'Sort top to bottom' })).toBeChecked();
  await expect(options.getByRole('checkbox', { name: 'Case sensitive' })).not.toBeChecked();
  await options.getByRole('checkbox', { name: 'Case sensitive' }).check();
  await options.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(options).toBeHidden();
  await expect(sort).toBeVisible();
  await sort.getByRole('combobox', { name: 'Sort order' }).selectOption('desc');
  await sort.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(sort).toBeHidden();
  await expect(page.locator('.edit-status')).toHaveText('Sorted A1:B3 by column A, largest to smallest. Records moved together.');
  await expect(page.getByRole('gridcell', { name: 'A2, b' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B2, 2' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'A3, a' })).toBeAttached();

  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByRole('gridcell', { name: 'A2, a' })).toBeAttached();
  await page.getByRole('button', { name: 'Sort & Filter', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Sort A to Z' }).click();
  await expect(page.getByRole('gridcell', { name: 'A2, a' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'A3, b' })).toBeAttached();

  await page.getByRole('button', { name: 'Sort & Filter', exact: true }).click();
  const filterItem = page.getByRole('menuitem', { name: 'Filter', exact: true });
  await filterItem.click();
  await expect(page.locator('.edit-status')).toHaveText('Filter applied to A1:B3. Choose values from a header funnel.');
  await page.getByRole('button', { name: 'Sort & Filter', exact: true }).click();
  await expect(page.getByRole('menuitem', { name: 'Reapply', exact: true })).toBeEnabled();
  await expect(page.getByRole('menuitem', { name: 'Clear', exact: true })).toBeDisabled();
  await page.getByRole('menuitem', { name: 'Filter', exact: true }).click();
});

test('sorts next to surrounding data with the desktop Sort Warning and honors table filter buttons', async ({ page }) => {
  await page.goto(path);
  await enter(page, [
    ['A1', 'x'], ['B1', 'keep'],
    ['A2', 'b'], ['B2', '2'],
    ['A3', 'a'], ['B3', '1'],
    ['D2', 'n'], ['E2', 'u'],
    ['D3', 'p'], ['E3', 'v'],
    ['D4', 'q'],
  ]);
  await extend(page, 'A1', 2, 0);
  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.getByRole('button', { name: 'A to Z', exact: true }).click();
  const warning = page.getByRole('dialog', { name: 'Sort Warning', exact: true });
  await expect(warning).toBeVisible();
  await expect(warning.getByText('There is data next to A1:A3. Choose which data to sort.')).toBeVisible();
  await expect(warning.getByRole('radio', { name: 'Expand the selection' })).toBeChecked();
  await expect(warning.getByText('Sort range: A1:B3')).toBeVisible();
  await warning.getByRole('radio', { name: 'Continue with the current selection' }).check();
  await expect(warning.getByText('Sort range: A1:A3. Adjacent columns will stay in place.')).toBeVisible();
  await warning.getByRole('button', { name: 'Sort', exact: true }).click();
  await expect(warning).toBeHidden();
  await expect(page.getByRole('gridcell', { name: 'A1, a' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'A2, b' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'A3, x' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B1, keep' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B2, 2' })).toBeAttached();

  await go(page, 'A1');
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Table', exact: true }).click();
  const create = page.getByRole('dialog', { name: 'Create Table' });
  await create.getByRole('textbox', { name: 'Range', exact: true }).fill("='Sheet1'!$A$1:$B$3");
  await create.getByRole('checkbox', { name: 'My table has headers' }).check();
  await create.getByRole('button', { name: 'OK', exact: true }).click();
  await go(page, 'A1');
  await expect(page.getByRole('checkbox', { name: 'Filter Button' })).toBeChecked();
  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  const toggle = page.getByRole('button', { name: 'Filter', exact: true });
  await expect(toggle).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByRole('button', { name: 'Filter a', exact: true })).toBeVisible();
  await toggle.click();
  await page.getByRole('tab', { name: 'Table Design', exact: true }).click();
  await expect(page.getByRole('checkbox', { name: 'Filter Button' })).not.toBeChecked();
  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');
  await expect(page.getByRole('button', { name: 'Filter a', exact: true })).toHaveCount(0);
  await toggle.click();
  await expect(toggle).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByRole('button', { name: 'Filter a', exact: true })).toBeVisible();

  // Outside the table the table's filter leaves the toggle off, and pressing it filters the selected cells.
  await extend(page, 'D2', 1, 1);
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');
  await expect(page.getByRole('button', { name: 'Filter a', exact: true })).toBeVisible();
  await toggle.click();
  await expect(page.locator('.edit-status')).toHaveText('Filter applied to D2:E3. Choose values from a header funnel.');
  await expect(toggle).toHaveAttribute('aria-pressed', 'true');

  await toggle.click();
  await expect(page.locator('.edit-status')).toHaveText('Filter removed. Manually hidden rows stay hidden.');
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');

  // A single non-empty cell filters the data block around it.
  await go(page, 'D3');
  await toggle.click();
  await expect(page.locator('.edit-status')).toHaveText('Filter applied to D2:E4. Choose values from a header funnel.');
  await expect(toggle).toHaveAttribute('aria-pressed', 'true');
  await toggle.click();

  await go(page, 'D5');
  await toggle.click();
  await expect(page.locator('.edit-status')).toHaveText('Select a range with data before turning on the filter.');
  await expect(toggle).toHaveAttribute('aria-pressed', 'false');
});

test('opens the Sort dialog with the sort already applied to the table', async ({ page }) => {
  await page.goto(path);
  await enter(page, [
    ['A1', 'Name'], ['B1', 'Score'],
    ['A2', 'bbb'], ['B2', '1'],
    ['A3', 'aaa'], ['B3', '2'],
  ]);
  await go(page, 'A1');
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Table', exact: true }).click();
  const create = page.getByRole('dialog', { name: 'Create Table' });
  await create.getByRole('textbox', { name: 'Range', exact: true }).fill("='Sheet1'!$A$1:$B$3");
  await create.getByRole('checkbox', { name: 'My table has headers' }).check();
  await create.getByRole('button', { name: 'OK', exact: true }).click();
  await go(page, 'A1');
  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  await page.getByRole('button', { name: 'Filter Score', exact: true }).click();
  await page.getByRole('dialog', { name: 'Filter Score', exact: true })
    .getByRole('button', { name: 'Sort Largest to Smallest' }).click();
  await expect(page.getByRole('gridcell', { name: 'A2, aaa' })).toBeAttached();

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Sort & Filter', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Custom Sort...' }).click();
  const sort = page.getByRole('dialog', { name: 'Sort', exact: true });
  await expect(sort.getByRole('combobox', { name: 'Sort column' })).toHaveValue('2');
  await expect(sort.getByRole('combobox', { name: 'Sort order' })).toHaveValue('desc');
  await expect(sort.getByRole('checkbox', { name: 'My data has headers' })).toBeChecked();
  await sort.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(sort).toBeHidden();
});
