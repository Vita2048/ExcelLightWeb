import { expect, test, type Page } from '@playwright/test';
import axe from 'axe-core';

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

async function seedSalesData(page: Page) {
  const rows: readonly (readonly [string, string])[] = [
    ['A1', 'Region'], ['B1', 'Product'], ['C1', 'Sales'],
    ['A2', 'North'], ['B2', 'Widget'], ['C2', '100'],
    ['A3', 'South'], ['B3', 'Gadget'], ['C3', '200'],
    ['A4', 'North'], ['B4', 'Gadget'], ['C4', '150'],
    ['A5', 'South'], ['B5', 'Widget'], ['C5', '50'],
  ];
  for (const [cell, val] of rows) {
    await go(page, cell, val);
  }
}

async function openCreatePivotDialog(page: Page) {
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'PivotTable', exact: true }).click();
  await page.getByRole('menuitem', { name: 'From Table/Range', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'PivotTable from table or range' });
  await expect(dialog).toBeVisible();
  return dialog;
}

test('creates, configures, formats and places PivotTables via Ribbon, Dialog and Fields Pane', async ({ page }) => {
  await page.context().addInitScript(axe.source);
  await page.goto(path);
  await seedSalesData(page);
  await go(page, 'A1');

  const createDialog = await openCreatePivotDialog(page);
  await expect(createDialog.getByLabel('Table/Range:')).toHaveValue(/Sheet1'?!\$A\$1:\$C\$5/u);
  await createDialog.getByRole('radio', { name: 'Existing Worksheet' }).check();
  await createDialog.getByLabel('Location:').fill("='Sheet1'!$E$1");
  await createDialog.getByRole('button', { name: 'OK' }).click();

  const pane = page.getByRole('complementary', { name: 'PivotTable Fields' });
  await expect(pane).toBeVisible();

  await pane.getByRole('checkbox', { name: 'Region' }).check();
  await pane.getByRole('checkbox', { name: 'Sales' }).check();

  await expect(page.getByRole('gridcell', { name: 'E1, Row Labels' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'F1, Sum of Sales' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'E2, North' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'F2, 250' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'E3, South' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'F3, 250' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'E4, Grand Total' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'F4, 500' })).toBeAttached();

  // Move Product to Column Labels via chip action menu.
  await pane.getByRole('checkbox', { name: 'Product' }).check();
  await pane.getByRole('button', { name: 'Product options' }).click();
  await pane.getByRole('menuitem', { name: 'Move to Column Labels' }).click();

  await expect(page.getByRole('gridcell', { name: 'F1, Column Labels' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'F2, Gadget' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'G2, Widget' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'H2, Grand Total' })).toBeAttached();

  // Open Value Field Settings and verify aggregation + Show Values As + number format.
  await pane.getByRole('button', { name: 'Sum of Sales options' }).click();
  await pane.getByRole('menuitem', { name: 'Value Field Settings...' }).click();
  const settingsDialog = page.getByRole('dialog', { name: 'Value Field Settings' });
  await expect(settingsDialog).toBeVisible();
  await settingsDialog.getByRole('option', { name: 'Average' }).click();
  await settingsDialog.getByRole('tab', { name: 'Show Values As' }).click();
  await settingsDialog.getByLabel('Show values as').selectOption('percentOfColumnTotal');
  await settingsDialog.getByLabel('Number Format:').fill('0.00%');
  await settingsDialog.getByRole('button', { name: 'OK' }).click();

  await expect(page.getByRole('gridcell', { name: 'H5, 100.00%' })).toBeAttached();

  // Defer Layout Update holds grid refresh until Update is clicked.
  await pane.getByRole('checkbox', { name: 'Defer Layout Update' }).check();
  await pane.getByRole('checkbox', { name: 'Product' }).uncheck();
  await expect(page.getByRole('gridcell', { name: 'F1, Column Labels' })).toBeAttached();
  await pane.getByRole('button', { name: 'Update' }).click();
  await expect(page.getByRole('gridcell', { name: 'F1, Column Labels' })).toHaveCount(0);

  // Axe accessibility check while PivotTable Fields pane is active.
  const violations = await page.evaluate(async () => (await (window as Window & { axe?: typeof axe }).axe!.run({
    runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] },
  })).violations.map(({ id, nodes }) => ({ id, targets: nodes.map(node => node.target) })));
  expect(violations).toEqual([]);

  // Close and auto-reopen when leaving and returning to the PivotTable.
  await pane.getByRole('button', { name: 'Close PivotTable Fields' }).click();
  await expect(pane).toBeHidden();
  await go(page, 'A1');
  await expect(pane).toBeHidden();
  await go(page, 'E1');
  await expect(pane).toBeVisible();

  // Creating another PivotTable on a New Worksheet activates the new sheet and opens the pane.
  await go(page, 'A1');
  const secondDialog = await openCreatePivotDialog(page);
  await expect(secondDialog.getByRole('radio', { name: 'New Worksheet' })).toBeChecked();
  await secondDialog.getByRole('button', { name: 'OK' }).click();
  await expect(page.getByRole('button', { name: 'Sheet2' })).toHaveAttribute('aria-current', 'true');
  await expect(page.getByRole('grid', { name: 'Sheet2' })).toBeVisible();
  await expect(pane).toBeVisible();
});

test('supports expand/collapse, in-sheet report/label filter & sort dropdowns, and double-click drill-down', async ({ page }) => {
  await page.context().addInitScript(axe.source);
  await page.goto(path);
  await seedSalesData(page);
  await go(page, 'A1');

  const createDialog = await openCreatePivotDialog(page);
  await createDialog.getByRole('radio', { name: 'Existing Worksheet' }).check();
  await createDialog.getByLabel('Location:').fill("='Sheet1'!$E$1");
  await createDialog.getByRole('button', { name: 'OK' }).click();

  const pane = page.getByRole('complementary', { name: 'PivotTable Fields' });
  await pane.getByRole('checkbox', { name: 'Region' }).check();
  await pane.getByRole('checkbox', { name: 'Product' }).check();
  await pane.getByRole('checkbox', { name: 'Sales' }).check();

  // Multi-level rows render +/- expand/collapse buttons in Compact layout.
  const collapseNorth = page.getByRole('button', { name: 'Collapse North' });
  await expect(collapseNorth).toHaveAttribute('aria-expanded', 'true');
  await collapseNorth.click();
  const expandNorth = page.getByRole('button', { name: 'Expand North' });
  await expect(expandNorth).toHaveAttribute('aria-expanded', 'false');
  await expandNorth.click();
  await expect(page.getByRole('button', { name: 'Collapse North' })).toHaveAttribute('aria-expanded', 'true');

  // Move Product to Report Filters and test single-select & Clear Filter.
  await pane.getByRole('button', { name: 'Product options' }).click();
  await pane.getByRole('menuitem', { name: 'Move to Report Filter' }).click();
  await expect(page.getByRole('gridcell', { name: 'E1, Product' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'F1, (All)' })).toBeAttached();

  await page.getByRole('button', { name: 'Pivot filter Product' }).click();
  const productMenu = page.getByRole('dialog', { name: 'Filter Product' });
  await productMenu.getByRole('option', { name: 'Gadget' }).click();
  await productMenu.getByRole('button', { name: 'OK' }).click();
  await expect(page.getByRole('gridcell', { name: 'F1, Gadget' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'F6, 350' })).toBeAttached();

  await page.getByRole('button', { name: 'Pivot filter Product' }).click();
  await page.getByRole('button', { name: 'Clear Filter From "Product"' }).click();
  await expect(page.getByRole('gridcell', { name: 'F6, 500' })).toBeAttached();

  // Row Labels in-sheet sort, filter checklist, axe check, and More Sort Options.
  await page.getByRole('button', { name: 'Pivot filter Row Labels' }).click();
  await page.getByRole('menuitem', { name: 'Sort Z to A' }).or(page.getByRole('button', { name: 'Sort Z to A' })).click();
  await expect(page.getByRole('gridcell', { name: 'E4, South' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'E5, North' })).toBeAttached();

  await page.getByRole('button', { name: 'Pivot filter Row Labels' }).click();
  const rowMenu = page.getByRole('dialog', { name: /Filter Row Labels/ });
  await rowMenu.getByRole('checkbox', { name: 'South' }).uncheck();
  const violations = await page.evaluate(async () => (await (window as Window & { axe?: typeof axe }).axe!.run({
    runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] },
  })).violations.map(({ id, nodes }) => ({ id, targets: nodes.map(node => node.target) })));
  expect(violations).toEqual([]);
  await rowMenu.getByRole('button', { name: 'OK' }).click();
  await expect(page.getByRole('gridcell', { name: 'E4, North' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'F5, 250' })).toBeAttached();

  await page.getByRole('button', { name: 'Pivot filter Row Labels' }).click();
  await page.getByRole('button', { name: 'Clear Filter From "Region"' }).click();
  await expect(page.getByRole('gridcell', { name: 'F6, 500' })).toBeAttached();

  await page.getByRole('button', { name: 'Pivot filter Row Labels' }).click();
  await page.getByRole('button', { name: 'More Sort Options...' }).click();
  await expect(page.getByRole('dialog', { name: 'Field Settings' })).toBeVisible();
  await page.getByRole('dialog', { name: 'Field Settings' }).getByRole('button', { name: 'Cancel' }).click();

  // Double-click F4 (South total = 250 on row 4, col 6) to trigger Show Details drill-down.
  const grid = page.getByRole('grid', { name: 'Sheet1' });
  const box = (await grid.boundingBox())!;
  const header = Number(await grid.getAttribute('data-row-header'));
  await page.mouse.dblclick(box.x + header + 5 * 96 + 48, box.y + 26 + 3 * 26 + 13);

  await expect(page.getByRole('button', { name: 'Sheet2' })).toHaveAttribute('aria-current', 'true');
  await expect(page.getByRole('grid', { name: 'Sheet2' })).toBeVisible();
  await expect(page.getByRole('gridcell', { name: 'A1, Region' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'A2, South' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'C2, 200' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'C3, 50' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: /North/ })).toHaveCount(0);

  // Undo restores original single-sheet workbook.
  await page.keyboard.press('Control+z');
  await expect(page.getByRole('button', { name: 'Sheet2' })).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Sheet1' })).toHaveAttribute('aria-current', 'true');
});

test('keeps chip menus unclipped, persists column filter switches, and inherits source number formats', async ({ page }) => {
  await page.goto(path);
  await seedSalesData(page);

  // Format source column C2:C5 as Currency ($#,##0.00) via Home ribbon before creating PivotTable.
  await go(page, 'C2');
  const grid = page.getByRole('grid', { name: 'Sheet1' });
  await grid.press('Shift+ArrowDown');
  await grid.press('Shift+ArrowDown');
  await grid.press('Shift+ArrowDown');
  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: /^Number Format(?: |$)/u }).click();
  await page.getByRole('menuitem', { name: /^Currency\b/u }).click();
  await expect(page.getByRole('gridcell', { name: 'C2, $100.00' })).toBeAttached();

  await go(page, 'A1');
  const createDialog = await openCreatePivotDialog(page);
  await createDialog.getByRole('radio', { name: 'Existing Worksheet' }).check();
  await createDialog.getByLabel('Location:').fill("='Sheet1'!$E$1");
  await createDialog.getByRole('button', { name: 'OK' }).click();

  const pane = page.getByRole('complementary', { name: 'PivotTable Fields' });
  await pane.getByRole('checkbox', { name: 'Product' }).check();
  await pane.getByRole('checkbox', { name: 'Region' }).check();
  await pane.getByRole('checkbox', { name: 'Sales' }).check();

  // Move Region to Column Labels and verify left-column chip menu in Rows is never clipped.
  await pane.getByRole('button', { name: 'Region options' }).click();
  await pane.getByRole('menuitem', { name: 'Move to Column Labels' }).click();

  await pane.getByRole('button', { name: 'Product options' }).click();
  const productMenu = page.getByRole('menu', { name: 'Product actions' });
  await expect(productMenu).toBeVisible();
  const menuBox = (await productMenu.boundingBox())!;
  const vp = page.viewportSize()!;
  expect(menuBox.x).toBeGreaterThanOrEqual(0);
  expect(menuBox.y).toBeGreaterThanOrEqual(0);
  expect(menuBox.x + menuBox.width).toBeLessThanOrEqual(vp.width);
  expect(menuBox.y + menuBox.height).toBeLessThanOrEqual(vp.height);
  expect(menuBox.width).toBeGreaterThanOrEqual(170);
  await page.keyboard.press('Escape');
  await expect(productMenu).toBeHidden();

  // PivotTable values automatically inherit the source Currency format.
  await expect(page.getByRole('gridcell', { name: 'H5, $500.00' })).toBeAttached();

  // Open Row Labels filter, then immediately switch to Column Labels filter and hide South.
  await page.getByRole('button', { name: 'Pivot filter Row Labels' }).click();
  await expect(page.getByRole('dialog', { name: /Filter Row Labels/ })).toBeVisible();
  await page.getByRole('button', { name: 'Pivot filter Column Labels' }).click();
  const colMenu = page.getByRole('dialog', { name: 'Filter Column Labels (Region)' });
  await expect(colMenu).toBeVisible();
  await colMenu.getByRole('checkbox', { name: 'South' }).uncheck();
  await colMenu.getByRole('button', { name: 'OK' }).click();

  // South column is removed and Grand Total reflects only North ($250.00).
  await expect(page.getByRole('gridcell', { name: 'G2, Grand Total' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'G5, $250.00' })).toBeAttached();

  // Reopening Column Labels filter remembers that South is unchecked and North is checked.
  await page.getByRole('button', { name: 'Pivot filter Column Labels' }).click();
  await expect(colMenu).toBeVisible();
  await expect(colMenu.getByRole('checkbox', { name: 'South' })).not.toBeChecked();
  await expect(colMenu.getByRole('checkbox', { name: 'North' })).toBeChecked();
});

async function dragByMouse(page: Page, from: { x: number; y: number }, to: { x: number; y: number }) {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(from.x + (to.x - from.x) / 2, from.y + (to.y - from.y) / 2, { steps: 4 });
  await page.mouse.move(to.x, to.y, { steps: 4 });
  await page.mouse.up();
}

test('supports mouse drag-and-drop in PivotTable Fields and indented hierarchy without button overlap', async ({ page }) => {
  await page.goto(path);
  await seedSalesData(page);
  await go(page, 'A1');

  const createDialog = await openCreatePivotDialog(page);
  await createDialog.getByRole('radio', { name: 'Existing Worksheet' }).check();
  await createDialog.getByLabel('Location:').fill("='Sheet1'!$E$1");
  await createDialog.getByRole('button', { name: 'OK' }).click();

  const pane = page.getByRole('complementary', { name: 'PivotTable Fields' });
  await pane.getByRole('checkbox', { name: 'Region' }).check();
  await pane.getByRole('checkbox', { name: 'Sales' }).check();

  // Drag Product from top field list directly into Columns area by mouse.
  const productLabel = pane.locator('.pivot-field-check', { hasText: 'Product' }).locator('span');
  const colsArea = pane.getByLabel('Columns area');
  const prodBox = (await productLabel.boundingBox())!;
  const colsBox = (await colsArea.boundingBox())!;
  await dragByMouse(
    page,
    { x: prodBox.x + prodBox.width / 2, y: prodBox.y + prodBox.height / 2 },
    { x: colsBox.x + colsBox.width / 2, y: colsBox.y + colsBox.height / 2 },
  );
  await expect(page.getByRole('gridcell', { name: 'F1, Column Labels' })).toBeAttached();
  await expect(colsArea.locator('.pivot-chip-label')).toHaveText(['Product']);

  // Drag Product chip from Columns to Rows above Region and verify multi-level hierarchy & indentation.
  const rowsArea = pane.getByLabel('Rows area');
  const prodChip = colsArea.locator('.pivot-chip-label', { hasText: 'Product' });
  const regionChip = rowsArea.locator('.pivot-chip', { hasText: 'Region' });
  const pChipBox = (await prodChip.boundingBox())!;
  const rChipBox = (await regionChip.boundingBox())!;
  await dragByMouse(
    page,
    { x: pChipBox.x + pChipBox.width / 2, y: pChipBox.y + pChipBox.height / 2 },
    { x: rChipBox.x + rChipBox.width / 2, y: rChipBox.y + 2 },
  );
  await expect(rowsArea.locator('.pivot-chip-label')).toHaveText(['Product', 'Region']);
  await expect(page.getByRole('button', { name: 'Collapse Gadget' })).toBeVisible();

  // Reorder Product below Region within Rows area by mouse.
  const prodRowChip = rowsArea.locator('.pivot-chip-label', { hasText: 'Product' });
  const regionRowChip = rowsArea.locator('.pivot-chip', { hasText: 'Region' });
  const prBox = (await prodRowChip.boundingBox())!;
  const rrBox = (await regionRowChip.boundingBox())!;
  await dragByMouse(
    page,
    { x: prBox.x + prBox.width / 2, y: prBox.y + prBox.height / 2 },
    { x: rrBox.x + rrBox.width / 2, y: rrBox.y + rrBox.height + 6 },
  );
  await expect(rowsArea.locator('.pivot-chip-label')).toHaveText(['Region', 'Product']);
  await expect(page.getByRole('button', { name: 'Collapse North' })).toBeVisible();

  // Drag Product chip outside the areas to remove it.
  const removeTarget = rowsArea.locator('.pivot-chip-label', { hasText: 'Product' });
  await removeTarget.scrollIntoViewIfNeeded();
  const removeBox = (await removeTarget.boundingBox())!;
  const paneBox = (await pane.boundingBox())!;
  await dragByMouse(
    page,
    { x: removeBox.x + removeBox.width / 2, y: removeBox.y + removeBox.height / 2 },
    { x: paneBox.x - 40, y: removeBox.y },
  );
  await expect(rowsArea.locator('.pivot-chip-label')).toHaveText(['Region']);
});
