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

test('uses the desktop table dialog, ribbon and contextual tab memory', async ({ page }) => {
  await page.goto(path);
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  const tableButton = page.getByRole('button', { name: 'Table', exact: true });
  await expect(tableButton).toHaveClass(/ribbon-large/);
  await tableButton.click();
  const create = page.getByRole('dialog', { name: 'Create Table' });
  await expect(create.getByRole('heading', { name: 'Create Table' })).toBeVisible();
  await expect(create.getByRole('textbox', { name: 'Name', exact: true })).toHaveValue('Table_1');
  await expect(create.getByText('Workbook', { exact: true })).toBeVisible();
  await expect(create.getByRole('textbox', { name: 'Comment' })).toBeVisible();
  await expect(create.getByRole('textbox', { name: 'Range', exact: true })).toHaveValue("='Sheet1'!$A$1");
  await expect(create.getByRole('checkbox', { name: 'My table has headers' })).not.toBeChecked();
  await expect(create.locator('.dialog-actions button')).toHaveText(['OK', 'Cancel']);
  const rangeButton = create.getByRole('button', { name: 'Select range from worksheet' });
  await expect(rangeButton).toHaveText('...');
  await rangeButton.click();
  const picker = page.locator('.range-picker-floating-panel');
  await expect(picker.getByText('Select a worksheet range')).toBeVisible();
  await expect(create).toBeHidden();
  await go(page, 'B2');
  await expect(picker.getByRole('textbox', { name: 'Selected worksheet range' })).toHaveValue("='Sheet1'!$B$2");
  await page.getByRole('grid').press('Escape');
  await expect(picker).toBeHidden();
  await expect(create.getByRole('textbox', { name: 'Range', exact: true })).toHaveValue("='Sheet1'!$A$1");
  await create.getByRole('button', { name: 'OK', exact: true }).click();
  await go(page, 'A1');

  const design = page.getByRole('tab', { name: 'Table Design', exact: true });
  await expect(design).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByRole('textbox', { name: 'Table Name' })).toHaveValue('Table_1');
  await expect(page.getByRole('button', { name: 'Resize Table', exact: true })).toBeVisible();
  await expect(page.getByRole('checkbox', { name: 'Header Row' })).toBeChecked();
  await expect(page.getByRole('checkbox', { name: 'Filter Button' })).toBeEnabled();
  await expect(page.getByRole('checkbox', { name: 'Banded Rows' })).toBeChecked();
  await expect(page.getByRole('checkbox', { name: 'Banded Columns' })).not.toBeChecked();
  for (const style of ['Blue', 'Green', 'Orange', 'Purple']) {
    await expect(page.getByRole('button', { name: style + ' table style' })).toBeVisible();
  }
  await expect(page.getByRole('button', { name: 'Blue table style' })).toHaveClass(/selected/);
  await page.getByRole('checkbox', { name: 'Header Row' }).click();
  await go(page, 'A2');
  await expect(page.getByRole('checkbox', { name: 'Header Row' })).not.toBeChecked();
  await expect(page.getByRole('checkbox', { name: 'Filter Button' })).toBeDisabled();
  await page.getByRole('checkbox', { name: 'Header Row' }).click();
  await go(page, 'A1');
  await expect(page.getByRole('checkbox', { name: 'Header Row' })).toBeChecked();

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await go(page, 'C3');
  await expect(page.getByRole('tab', { name: 'Table Design' })).toHaveCount(0);
  await go(page, 'A1');
  await expect(page.getByRole('tab', { name: 'Home', exact: true })).toHaveAttribute('aria-selected', 'true');
  await design.click();
  await go(page, 'C3');
  await go(page, 'A1');
  await expect(design).toHaveAttribute('aria-selected', 'true');

  const handle = page.getByRole('button', { name: 'Resize table Table_1' });
  const handleBox = await handle.boundingBox();
  if (!handleBox) throw new Error('The table resize handle is not visible.');
  await page.mouse.move(handleBox.x + handleBox.width / 2, handleBox.y + handleBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(handleBox.x + 100, handleBox.y + handleBox.height / 2);
  await expect(page.getByRole('gridcell', { name: 'B1, blank' })).toBeAttached();
  const dashed = await page.locator('.worksheet-canvas').evaluate((element: HTMLCanvasElement) => {
    const ratio = window.devicePixelRatio || 1;
    const data = element.getContext('2d')!.getImageData(0, Math.round(76 * ratio), element.width, 4).data;
    let accent = 0;
    for (let index = 0; index < data.length; index += 4) {
      const red = data[index] ?? 0;
      const green = data[index + 1] ?? 0;
      const blue = data[index + 2] ?? 0;
      const alpha = data[index + 3] ?? 0;
      if (alpha > 200 && red > 70 && red < 130 && green > 180 && blue > 230) accent++;
    }
    return accent;
  });
  expect(dashed).toBeGreaterThan(8);
  await page.mouse.up();
  await expect(page.getByRole('gridcell', { name: 'B1, Column2' })).toBeAttached();
  await page.getByRole('checkbox', { name: 'Total Row' }).check();
  await go(page, 'A3');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('Total');
  await go(page, 'B3');
  await expect(page.getByRole('button', { name: 'Table total function options' })).toBeVisible();
  await page.getByRole('button', { name: 'Table total function options' }).click();
  await expect(page.getByRole('menuitemcheckbox', { name: 'Sum' })).toHaveAttribute('aria-checked', 'true');
  await page.getByRole('menuitemcheckbox', { name: 'Average' }).click();
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('=SUBTOTAL(101,B2:B2)');
});

test('uses the desktop New Name and Name Manager dialogs', async ({ page }) => {
  await page.goto(path);
  await go(page, 'A1', '5');
  await page.getByRole('tab', { name: 'Formulas', exact: true }).click();
  const define = page.getByRole('button', { name: 'Define Name', exact: true });
  await expect(define).not.toHaveClass(/ribbon-large/);
  await expect(page.getByRole('button', { name: 'Name Manager', exact: true })).toHaveClass(/ribbon-large/);
  await define.click();
  await page.getByRole('menuitem', { name: 'Define Name...' }).click();
  const named = page.getByRole('dialog', { name: 'New Name' });
  await expect(named.getByText('Workbook', { exact: true })).toBeVisible();
  await expect(named.locator('.dialog-actions button')).toHaveText(['OK', 'Cancel']);
  await named.getByRole('textbox', { name: 'Name', exact: true }).fill('A1');
  await named.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(named.getByRole('alert')).toContainText('Invalid workbook name');
  await named.getByRole('textbox', { name: 'Name', exact: true }).fill('TaxRate');
  await named.getByRole('textbox', { name: 'Comment' }).fill('Local rate');
  await named.getByRole('textbox', { name: 'Refers to' }).fill("='Sheet1'!$A$1");
  await named.getByRole('button', { name: 'OK', exact: true }).click();

  await go(page, 'C3');
  await page.getByRole('button', { name: 'Name Manager', exact: true }).click();
  const manager = page.getByRole('dialog', { name: 'Name Manager' });
  await expect(manager.getByRole('combobox', { name: 'Filter names' })).toHaveValue('0');
  await expect(manager.locator('.dialog-actions button')).toHaveText(['Close']);
  await expect(manager.getByRole('button', { name: 'Select range from worksheet' })).toHaveText('...');
  await manager.getByRole('button', { name: 'Select range from worksheet' }).click();
  await expect(manager).toBeHidden();
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('A1');
  await expect(page.getByRole('gridcell', { name: 'A1, 5' })).toHaveAttribute('aria-selected', 'true');
  await page.getByRole('grid').press('Escape');
  await expect(manager).toBeVisible();
  const row = manager.getByRole('row', { name: /TaxRate, named range/ });
  await expect(row).toContainText('Workbook');
  await expect(row).toContainText('Local rate');
  await manager.getByRole('button', { name: 'New...' }).click();
  const another = page.getByRole('dialog', { name: 'New Name' });
  await expect(manager).toBeHidden();
  await another.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(manager).toBeVisible();
  await row.click();
  await manager.getByRole('button', { name: 'Edit...' }).click();
  const edit = page.getByRole('dialog', { name: 'Edit Name' });
  await expect(edit.getByRole('textbox', { name: 'Name', exact: true })).toHaveValue('TaxRate');
  await expect(edit.getByRole('textbox', { name: 'Refers to' })).toBeEnabled();
  await expect(edit.getByRole('textbox', { name: 'Comment' })).toBeEnabled();
  await edit.getByRole('button', { name: 'Cancel', exact: true }).click();
  await manager.getByRole('button', { name: 'Delete', exact: true }).click();
  await expect(manager.getByRole('row', { name: /TaxRate/ })).toHaveCount(0);
  await manager.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(page.getByRole('grid')).toBeFocused();
});

test('edits a table name and comment without changing its size, and still edits a named range', async ({ page }) => {
  await page.goto(path);
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Table', exact: true }).click();
  await page.getByRole('dialog', { name: 'Create Table' }).getByRole('button', { name: 'OK', exact: true }).click();
  await go(page, 'E1', '5');
  await page.getByRole('tab', { name: 'Formulas', exact: true }).click();
  await page.getByRole('button', { name: 'Define Name', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Define Name...' }).click();
  const named = page.getByRole('dialog', { name: 'New Name' });
  await named.getByRole('textbox', { name: 'Name', exact: true }).fill('TaxRate');
  await named.getByRole('textbox', { name: 'Refers to' }).fill("='Sheet1'!$E$1");
  await named.getByRole('button', { name: 'OK', exact: true }).click();

  await page.getByRole('button', { name: 'Name Manager', exact: true }).click();
  const manager = page.getByRole('dialog', { name: 'Name Manager' });
  const tableRow = manager.getByRole('row', { name: /Table_1, table/ });
  await tableRow.click();
  const refers = manager.getByRole('textbox', { name: 'Refers to' });
  await expect(refers).toBeDisabled();
  await expect(refers).toHaveValue("='Sheet1'!$A$1:$A$2");
  await expect(manager.getByRole('button', { name: 'Apply refers to' })).toBeDisabled();
  await expect(manager.getByRole('button', { name: 'Cancel refers to edit' })).toBeDisabled();
  await expect(manager.getByRole('button', { name: 'Select range from worksheet' })).toBeDisabled();

  await manager.getByRole('button', { name: 'Edit...' }).click();
  const editTable = page.getByRole('dialog', { name: 'Edit Table' });
  await expect(editTable.getByRole('textbox', { name: 'Name', exact: true })).toBeEnabled();
  await expect(editTable.getByRole('textbox', { name: 'Comment' })).toBeEnabled();
  await expect(editTable.getByRole('textbox', { name: 'Range', exact: true })).toBeDisabled();
  await expect(editTable.getByRole('button', { name: 'Select range from worksheet' })).toBeDisabled();
  await editTable.getByRole('textbox', { name: 'Name', exact: true }).fill('Orders');
  await editTable.getByRole('textbox', { name: 'Comment' }).fill('Kept');
  await editTable.getByRole('button', { name: 'OK', exact: true }).click();

  const orders = manager.getByRole('row', { name: /Orders, table/ });
  await expect(orders).toContainText('Kept');
  await expect(orders).toContainText("='Sheet1'!$A$1:$A$2");
  await expect(manager.getByRole('textbox', { name: 'Refers to' })).toBeDisabled();
  await expect(manager.getByRole('textbox', { name: 'Refers to' })).toHaveValue("='Sheet1'!$A$1:$A$2");

  await manager.getByRole('row', { name: /TaxRate, named range/ }).click();
  await expect(manager.getByRole('textbox', { name: 'Refers to' })).toBeEnabled();
  await manager.getByRole('textbox', { name: 'Refers to' }).fill("='Sheet1'!$F$1");
  await manager.getByRole('button', { name: 'Apply refers to' }).click();
  const rate = manager.getByRole('row', { name: /TaxRate, named range/ });
  await expect(rate).toContainText("='Sheet1'!$F$1");

  await manager.getByRole('button', { name: 'Edit...' }).click();
  const editName = page.getByRole('dialog', { name: 'Edit Name' });
  await expect(editName.getByRole('textbox', { name: 'Refers to' })).toBeEnabled();
  await editName.getByRole('textbox', { name: 'Name', exact: true }).fill('Rate');
  await editName.getByRole('textbox', { name: 'Comment' }).fill('Updated');
  await editName.getByRole('textbox', { name: 'Refers to' }).fill("='Sheet1'!$E$1");
  await editName.getByRole('button', { name: 'OK', exact: true }).click();
  const renamed = manager.getByRole('row', { name: /Rate, named range/ });
  await expect(renamed).toContainText('Updated');
  await expect(renamed).toContainText("='Sheet1'!$E$1");
  await expect(manager.getByRole('textbox', { name: 'Refers to' })).toBeEnabled();
  await expect(manager.getByRole('textbox', { name: 'Refers to' })).toHaveValue("='Sheet1'!$E$1");
});

test('keeps the grid still when Table Design closes and shrinks a table that has a total row', async ({ page }) => {
  await page.goto(path);
  await go(page, 'A1', 'Kind');
  await go(page, 'B1', 'Amount');
  await go(page, 'A2', 'a');
  await go(page, 'B2', '10');
  await go(page, 'A3', 'b');
  await go(page, 'B3', '4');
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Table', exact: true }).click();
  const create = page.getByRole('dialog', { name: 'Create Table' });
  await create.getByRole('textbox', { name: 'Range', exact: true }).fill("='Sheet1'!$A$1:$B$3");
  await create.getByRole('checkbox', { name: 'My table has headers' }).check();
  await create.getByRole('button', { name: 'OK', exact: true }).click();
  await go(page, 'A1');
  await expect(page.getByRole('tab', { name: 'Table Design', exact: true })).toHaveAttribute('aria-selected', 'true');

  await page.evaluate(() => {
    const ribbon = document.querySelector('.ribbon');
    const scroller = document.querySelector('.worksheet-scroll');
    if (!(ribbon instanceof HTMLElement) || !(scroller instanceof HTMLElement)) throw new Error('Worksheet layout is missing.');
    const samples: { ribbon: number; top: number; scroll: number }[] = [];
    let frame = 0;
    const tick = () => {
      samples.push({
        ribbon: ribbon.getBoundingClientRect().height,
        top: scroller.getBoundingClientRect().top,
        scroll: scroller.scrollTop,
      });
      frame = requestAnimationFrame(tick);
    };
    frame = requestAnimationFrame(tick);
    Object.assign(window, {
      __stopLayoutWatch: () => {
        cancelAnimationFrame(frame);
        return samples;
      },
    });
  });
  await go(page, 'D5');
  const samples = await page.evaluate(() =>
    (window as unknown as { __stopLayoutWatch: () => { ribbon: number; top: number; scroll: number }[] }).__stopLayoutWatch());
  expect(samples.length).toBeGreaterThan(2);
  const span = (values: number[]) => Math.max(...values) - Math.min(...values);
  expect(span(samples.map(sample => sample.ribbon))).toBeLessThan(2);
  expect(span(samples.map(sample => sample.top))).toBeLessThan(2);
  expect(span(samples.map(sample => sample.scroll))).toBe(0);
  await expect(page.getByRole('tab', { name: 'Table Design' })).toHaveCount(0);

  await go(page, 'A1');
  await page.getByRole('checkbox', { name: 'Total Row' }).check();
  await expect(page.getByRole('gridcell', { name: 'A4, Total' })).toBeAttached();
  const handle = page.getByRole('button', { name: 'Resize table Table_1' });
  const handleBox = await handle.boundingBox();
  if (!handleBox) throw new Error('The table resize handle is not visible.');
  const originX = handleBox.x + handleBox.width / 2;
  const originY = handleBox.y + handleBox.height / 2;
  await page.mouse.move(originX, originY);
  await page.mouse.down();
  await page.mouse.move(originX, originY - 36);
  await page.mouse.up();
  await expect(page.getByRole('gridcell', { name: 'A2, a' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'A3, Total' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'A4, b' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B4, 4' })).toBeAttached();
  await go(page, 'B3');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('=SUBTOTAL(109,B2:B2)');

  const grown = await page.getByRole('button', { name: 'Resize table Table_1' }).boundingBox();
  if (!grown) throw new Error('The table resize handle is not visible.');
  const growX = grown.x + grown.width / 2;
  const growY = grown.y + grown.height / 2;
  await page.mouse.move(growX, growY);
  await page.mouse.down();
  await page.mouse.move(growX, growY + 30);
  await page.mouse.up();
  await expect(page.getByRole('gridcell', { name: 'A3, b' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B3, 4' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'A4, Total' })).toBeAttached();
  await go(page, 'B4');
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('=SUBTOTAL(109,B2:B3)');
});
