import { expect, test, type Page } from '@playwright/test';

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

test('formula bar picks a reverse drag, replaces the picked reference, and commits to the original cell', async ({ page }) => {
  await page.goto('/ExcelLightWeb/');
  await enter(page, 'A1', '2');
  await enter(page, 'B1', '3');
  await enter(page, 'B2', '5');
  await go(page, 'D4');
  const editor = page.getByRole('textbox', { name: 'Formula bar' });
  await editor.fill('=SUM()');
  await editor.press('ArrowLeft');
  const b2 = await point(page, 2, 2), a1 = await point(page, 1, 1);
  await page.mouse.move(b2.x, b2.y);
  await page.mouse.down();
  await expect(editor).toHaveValue('=SUM(B2)');
  await page.mouse.move(a1.x, a1.y, { steps: 5 });
  await page.mouse.up();
  await expect(editor).toHaveValue('=SUM(A1:B2)');
  await expect(editor).toBeFocused();
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('D4');
  await expect(page.getByRole('gridcell', { name: 'D4, blank', exact: true })).toBeAttached();
  await page.mouse.click(b2.x, b2.y);
  await expect(editor).toHaveValue('=SUM(B2)');
  await editor.press('Enter');
  await expect(page.getByRole('gridcell', { name: 'D4, 5', exact: true })).toBeAttached();
  await expect(page.getByRole('grid')).toBeFocused();
});

test('inline picking can add a cross-sheet reference with the formula bar hidden', async ({ page }) => {
  await page.goto('/ExcelLightWeb/');
  await page.getByRole('button', { name: 'New sheet' }).click();
  await enter(page, 'A1', '7');
  await page.getByRole('button', { name: 'Sheet2', exact: true }).dblclick();
  const rename = page.getByRole('textbox', { name: 'Sheet name' });
  await rename.fill("O'Brien");
  await rename.press('Enter');
  const sourceSheet = page.getByRole('button', { name: 'Sheet1', exact: true });
  const otherSheet = page.getByRole('button', { name: "O'Brien", exact: true });
  await otherSheet.click();
  await enter(page, 'A1', '7');
  await sourceSheet.click();
  await go(page, 'C3');
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('checkbox', { name: 'Formula Bar' }).uncheck();
  const grid = page.getByRole('grid');
  await grid.press('F2');
  const inline = page.getByRole('textbox', { name: 'Edit C3' });
  await inline.fill('=SUM(');
  const local = await point(page, 1, 2);
  await page.mouse.click(local.x, local.y);
  await expect(inline).toHaveValue('=SUM(B1');
  await expect(inline).toBeFocused();
  await page.keyboard.insertText(',');
  await otherSheet.click();
  await expect(otherSheet).toHaveAttribute('aria-current', 'true');
  const editor = page.getByRole('textbox', { name: 'Formula bar' });
  await expect(editor).toBeFocused();
  const remote = await point(page, 1, 1);
  await page.mouse.click(remote.x, remote.y);
  await expect(editor).toHaveValue("=SUM(B1,'O''Brien'!A1");
  await editor.press('Escape');
  await expect(sourceSheet).toHaveAttribute('aria-current', 'true');
  await expect(editor).toHaveCount(0);
  await expect(page.getByRole('gridcell', { name: 'C3, blank', exact: true })).toBeAttached();

  await grid.press('F2');
  await inline.fill('=');
  await otherSheet.click();
  await expect(editor).toBeFocused();
  const remoteAgain = await point(page, 1, 1);
  await page.mouse.click(remoteAgain.x, remoteAgain.y);
  await expect(editor).toHaveValue("='O''Brien'!A1");
  await editor.press('Enter');
  await expect(sourceSheet).toHaveAttribute('aria-current', 'true');
  await expect(page.getByRole('gridcell', { name: 'C3, 7', exact: true })).toBeAttached();
  await expect(editor).toHaveCount(0);
});
