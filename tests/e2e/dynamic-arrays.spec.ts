import { expect, test, type Page } from '@playwright/test';

const path = '/ExcelLightWeb/';

async function go(page: Page, address: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(address);
  await name.press('Enter');
}

async function enter(page: Page, address: string, source: string) {
  await go(page, address);
  const editor = page.getByRole('textbox', { name: 'Formula bar' });
  await editor.fill(source);
  await editor.press('Enter');
}

function cell(page: Page, address: string, value: string) {
  return page.getByRole('gridcell', { name: address + ', ' + value, exact: true });
}

test('spills dynamic array results, renders spill border, enforces read-only children, and recovers from #SPILL!', async ({ page }) => {
  await page.goto(path);
  await enter(page, 'A1', '=MAKEARRAY(2,2,LAMBDA(r,c,r*10+c))');

  await expect(cell(page, 'A1', '11')).toBeAttached();
  await expect(cell(page, 'B1', '12')).toBeAttached();
  await expect(cell(page, 'A2', '21')).toBeAttached();
  await expect(cell(page, 'B2', '22')).toBeAttached();

  const grid = page.getByRole('grid');
  await expect(grid).toHaveAttribute('data-spill-range', 'A1:B2');

  await go(page, 'B2');
  await expect(grid).toHaveAttribute('data-spill-range', 'A1:B2');
  await expect(cell(page, 'B2', '22')).toHaveAttribute('aria-readonly', 'true');

  // Attempting to overwrite a spilled cell keeps its generated value and shows the read-only error.
  const editor = page.getByRole('textbox', { name: 'Formula bar' });
  await editor.fill('99');
  await editor.press('Enter');
  await expect(page.getByRole('status').filter({ hasText: /read-only/i })).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(cell(page, 'B2', '22')).toBeAttached();

  // Downstream spill-ref consumer A1# recalculates.
  await enter(page, 'D1', '=SUM(A1#)');
  await expect(cell(page, 'D1', '66')).toBeAttached();

  // Collision with a blocker produces #SPILL!, then clearing the blocker recovers automatically.
  await enter(page, 'E3', 'blocker');
  await enter(page, 'E1', '=MAKEARRAY(3,1,LAMBDA(r,c,r))');
  await expect(cell(page, 'E1', '#SPILL!')).toBeAttached();

  await go(page, 'E3');
  await page.keyboard.press('Delete');
  await expect(cell(page, 'E1', '1')).toBeAttached();
  await expect(cell(page, 'E2', '2')).toBeAttached();
  await expect(cell(page, 'E3', '3')).toBeAttached();
});
