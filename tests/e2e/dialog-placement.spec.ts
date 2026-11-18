import { expect, test, type Locator, type Page } from '@playwright/test';

async function moveDialog(page: Page, dialog: Locator, dx: number, dy: number) {
  const before = (await dialog.boundingBox())!;
  const bar = (await dialog.locator('.format-dialog-heading').boundingBox())!;
  const x = bar.x + 60, y = bar.y + bar.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + dx, y + dy, { steps: 6 });
  await page.mouse.up();
  const after = (await dialog.boundingBox())!;
  expect(after.x - before.x).toBeCloseTo(dx, 0);
  expect(after.y - before.y).toBeCloseTo(dy, 0);
}

test('drags the main and nested dialogs independently from their title bars', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 900 });
  await page.goto('/ExcelLightWeb/');
  await page.getByRole('grid').press('Control+1');
  const format = page.getByRole('dialog', { name: 'Format Cells' });
  await moveDialog(page, format, 70, 50);
  const position = await format.boundingBox();
  await format.getByRole('tab', { name: 'Fill', exact: true }).click();
  await format.getByRole('button', { name: 'Fill Effects…' }).click();
  const effects = page.getByRole('dialog', { name: 'Fill Effects' });
  await moveDialog(page, effects, -70, -40);
  await effects.getByRole('button', { name: 'Color 1', exact: true }).click();
  const picker = page.getByRole('dialog', { name: 'Color 1', exact: true });
  await moveDialog(page, picker, 40, 30);
  await picker.getByRole('button', { name: 'Cancel', exact: true }).click();
  await effects.getByRole('button', { name: 'Close Fill Effects' }).click();
  expect(await format.boundingBox()).toEqual(position);
  await format.getByRole('button', { name: 'Close Format Cells' }).click();
  await expect(page.getByRole('grid')).toBeFocused();
});

test('shows the full Borders menu when it fits and scrolls it only in a shorter viewport', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 1000 });
  await page.goto('/ExcelLightWeb/');
  await page.getByRole('button', { name: 'Borders menu' }).click();
  const menu = page.getByRole('group', { name: 'Borders actions' });
  const overflow = () => menu.evaluate(node => node.scrollHeight > node.clientHeight);
  await expect.poll(overflow).toBe(false);
  await expect(menu.getByRole('button', { name: 'More Borders…' })).toBeInViewport();
  await page.setViewportSize({ width: 1280, height: 520 });
  await expect.poll(overflow).toBe(true);
  await menu.getByRole('button', { name: 'More Borders…' }).click();
  await expect(page.getByRole('tab', { name: 'Border', selected: true })).toBeVisible();
});
