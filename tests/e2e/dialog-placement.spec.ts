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

test('ribbon combo boxes (AutoSum, Merge & Center, Font, Font size, Underline, Borders, Fill Color, Font Color) expand aligned with their control', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 900 });
  await page.goto('/ExcelLightWeb/');

  const cases: {
    name: string;
    comboSelector: string;
    triggerRole: 'button';
    triggerName: string | RegExp;
    menuSelector: string;
  }[] = [
    {
      name: 'Font family',
      comboSelector: '.ribbon-font-input:has(input[aria-label="Font family"])',
      triggerRole: 'button',
      triggerName: 'Font family options',
      menuSelector: '.ribbon-flyout[role="menu"][aria-label="Font family options"]',
    },
    {
      name: 'Font size',
      comboSelector: '.ribbon-font-input:has(input[aria-label="Font size"])',
      triggerRole: 'button',
      triggerName: 'Font size options',
      menuSelector: '.ribbon-flyout[role="menu"][aria-label="Font size options"]',
    },
    {
      name: 'Underline',
      comboSelector: '.ribbon-split:has(button[aria-label="Underline"])',
      triggerRole: 'button',
      triggerName: 'Underline options',
      menuSelector: '.ribbon-flyout[role="menu"][aria-label="Underline options"]',
    },
    {
      name: 'Borders',
      comboSelector: '.border-ribbon.ribbon-split',
      triggerRole: 'button',
      triggerName: 'Borders menu',
      menuSelector: '.border-menu[aria-label="Borders actions"]',
    },
    {
      name: 'Fill Color',
      comboSelector: '.ribbon-split:has(button[aria-label="Apply last fill color"])',
      triggerRole: 'button',
      triggerName: 'Fill Color',
      menuSelector: '.ribbon-palette[aria-label="Fill Color"]',
    },
    {
      name: 'Font Color',
      comboSelector: '.ribbon-split:has(button[aria-label="Apply last font color"])',
      triggerRole: 'button',
      triggerName: 'Font Color',
      menuSelector: '.ribbon-palette[aria-label="Font Color"]',
    },
    {
      name: 'Merge & Center',
      comboSelector: '.ribbon-split:has(button:has-text("Merge & Center"))',
      triggerRole: 'button',
      triggerName: 'Merge options',
      menuSelector: '.ribbon-flyout[role="menu"][aria-label="Merge options"]',
    },
    {
      name: 'AutoSum',
      comboSelector: '.ribbon-split:has(button:text-is("AutoSum"))',
      triggerRole: 'button',
      triggerName: 'AutoSum options',
      menuSelector: '.ribbon-flyout[role="menu"][aria-label="AutoSum options"]',
    },
  ];

  for (const item of cases) {
    const combo = page.locator(item.comboSelector).first();
    const trigger = page.getByRole(item.triggerRole, { name: item.triggerName, exact: true }).first();
    await trigger.click();
    const menu = page.locator(item.menuSelector);
    await expect(menu).toBeVisible();

    const comboRect = (await combo.boundingBox())!;
    const menuRect = (await menu.boundingBox())!;
    expect(Math.abs(menuRect.x - comboRect.x), `${item.name} aligned with combo box`).toBeLessThanOrEqual(1.5);
    expect(menuRect.y, `${item.name} opens below combo box`).toBeGreaterThanOrEqual(comboRect.y + comboRect.height - 1);

    await page.keyboard.press('Escape');
    await expect(menu).not.toBeVisible();
  }
});
