import axe from 'axe-core';
import { expect, test, type Page } from '@playwright/test';

const path = '/ExcelLightWeb/';

async function expectActive(page: Page, address: string) {
  const grid = page.getByRole('grid');
  await expect(grid).toHaveAttribute('aria-activedescendant', new RegExp('-' + address + '$'));
  const id = await grid.getAttribute('aria-activedescendant');
  await expect(grid.locator(`[id="${id}"]`)).toHaveAttribute('role', 'gridcell');
}

async function gotoCell(page: Page, address: string) {
  const box = page.getByRole('textbox', { name: 'Name box' });
  await box.fill(address);
  await box.press('Enter');
}

test('keyboard extends ranges, reveals offscreen cells, and keeps keyboard focus escapable', async ({ page }) => {
  await page.goto(path);
  const grid = page.getByRole('grid');
  const name = page.getByRole('textbox', { name: 'Name box' });
  await grid.focus();
  await grid.press('ArrowDown');
  await grid.press('ArrowRight');
  await grid.press('Shift+ArrowDown');
  await grid.press('Shift+ArrowRight');
  await expect(name).toHaveValue('B2:C3');
  await expectActive(page, 'B2');
  await expect(page.getByRole('gridcell', { name: 'C3, blank' })).toHaveAttribute('aria-selected', 'true');
  await grid.press('Escape');
  await expect(name).toHaveValue('B2');
  await grid.press('Control+ArrowDown');
  await expectActive(page, 'B1000');
  await expect.poll(() => grid.evaluate(node => node.scrollTop)).toBeGreaterThan(20_000);
  await grid.press('Control+Home');
  await expectActive(page, 'A1');
  await grid.press('PageDown');
  await expect.poll(() => grid.evaluate(node => node.scrollTop)).toBeGreaterThan(0);
  await grid.press('Shift+Space');
  await expect(name).toHaveValue(/^A\d+:Z\d+$/);
  await grid.press('Control+Space');
  await expect(name).toHaveValue('A1:A1000');
  await grid.press('Control+a');
  await expect(name).toHaveValue('A1:Z1000');
  await grid.press('Tab');
  await expectActive(page, 'B1');
  await grid.press('Shift+Tab');
  await expectActive(page, 'A1');
  await grid.press('Escape');
  await grid.press('Tab');
  await expect(page.getByRole('button', { name: 'Sheet1' })).toBeFocused();
});

test('pointer drag selects a rectangle and Control-click toggles independent areas', async ({ page }) => {
  await page.goto(path);
  const grid = page.getByRole('grid');
  const { x, y } = (await grid.boundingBox())!;
  const name = page.getByRole('textbox', { name: 'Name box' });
  await page.mouse.move(x + 75, y + 36);
  await page.mouse.down();
  await page.mouse.move(x + 176, y + 88, { steps: 8 });
  await page.mouse.up();
  await expect(name).toHaveValue('A1:B3');
  await expect(grid).toBeFocused();
  await page.keyboard.down('Control');
  await page.mouse.click(x + 365, y + 138);
  await expect(name).toHaveValue('A1:B3, D5');
  await expectActive(page, 'D5');
  await page.mouse.click(x + 75, y + 36);
  await page.keyboard.up('Control');
  await expect(name).toHaveValue('D5');
  await page.mouse.click(x + 175, y + 9);
  await expect(name).toHaveValue('B1:B1000');
  await page.mouse.click(x + 12, y + 62);
  await expect(name).toHaveValue('A2:Z2');
  await page.mouse.click(x + 12, y + 8);
  await expect(name).toHaveValue('A1:Z1000');
});

test('native scrolling virtualizes both axes and retains an offscreen active descendant', async ({ page }) => {
  await page.setViewportSize({ width: 850, height: 520 });
  await page.goto(path);
  const grid = page.getByRole('grid');
  await grid.evaluate(node => node.scrollTo({ top: 10_417, left: 1200 }));
  await expect(page.getByRole('gridcell', { name: 'N401, blank' })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B2, blank' })).toHaveCount(0);
  await expectActive(page, 'A1');
  expect(await page.getByRole('gridcell').count()).toBeLessThan(250);
  await gotoCell(page, 'z1000');
  await expect(grid).toBeFocused();
  await expectActive(page, 'Z1000');
  await expect.poll(() => grid.evaluate(node => node.scrollLeft)).toBeGreaterThan(1500);
  await expect.poll(() => grid.evaluate(node => node.scrollTop)).toBeGreaterThan(24_000);
  const canvas = page.locator('.worksheet-canvas');
  const dimensions = await grid.evaluate(node => ({ width: node.clientWidth, height: node.clientHeight }));
  await expect(canvas).toHaveCSS('width', dimensions.width + 'px');
  await expect(canvas).toHaveCSS('height', dimensions.height + 'px');
});

test('dragging beyond the worksheet scrolls the range and the name box rejects invalid coordinates', async ({ page }) => {
  await page.setViewportSize({ width: 880, height: 540 });
  await page.goto(path);
  const grid = page.getByRole('grid');
  const { x, y, height } = (await grid.boundingBox())!;
  await page.mouse.move(x + 70, y + 38);
  await page.mouse.down();
  await page.mouse.move(x + 70, y + height + 18, { steps: 8 });
  await expect.poll(() => grid.evaluate(node => node.scrollTop)).toBeGreaterThan(30);
  await page.mouse.up();
  const name = page.getByRole('textbox', { name: 'Name box' });
  const range = await name.inputValue();
  expect(range).toMatch(/^A1:A\d+$/);
  await gotoCell(page, 'AA1001');
  expect(await name.evaluate((input: HTMLInputElement) => input.validity.customError)).toBe(true);
  await expectActive(page, 'A1');
  await name.press('Escape');
  await expect(name).toHaveValue(range);
  await expect(grid).toBeFocused();
});

test('ribbon operates by keyboard, changes view independently of chrome, and exposes availability', async ({ page }) => {
  await page.goto(path);
  const home = page.getByRole('tab', { name: 'Home' });
  await home.focus();
  await home.press('ArrowRight');
  await expect(page.getByRole('tab', { name: 'Insert' })).toBeFocused();
  await expect(page.getByRole('tabpanel').getByRole('button', { name: 'Table', exact: true })).toBeVisible();
  await page.keyboard.press('End');
  await expect(page.getByRole('tab', { name: 'View' })).toBeFocused();
  const titleHeight = (await page.locator('.title-bar').boundingBox())!.height;
  const formulaBar = page.getByRole('textbox', { name: 'Formula bar' });
  await page.getByRole('checkbox', { name: 'Formula Bar' }).uncheck();
  await expect(formulaBar).toHaveCount(0);
  await page.getByRole('checkbox', { name: 'Formula Bar' }).check();
  await expect(formulaBar).toBeVisible();
  await page.getByRole('button', { name: 'Zoom in' }).click();
  await expect(page.getByRole('slider', { name: 'Zoom' })).toHaveAttribute('aria-valuetext', '110%');
  await page.getByRole('button', { name: '100%' }).click();
  await expect(page.getByRole('slider', { name: 'Zoom' })).toHaveAttribute('aria-valuetext', '100%');
  await page.getByRole('button', { name: 'Zoom to Selection' }).click();
  await expect(page.getByRole('slider', { name: 'Zoom' })).toHaveAttribute('aria-valuetext', '400%');
  await expect(page.getByRole('grid')).toBeFocused();
  expect((await page.locator('.title-bar').boundingBox())!.height).toBe(titleHeight);
});

test('semantic worksheet and command controls pass axe on a narrow viewport', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 640 });
  await page.goto(path);
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('grid').focus();
  await page.addScriptTag({ content: axe.source });
  const result = await page.evaluate(async () => (await (window as Window & { axe?: typeof axe }).axe!.run({
    runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] },
  })).violations.map(({ id, nodes }) => ({ id, targets: nodes.map(node => node.target) })));
  expect(result).toEqual([]);
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
});

test('forced colors redraw canvas and preserve keyboard focus and selection indicators', async ({ page }) => {
  await page.goto(path);
  const grid = page.getByRole('grid');
  const canvas = page.locator('.worksheet-canvas');
  const sample = () => canvas.evaluate((node: HTMLCanvasElement) =>
    Array.from(node.getContext('2d')!.getImageData(Math.round(50 * devicePixelRatio), Math.round(32 * devicePixelRatio), 1, 1).data));
  await grid.focus();
  const before = await sample();
  await page.emulateMedia({ colorScheme: 'light', forcedColors: 'active' });
  expect(await page.evaluate(() => matchMedia('(forced-colors: active)').matches)).toBe(true);
  await expect.poll(sample).not.toEqual(before);
  await expect(grid).toBeFocused();
  await expect(page.locator('.worksheet')).toHaveCSS('background-color',
    await page.evaluate(() => {
      const sample = document.createElement('span');
      sample.style.backgroundColor = 'Canvas';
      document.body.append(sample);
      const color = getComputedStyle(sample).backgroundColor;
      sample.remove();
      return color;
    }));
  await grid.press('ArrowRight');
  await expectActive(page, 'B1');
});
