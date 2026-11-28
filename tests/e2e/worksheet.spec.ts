import axe from 'axe-core';
import { expect, test, type Page } from '@playwright/test';

const path = '/ExcelLightWeb/';
test.describe.configure({ timeout: 60_000 });

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

async function rowHeader(page: Page) {
  return Number(await page.getByRole('grid').getAttribute('data-row-header'));
}

/** Painted caret and row-number ink along the bottom of the canvas. `cellOffset` is the active column's CSS x inside the grid body. */
async function bottomOfGrid(page: Page, cellOffset: number, band = 36) {
  return page.locator('.worksheet-canvas').evaluate((canvas: HTMLCanvasElement, probe: { cellOffset: number; band: number }) => {
    const grid = canvas.closest('.worksheet')!.querySelector('.worksheet-scroll') as HTMLElement;
    const header = Number(grid.getAttribute('data-row-header'));
    const ratio = window.devicePixelRatio || 1;
    const context = canvas.getContext('2d')!;
    const bandTop = Math.max(0, canvas.height - Math.round(probe.band * ratio));
    const origin = Math.round((header + probe.cellOffset + 6) * ratio);
    const sampleWidth = Math.min(Math.round(70 * ratio), canvas.width - origin);
    const cell = context.getImageData(origin, bandTop, Math.max(1, sampleWidth), canvas.height - bandTop).data;
    let accent = 0;
    for (let index = 0; index < cell.length; index += 4) {
      const red = cell[index] ?? 0;
      const green = cell[index + 1] ?? 0;
      const blue = cell[index + 2] ?? 0;
      if (blue > 170 && blue > red + 40 && green > 140) accent += 1;
    }
    const labelHeight = Math.min(canvas.height, Math.round(120 * ratio));
    const labelWidth = Math.min(canvas.width, Math.max(1, Math.round(header * ratio)));
    const label = context.getImageData(0, canvas.height - labelHeight, labelWidth, labelHeight).data;
    let minX = labelWidth;
    let maxX = -1;
    for (let index = 0; index < label.length; index += 4) {
      const red = label[index] ?? 0;
      const green = label[index + 1] ?? 0;
      const blue = label[index + 2] ?? 0;
      if (red > 210 && green > 210 && blue > 210) {
        const x = (index / 4) % labelWidth;
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
      }
    }
    return {
      accent,
      minX: minX / ratio,
      maxX: maxX / ratio,
      header,
      gap: grid.scrollHeight - grid.clientHeight - grid.scrollTop,
      spacer: Number.parseFloat((grid.querySelector('.worksheet-spacer') as HTMLElement).style.height),
    };
  }, { cellOffset, band });
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
  await expectActive(page, 'B1048576');
  let bottom = await bottomOfGrid(page, 96);
  await expect.poll(async () => {
    bottom = await bottomOfGrid(page, 96);
    return bottom.accent;
  }).toBeGreaterThan(8);
  expect(bottom.gap).toBeLessThan(2);
  expect(bottom.spacer).toBeGreaterThan(16_777_000);
  expect(bottom.spacer).toBeLessThan(16_777_216);
  expect(bottom.minX).toBeGreaterThan(1);
  expect(bottom.maxX).toBeLessThan(bottom.header - 1);
  const headerFit = await page.evaluate(() => {
    const grid = document.querySelector('.worksheet-scroll')!;
    const header = Number(grid.getAttribute('data-row-header'));
    const context = document.createElement('canvas').getContext('2d')!;
    context.font = `12px ${getComputedStyle(document.documentElement).fontFamily}`;
    return { header, text: context.measureText('1048576').width };
  });
  expect(headerFit.header).toBeGreaterThan(headerFit.text);
  await grid.press('Control+Home');
  await expectActive(page, 'A1');
  await grid.press('PageDown');
  await expect.poll(() => grid.evaluate(node => node.scrollTop)).toBeGreaterThan(0);
  await grid.press('Shift+Space');
  await expect(name).toHaveValue(/^A\d+:XFD\d+$/);
  await grid.press('Control+Space');
  await expect(name).toHaveValue('A1:A1048576');
  await grid.press('Control+a');
  await expect(name).toHaveValue('A1:XFD1048576');
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
  const header = await rowHeader(page);
  const { x, y } = (await grid.boundingBox())!;
  const name = page.getByRole('textbox', { name: 'Name box' });
  await page.mouse.move(x + header + 30, y + 36);
  await page.mouse.down();
  await page.mouse.move(x + header + 140, y + 88, { steps: 8 });
  await page.mouse.up();
  await expect(name).toHaveValue('A1:B3');
  await expect(grid).toBeFocused();
  await page.keyboard.down('Control');
  await page.mouse.click(x + header + 96 * 3 + 40, y + 138);
  await expect(name).toHaveValue('A1:B3, D5');
  await expectActive(page, 'D5');
  await page.mouse.click(x + header + 30, y + 36);
  await page.keyboard.up('Control');
  await expect(name).toHaveValue('D5');
  await page.mouse.click(x + header + 96 + 40, y + 9);
  await expect(name).toHaveValue('B1:B1048576');
  await page.mouse.click(x + 12, y + 62);
  await expect(name).toHaveValue('A2:XFD2');
  await page.mouse.click(x + 12, y + 8);
  await expect(name).toHaveValue('A1:XFD1048576');
});

test('native scrolling virtualizes both axes and retains an offscreen active descendant', async ({ page }) => {
  await page.setViewportSize({ width: 850, height: 520 });
  await page.goto(path);
  const grid = page.getByRole('grid');
  await grid.evaluate(node => node.scrollTo({ top: node.scrollHeight / 2, left: 1200 }));
  await expect(page.getByRole('gridcell', { name: /^M\d+, blank$/ }).first()).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'B2, blank' })).toHaveCount(0);
  await expectActive(page, 'A1');
  expect(await page.getByRole('gridcell').count()).toBeLessThan(250);
  await gotoCell(page, 'z1000');
  await expect(grid).toBeFocused();
  await expectActive(page, 'Z1000');
  await expect.poll(() => grid.evaluate(node => node.scrollLeft)).toBeGreaterThan(1500);
  await expect.poll(() => grid.evaluate(node => node.scrollTop)).toBeGreaterThan(10_000);
  const canvas = page.locator('.worksheet-canvas');
  const dimensions = await grid.evaluate(node => ({ width: node.clientWidth, height: node.clientHeight }));
  await expect(canvas).toHaveCSS('width', dimensions.width + 'px');
  await expect(canvas).toHaveCSS('height', dimensions.height + 'px');
});

test('dragging beyond the worksheet scrolls the range and the name box rejects invalid coordinates', async ({ page }) => {
  await page.setViewportSize({ width: 880, height: 540 });
  await page.goto(path);
  const grid = page.getByRole('grid');
  const header = await rowHeader(page);
  const { x, y, height } = (await grid.boundingBox())!;
  await page.mouse.move(x + header + 30, y + 38);
  await page.mouse.down();
  await page.mouse.move(x + header + 30, y + height + 18, { steps: 8 });
  await expect.poll(() => grid.evaluate(node => node.scrollTop)).toBeGreaterThan(30);
  await page.mouse.up();
  const name = page.getByRole('textbox', { name: 'Name box' });
  const range = await name.inputValue();
  expect(range).toMatch(/^A1:A\d+$/);
  await gotoCell(page, 'XFE1');
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
  // The test runner injects axe as a pre-document init script; production CSP stays restrictive.
  await page.context().addInitScript(axe.source);
  await page.goto(path);
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('grid').focus();
  const result = await page.evaluate(async () => (await (window as Window & { axe?: typeof axe }).axe!.run({
    runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] },
  })).violations.map(({ id, nodes }) => ({ id, targets: nodes.map(node => node.target) })));
  expect(result).toEqual([]);
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
});

test('name box edits the Excel corner and select-all addresses the full grid', async ({ page }) => {
  await page.goto(path);
  const grid = page.getByRole('grid');
  const name = page.getByRole('textbox', { name: 'Name box' });
  await gotoCell(page, 'XFD1048576');
  await expect(grid).toBeFocused();
  await expectActive(page, 'XFD1048576');
  await expect(name).toHaveValue('XFD1048576');
  await page.keyboard.press('4');
  await page.keyboard.press('2');
  await page.keyboard.press('Enter');
  await gotoCell(page, 'XFD1048576');
  await expect(page.getByRole('gridcell', { name: 'XFD1048576, 42', exact: true })).toBeAttached();
  await gotoCell(page, 'A1');
  await expectActive(page, 'A1');
  await grid.press('Control+ArrowDown');
  await expectActive(page, 'A1048576');
  await grid.press('Control+a');
  await expect(name).toHaveValue('A1:XFD1048576');
  expect(await page.getByRole('gridcell').count()).toBeLessThan(250);
});

test('the last row stays on screen, a sparse column sums, and one wheel notch stays short', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto(path);
  const grid = page.getByRole('grid');
  const editor = page.getByRole('textbox', { name: 'Formula bar' });
  await grid.focus();
  await grid.press('Control+ArrowDown');
  await expectActive(page, 'A1048576');
  let bottom = await bottomOfGrid(page, 0);
  await expect.poll(async () => {
    bottom = await bottomOfGrid(page, 0);
    return bottom.accent;
  }).toBeGreaterThan(8);
  expect(bottom.gap).toBeLessThan(2);
  expect(bottom.minX).toBeGreaterThan(1);
  expect(bottom.maxX).toBeLessThan(bottom.header - 1);

  const box = (await grid.boundingBox())!;
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  const atEnd = await grid.evaluate(node => node.scrollTop);
  await page.mouse.wheel(0, -120);
  await expect.poll(() => grid.evaluate(node => node.scrollTop)).toBeLessThan(atEnd - 10);
  expect(atEnd - await grid.evaluate(node => node.scrollTop)).toBeLessThan(2_000);
  await page.mouse.wheel(0, 8_000);
  await expect.poll(async () => (await bottomOfGrid(page, 0)).accent).toBeGreaterThan(8);

  async function enter(address: string, source: string) {
    await gotoCell(page, address);
    await editor.fill(source);
    await editor.press('Enter');
  }
  await enter('A1', '2');
  await enter('A1048576', '3');
  await expect(page.getByRole('gridcell', { name: 'A1048576, 3', exact: true })).toBeAttached();
  await enter('B1', '=SUM(A:A)');
  await expect(page.getByRole('gridcell', { name: 'B1, 5', exact: true })).toBeAttached();
  await gotoCell(page, 'B1');
  await grid.press('Control+z');
  await expect(page.getByRole('gridcell', { name: 'B1, blank', exact: true })).toBeAttached();

  await gotoCell(page, 'A1048576');
  const end = (await grid.boundingBox())!;
  await page.mouse.click(end.x + (await rowHeader(page)) / 2, end.y + end.height - 13, { button: 'right' });
  await page.getByRole('group', { name: 'Structure commands' }).getByRole('button', { name: 'Hide Rows', exact: true }).click();
  await gotoCell(page, 'A1');
  await grid.press('Control+ArrowDown');
  await expectActive(page, 'A1048575');
  await expect.poll(async () => (await bottomOfGrid(page, 0)).accent).toBeGreaterThan(8);

  await page.getByRole('button', { name: /^Zoom percentage/ }).click();
  const zoomDialog = page.getByRole('dialog', { name: 'Zoom', exact: true });
  await zoomDialog.getByRole('spinbutton', { name: 'Custom zoom percentage, 10 to 400' }).fill('400');
  await zoomDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(page.getByRole('slider', { name: 'Zoom' })).toHaveAttribute('aria-valuetext', '400%');
  await expect.poll(async () => (await bottomOfGrid(page, 0, 220)).accent).toBeGreaterThan(8);
  expect((await bottomOfGrid(page, 0, 220)).gap).toBeLessThan(2);
  const zoomBox = (await grid.boundingBox())!;
  await page.mouse.move(zoomBox.x + zoomBox.width / 2, zoomBox.y + zoomBox.height / 2);
  const atZoom = await grid.evaluate(node => node.scrollTop);
  await page.mouse.wheel(0, -120);
  await expect.poll(() => grid.evaluate(node => node.scrollTop)).toBeLessThan(atZoom - 10);
  expect(atZoom - await grid.evaluate(node => node.scrollTop)).toBeLessThan(2_000);
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

test('repeated arrow navigation scrolls smoothly without vertical drift on horizontal moves', async ({ page }) => {
  await page.goto(path);
  const grid = page.getByRole('grid');
  await grid.focus();
  const sample = await grid.evaluate(async node => {
    for (let index = 0; index < 35; index++) {
      node.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', repeat: index > 0, bubbles: true, cancelable: true }));
      await new Promise(resolve => requestAnimationFrame(resolve));
    }
    const verticalBeforeHorizontal = node.scrollTop;
    const verticalAfterHorizontal: number[] = [];
    for (let index = 0; index < 10; index++) {
      node.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', repeat: index > 0, bubbles: true, cancelable: true }));
      await new Promise(resolve => requestAnimationFrame(resolve));
      verticalAfterHorizontal.push(node.scrollTop);
    }
    return { verticalBeforeHorizontal, verticalAfterHorizontal, active: node.getAttribute('aria-activedescendant') };
  });
  expect(sample.verticalBeforeHorizontal).toBeGreaterThan(0);
  expect(sample.verticalAfterHorizontal.every(top => top === sample.verticalBeforeHorizontal)).toBe(true);
  expect(sample.active).toMatch(/-K36$/u);
});
