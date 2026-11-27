import { expect, test, type Locator, type Page } from '@playwright/test';

test.describe.configure({ timeout: 60_000 });

const path = '/ExcelLightWeb/';
const toolbar = (page: Page) => page.getByRole('tabpanel', { name: 'Home' });

async function go(page: Page, address: string) {
  const box = page.getByRole('textbox', { name: 'Name box' });
  await box.fill(address);
  await box.press('Enter');
}

async function enter(page: Page, address: string, text: string) {
  await go(page, address);
  const bar = page.getByRole('textbox', { name: 'Formula bar' });
  await bar.fill(text);
  await bar.press('Enter');
}

function region(canvas: Locator, x: number, y: number, width: number, height: number) {
  return canvas.evaluate((node: HTMLCanvasElement, box) => {
    const ratio = devicePixelRatio;
    return Array.from(node.getContext('2d')!.getImageData(box.x * ratio, box.y * ratio, box.width * ratio, box.height * ratio).data);
  }, { x, y, width, height });
}

function pixel(canvas: Locator, x: number, y: number) {
  return canvas.evaluate((node: HTMLCanvasElement, point) => {
    const ratio = devicePixelRatio;
    return Array.from(node.getContext('2d')!.getImageData(
      Math.floor((point.x + .5) * ratio), Math.floor((point.y + .5) * ratio), 1, 1).data);
  }, { x, y });
}
const canvasFor = (page: Page) => page.locator('.worksheet-canvas');
const fillAt = (page: Page, column: number, row: number, x = 10, y = 8) =>
  canvasFor(page).evaluate((node: HTMLCanvasElement, point) => {
    const header = Number(node.parentElement?.querySelector('.worksheet-scroll')?.getAttribute('data-row-header') ?? 0);
    const ratio = devicePixelRatio;
    const sampleX = header + (point.column - 1) * 96 + point.x;
    const sampleY = 26 + (point.row - 1) * 26 + point.y;
    return Array.from(node.getContext('2d')!.getImageData(
      Math.floor((sampleX + .5) * ratio), Math.floor((sampleY + .5) * ratio), 1, 1).data);
  }, { column, row, x, y });

async function color(dialog: Locator, hex: string) {
  const input = dialog.getByRole('textbox', { name: 'Custom color hex' });
  await input.fill(hex);
  await dialog.getByRole('button', { name: 'Apply custom color' }).click();
}

test('retains typed unavailable font names and paints the browser fallback, then validates ribbon sizes', async ({ page }) => {
  await page.goto(path);
  await enter(page, 'A1', 'Fallback');
  await enter(page, 'B1', 'Fallback');
  await go(page, 'A1');
  const font = toolbar(page).getByRole('combobox', { name: 'Font family' });
  const size = toolbar(page).getByRole('combobox', { name: 'Font size' });
  const unavailable = 'ExcelLight Missing Font 938';
  await font.fill(unavailable);
  await font.press('Enter');
  await expect(font).toHaveValue(unavailable);
  await go(page, 'C1');
  const canvas = canvasFor(page);
  const fallback = await region(canvas, 47, 28, 88, 22);
  const original = await region(canvas, 143, 28, 88, 22);
  const ink = original.filter((value, index) => index % 4 === 0 && (value > 128 || fallback[index]! > 128));
  const overlap = original.filter((value, index) => index % 4 === 0 && value > 128 && fallback[index]! > 128);
  expect(ink.length).toBeGreaterThan(30);
  expect(overlap.length / ink.length).toBeGreaterThan(.82);
  await go(page, 'A1');
  await size.fill('73');
  await size.press('Enter');
  await expect(page.getByRole('status').first()).toContainText('6 to 72');
  await expect(size).toHaveValue('11');
  await size.fill('6');
  await size.press('Enter');
  await expect(size).toHaveValue('6');
  await toolbar(page).getByRole('button', { name: 'Increase Font Size' }).click();
  await expect(size).toHaveValue('7');
  await toolbar(page).getByRole('button', { name: 'Decrease Font Size' }).click();
  await expect(size).toHaveValue('6');
  await toolbar(page).getByRole('button', { name: 'Format Cell Font' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog.getByRole('tab', { name: 'Font', selected: true })).toBeFocused();
  await expect(dialog.getByRole('textbox', { name: 'Font family' })).toHaveValue(unavailable);
  await expect(dialog.getByRole('textbox', { name: 'Font size' })).toHaveValue('6');
  await dialog.getByRole('checkbox', { name: 'Normal font' }).check();
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(font).toHaveValue('Segoe UI');

});

test('commits font, number and pattern changes together and keeps nested color and gradient cancellation local', async ({ page }) => {
  await page.goto(path);
  await enter(page, 'A1', '12.5');
  await go(page, 'A1');
  const grid = page.getByRole('grid');
  await grid.press('Control+1');
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await dialog.getByLabel('Category').selectOption('Number');
  const numberTab = dialog.getByRole('tab', { name: 'Number' });
  await numberTab.focus();
  await page.keyboard.press('ArrowRight');
  await expect(dialog.getByRole('tab', { name: 'Alignment', selected: true })).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(dialog.getByRole('tab', { name: 'Font', selected: true })).toBeFocused();
  await dialog.getByRole('listbox', { name: 'Font styles' }).selectOption('Bold Italic');
  await dialog.getByRole('combobox', { name: 'Underline' }).selectOption('Double Accounting');
  await dialog.getByRole('checkbox', { name: 'Strikethrough' }).check();
  await dialog.getByRole('checkbox', { name: 'Superscript' }).check();
  await dialog.getByRole('checkbox', { name: 'Subscript' }).check();
  await expect(dialog.getByRole('checkbox', { name: 'Superscript' })).not.toBeChecked();
  await dialog.getByRole('button', { name: 'Font Color' }).click();
  const fontColor = page.getByRole('dialog', { name: 'Font Color', exact: true });
  await fontColor.getByRole('textbox', { name: 'Custom color hex' }).fill('#GG1234');
  await expect(fontColor.getByRole('button', { name: 'Apply custom color' })).toBeDisabled();
  await page.keyboard.press('Escape');
  await expect(fontColor).toHaveCount(0);
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Font Color' }).click();
  await color(fontColor, '#ff0000');
  const fontColorTrigger = dialog.getByRole('button', { name: 'Font Color' });
  await expect(fontColorTrigger).toHaveText('');
  await expect(fontColorTrigger.locator('.color-swatch-bar')).toHaveCSS('background-color', 'rgb(255, 0, 0)');
  const [underlineHeight, colorHeight] = await Promise.all([
    dialog.getByRole('combobox', { name: 'Underline', exact: true }).evaluate(node => node.getBoundingClientRect().height),
    fontColorTrigger.evaluate(node => node.getBoundingClientRect().height),
  ]);
  expect(colorHeight).toBe(underlineHeight);

  await dialog.getByRole('tab', { name: 'Fill', exact: true }).click();
  await dialog.getByRole('group', { name: 'Background Color' }).getByRole('button', { name: 'Color #FFFFFF' }).first().click();
  await dialog.getByRole('button', { name: 'Pattern Color' }).click();
  await color(page.getByRole('dialog', { name: 'Pattern Color' }), '#0070C0');
  await dialog.getByLabel('Pattern Style: None').click();
  await dialog.getByRole('button', { name: 'Dark vertical', exact: true }).click();
  const sample = dialog.getByRole('img', { name: 'Fill sample' });
  await expect.poll(() => pixel(sample, 8, 8)).toEqual([0, 112, 192, 255]);
  await expect.poll(() => pixel(sample, 10, 8)).toEqual([255, 255, 255, 255]);
  await dialog.getByRole('button', { name: 'Fill Effects…' }).click();
  const effects = page.getByRole('dialog', { name: 'Fill Effects', exact: true });
  await effects.getByRole('radio', { name: 'From corner' }).check();
  await effects.getByRole('button', { name: 'Gradient variant 4' }).click();
  await effects.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog.getByLabel('Pattern Style: Dark vertical')).toBeVisible();
  await dialog.getByRole('button', { name: 'OK' }).click();

  await expect(page.getByRole('gridcell', { name: 'A1, 12.50', exact: true })).toBeAttached();
  await expect(toolbar(page).getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await expect.poll(() => fillAt(page, 1, 1, 8)).toEqual([0, 112, 192, 255]);
  await expect.poll(() => fillAt(page, 1, 1, 10)).toEqual([255, 255, 255, 255]);
  const ink = await region(canvasFor(page), 90, 29, 45, 21);
  expect(ink.some((channel, index) => index % 4 === 0 && channel > 180 && ink[index + 1]! < 50 && ink[index + 2]! < 50)).toBe(true);
  await grid.press('Control+z');
  await expect(page.getByRole('gridcell', { name: 'A1, 12.5', exact: true })).toBeAttached();
  await expect(toolbar(page).getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'false');
  await grid.press('Control+y');
  await expect(page.getByRole('gridcell', { name: 'A1, 12.50', exact: true })).toBeAttached();
  await expect.poll(() => fillAt(page, 1, 1, 8)).toEqual([0, 112, 192, 255]);
});

test('paints the chosen gradient, cancels the Fill draft and replaces the effect with ribbon colors', async ({ page }) => {
  await page.goto(path);
  const grid = page.getByRole('grid');
  await grid.press('Control+1');
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await dialog.getByRole('tab', { name: 'Fill', exact: true }).click();
  await dialog.getByRole('button', { name: 'Fill Effects…' }).click();
  const effects = page.getByRole('dialog', { name: 'Fill Effects', exact: true });
  await effects.getByRole('radio', { name: 'Vertical', exact: true }).check();
  await effects.getByRole('button', { name: 'Gradient variant 3' }).click();
  const sample = effects.getByRole('img', { name: 'Gradient sample' });
  const bounds = (await sample.boundingBox())!;
  const [red] = await pixel(sample, bounds.width / 2, 40);
  expect(red).toBeLessThan(115);
  await effects.getByRole('button', { name: 'OK' }).click();
  await dialog.getByRole('button', { name: 'OK' }).click();
  const edge = await fillAt(page, 1, 1, 6);
  const middle = await fillAt(page, 1, 1, 48);
  expect(edge[0]).toBeGreaterThan(middle[0]! + 100);

  await grid.press('Control+1');
  await dialog.getByRole('tab', { name: 'Fill', exact: true }).click();
  await dialog.getByRole('button', { name: 'More Colors…' }).click();
  await page.getByRole('dialog', { name: 'More Colors…' }).getByRole('button', { name: 'No Color' }).click();
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(grid).toBeFocused();
  await expect.poll(() => fillAt(page, 1, 1, 48)).toEqual(middle);
  await toolbar(page).getByRole('button', { name: 'Fill Color', exact: true }).click();
  const fillColor = page.getByRole('group', { name: 'Fill Color', exact: true });
  await fillColor.getByRole('button', { name: 'Color #156082', exact: true }).click();
  await expect.poll(() => fillAt(page, 1, 1)).toEqual([21, 96, 130, 255]);
  await toolbar(page).getByRole('button', { name: 'Fill Color', exact: true }).click();
  await fillColor.getByRole('button', { name: 'No Fill' }).click();
  await expect.poll(() => fillAt(page, 1, 1)).toEqual([32, 32, 32, 255]);
});

test('Format Painter tiles a locked source across sheets, applies once in single mode, and cancels with Escape', async ({ page }) => {
  await page.goto(path);
  const grid = page.getByRole('grid');
  await expect(page.getByRole('button', { name: 'Save', exact: true })).toBeEnabled();
  await page.getByRole('button', { name: 'New sheet' }).click();
  const destinationSheet = page.getByRole('button', { name: 'Sheet2', exact: true });
  await expect(destinationSheet).toBeVisible();
  await destinationSheet.click();
  await expect(destinationSheet).toHaveAttribute('aria-current', 'true');
  await enter(page, 'C3', '=1+2');
  await enter(page, 'D3', 'keep');
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click();
  await enter(page, 'A1', '123');
  await go(page, 'A1');
  await toolbar(page).getByRole('button', { name: 'Apply last fill color' }).click();
  await toolbar(page).getByRole('button', { name: /^Number Format/u }).click();
  await page.getByRole('menuitem', { name: /^Number(?:$| )/ }).click();
  await grid.press('Shift+ArrowRight');

  const painter = toolbar(page).getByRole('button', { name: 'Format Painter', exact: true });
  const lock = toolbar(page).getByRole('button', { name: 'Lock Format Painter' });
  await painter.dblclick();
  await expect(lock).toHaveAttribute('aria-pressed', 'true');
  await page.getByRole('button', { name: 'Sheet2', exact: true }).click();
  const box = (await grid.boundingBox())!;
  const dragTo = async (left: number, top: number, right: number, bottom: number) => {
    await page.mouse.move(box.x + 60 + (left - 1) * 96, box.y + 38 + (top - 1) * 26);
    await page.mouse.down();
    await page.mouse.move(box.x + 60 + (right - 1) * 96, box.y + 38 + (bottom - 1) * 26, { steps: 5 });
    await page.mouse.up();
  };
  await dragTo(3, 3, 5, 4);
  await expect(page.getByRole('gridcell', { name: 'C3, 3.00', exact: true })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'D3, keep', exact: true })).toBeAttached();
  await go(page, 'C3');
  const bar = page.getByRole('textbox', { name: 'Formula bar' });
  await expect(bar).toHaveValue('=1+2');
  await expect.poll(() => fillAt(page, 5, 4)).toEqual([255, 255, 0, 255]);
  await expect(lock).toHaveAttribute('aria-pressed', 'true');
  await go(page, 'G3');
  await grid.press('Enter');
  await expect.poll(() => fillAt(page, 7, 3)).toEqual([255, 255, 0, 255]);
  await grid.press('Control+z');
  await expect.poll(() => fillAt(page, 7, 3)).toEqual([32, 32, 32, 255]);
  await expect.poll(() => fillAt(page, 5, 4)).toEqual([255, 255, 0, 255]);
  await grid.press('Escape');
  await expect(painter).toHaveAttribute('aria-pressed', 'false');

  await go(page, 'C3');
  await painter.click();
  await go(page, 'H3');
  await grid.press('Enter');
  await expect(painter).toHaveAttribute('aria-pressed', 'false');
  await expect.poll(() => fillAt(page, 8, 3)).toEqual([255, 255, 0, 255]);
  await painter.click();
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click();
  await expect(painter).toHaveAttribute('aria-pressed', 'false');
});

test('Format Cells tabs keep desktop order and edited panes on a narrow viewport', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 550 });
  await page.goto(path);
  const grid = page.getByRole('grid');
  await grid.press('Control+1');
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  const tabs = dialog.getByRole('tablist');
  await expect(tabs.getByRole('tab')).toHaveText(['Number', 'Alignment', 'Font', 'Border', 'Fill']);
  await expect(tabs.getByRole('tab', { name: 'Alignment' })).toBeEnabled();
  await expect(tabs.getByRole('tab', { name: 'Border' })).toBeEnabled();
  await tabs.getByRole('tab', { name: 'Number' }).focus();
  await page.keyboard.press('ArrowRight');
  await expect(tabs.getByRole('tab', { name: 'Alignment', selected: true })).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(tabs.getByRole('tab', { name: 'Font', selected: true })).toBeFocused();
  const font = dialog.getByRole('tabpanel', { name: 'Font' });
  await font.getByRole('listbox', { name: 'Font sizes' }).selectOption('18');
  await font.getByRole('listbox', { name: 'Font catalog' }).selectOption('Georgia');
  await tabs.getByRole('tab', { name: 'Fill' }).click();
  await expect(font).toHaveCount(0);
  await expect(dialog.getByRole('tabpanel', { name: 'Fill' })).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeInViewport();
  await tabs.getByRole('tab', { name: 'Font' }).click();
  await expect(font.getByRole('textbox', { name: 'Font family' })).toHaveValue('Georgia');
  await expect(font.getByRole('textbox', { name: 'Font size' })).toHaveValue('18');
  await dialog.getByRole('button', { name: 'Close Format Cells' }).click();
  await expect(grid).toBeFocused();
  await grid.press('Control+1');
  await tabs.getByRole('tab', { name: 'Font' }).click();
  await expect(font.getByRole('textbox', { name: 'Font size' })).toHaveValue('11');
});
