import { expect, test, type Page } from '@playwright/test';

test.describe.configure({ timeout: 60_000 });

async function menu(page: Page) {
  await page.getByRole('button', { name: 'Borders menu' }).click();
  return page.getByRole('group', { name: 'Borders actions' });
}
async function go(page: Page, at: string) {
  const input = page.getByRole('textbox', { name: 'Name box' });
  await input.fill(at);
  await input.press('Enter');
  await expect(input).toHaveValue(at);
}
async function pixel(page: Page, x: number, y: number) {
  return page.locator('.worksheet-canvas').evaluate((element: HTMLCanvasElement, point) => {
    const ratio = devicePixelRatio;
    return Array.from(element.getContext('2d')!.getImageData(Math.floor(point.x * ratio), Math.floor(point.y * ratio), 1, 1).data);
  }, { x, y });
}
async function thickRed(page: Page) {
  let commands = await menu(page);
  await commands.getByRole('button', { name: 'Border Line Color' }).click();
  const colors = page.getByRole('group', { name: 'Border Line Color' });
  await colors.getByRole('button', { name: 'More Colors…' }).click();
  const picker = page.getByRole('dialog', { name: 'Border Line Color' });
  await picker.getByRole('textbox', { name: 'Custom color hex' }).fill('#FF0000');
  await picker.getByRole('button', { name: 'Apply custom color' }).click();
  commands = await menu(page);
  await commands.getByRole('button', { name: 'Border Line Color' }).click();
  await colors.getByRole('group', { name: 'Recent colors' }).getByRole('button', { name: 'Color #FF0000', exact: true }).click();
  commands = await menu(page);
  await commands.getByText('Line Style', { exact: true }).click();
  await commands.getByRole('button', { name: 'Thick', exact: true }).click();
}
async function gridlinesOff(page: Page) {
  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await page.getByRole('checkbox', { name: 'Gridlines', exact: true }).uncheck();
  await page.getByRole('tab', { name: 'Home', exact: true }).click();
}

test('draws edge-only strokes and rectangular grids, erases a physical line and repeats the remembered tool', async ({ page }) => {
  await page.goto('/ExcelLightWeb/');
  await gridlinesOff(page);
  await thickRed(page);
  const grid = page.getByRole('grid');
  const status = page.locator('.status-bar [role="status"]');
  const box = (await grid.boundingBox())!;
  const headerDelta = Number(await grid.getAttribute('data-row-header')) - 44;
  const xAt = (x: number) => x + headerDelta;
  const drag = async (x: number, y: number, toX: number, toY: number) => {
    await page.mouse.move(box.x + xAt(x), box.y + y);
    await page.mouse.down();
    await page.mouse.move(box.x + xAt(toX), box.y + toY, { steps: 8 });
    await page.mouse.up();
  };
  const red = [255, 0, 0, 255], dark = [32, 32, 32, 255];
  await (await menu(page)).getByRole('button', { name: 'Draw Border', exact: true }).click();
  await expect(status).toContainText('Draw Border');
  await grid.hover({ position: { x: xAt(170), y: 53 } });
  await expect(grid).toHaveCSS('cursor', /crosshair/);
  await drag(170, 53, 275, 53);
  await expect.poll(() => pixel(page, xAt(173), 52)).toEqual(red);
  await expect.poll(() => pixel(page, xAt(270), 52)).toEqual(red);
  await expect.poll(() => pixel(page, xAt(173), 78)).toEqual(dark);
  await expect(status).toContainText('Draw Border');
  await grid.press('Control+z');
  await expect.poll(() => pixel(page, xAt(173), 52)).toEqual(dark);
  await grid.press('Control+y');
  await expect.poll(() => pixel(page, xAt(270), 52)).toEqual(red);

  await (await menu(page)).getByRole('button', { name: 'Draw Border Grid', exact: true }).click();
  await drag(177, 65, 267, 91);
  await expect.poll(() => pixel(page, xAt(179), 78)).toEqual(red);
  await expect.poll(() => pixel(page, xAt(235), 93)).toEqual(red);
  await (await menu(page)).getByRole('button', { name: 'Erase Border', exact: true }).click();
  await drag(177, 69, 269, 69);
  await expect.poll(() => pixel(page, xAt(235), 69)).toEqual(dark);
  await expect.poll(() => pixel(page, xAt(178), 78)).toEqual(red);
  await grid.press('Control+z');
  await expect.poll(() => pixel(page, xAt(235), 69)).toEqual(red);
  await grid.press('Escape');
  await expect(status).not.toContainText('Erase Border');
  await page.getByRole('button', { name: 'Repeat Erase Border' }).click();
  await expect(status).toContainText('Erase Border');
  await page.mouse.click(box.x + xAt(170), box.y + 15);
  await expect(status).not.toContainText('Erase Border');
  await page.getByRole('button', { name: 'Repeat Erase Border' }).click();
  await page.getByRole('button', { name: 'New sheet' }).click();
  await expect(status).not.toContainText('Erase Border');
});

test('edits mixed edges in the shared Border tab and remembers the last ribbon preset', async ({ page }) => {
  await page.goto('/ExcelLightWeb/');
  const grid = page.getByRole('grid');
  await go(page, 'B2');
  await (await menu(page)).getByRole('button', { name: 'Top Border', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Repeat Top Border' })).toBeVisible();
  await go(page, 'D2');
  await page.getByRole('button', { name: 'Repeat Top Border' }).click();
  await go(page, 'B2');
  await grid.press('Shift+ArrowRight');
  const commandMenu = await menu(page);
  await commandMenu.getByRole('button', { name: 'More Borders…' }).click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog.getByRole('tab', { name: 'Border', selected: true })).toBeFocused();
  const border = dialog.getByRole('tabpanel', { name: 'Border' });
  await expect(border.getByRole('button', { name: 'Top', exact: true })).toHaveAttribute('aria-pressed', 'mixed');
  await expect(border.getByRole('button', { name: 'Inside horizontal' })).toBeDisabled();
  await border.getByRole('button', { name: 'Double', exact: true }).click();
  await expect(border.getByRole('button', { name: 'Top', exact: true })).toHaveAttribute('aria-pressed', 'mixed');
  await border.getByRole('button', { name: 'Bottom', exact: true }).click();
  await border.getByRole('button', { name: 'Diagonal up', exact: true }).click();
  await dialog.getByRole('tab', { name: 'Font', exact: true }).click();
  await dialog.getByRole('listbox', { name: 'Font sizes' }).selectOption('18');
  await dialog.getByRole('tab', { name: 'Border', exact: true }).click();
  await expect(border.getByRole('button', { name: 'Bottom', exact: true })).toHaveAttribute('aria-pressed', 'true');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  const size = page.getByRole('combobox', { name: 'Font size', exact: true });
  await expect(size).toHaveValue('18');
  await grid.press('Control+z');
  await expect(size).toHaveValue('11');
  await grid.press('Control+1');
  await dialog.getByRole('tab', { name: 'Border', exact: true }).click();
  await expect(border.getByRole('button', { name: 'Bottom', exact: true })).toHaveAttribute('aria-pressed', 'false');
  await expect(border.getByRole('button', { name: 'Top', exact: true })).toHaveAttribute('aria-pressed', 'mixed');
  await border.getByRole('button', { name: 'Outline', exact: true }).click();
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  await grid.press('Control+1');
  await dialog.getByRole('tab', { name: 'Border', exact: true }).click();
  await expect(border.getByRole('button', { name: 'Top', exact: true })).toHaveAttribute('aria-pressed', 'mixed');
});

test('restyles and removes every visible side of a ribbon outline from Format Cells', async ({ page }) => {
  await page.goto('/ExcelLightWeb/');
  await gridlinesOff(page);
  await go(page, 'B2');
  await (await menu(page)).getByRole('button', { name: 'Outside Borders', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  const border = dialog.getByRole('tabpanel', { name: 'Border' });
  const openBorder = async () => {
    await (await menu(page)).getByRole('button', { name: 'More Borders…' }).click();
    await expect(border).toBeVisible();
  };
  await openBorder();
  await expect(border.getByRole('group', { name: 'Border line style' }).getByRole('button')).toHaveCount(14);
  await expect(border.getByRole('button', { name: 'Inside', exact: true })).toBeDisabled();
  const preview = border.locator('.border-preview');
  await expect(preview).toHaveText('Text');
  await preview.locator('canvas').click({ position: { x: 7, y: 60 } });
  await expect(border.getByRole('button', { name: 'Left', exact: true })).toHaveAttribute('aria-pressed', 'false');

  await border.getByRole('button', { name: 'Thick', exact: true }).click();
  await border.getByRole('button', { name: 'Border Color', exact: true }).click();
  await page.getByRole('group', { name: 'Standard colors' }).getByRole('button', { name: 'Color #FF0000', exact: true }).click();
  await border.getByRole('button', { name: 'Outline', exact: true }).click();
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await go(page, 'A1');

  const header = Number(await page.getByRole('grid').getAttribute('data-row-header'));
  const sites = [[header + 126, 52], [header + 192, 63], [header + 126, 78], [header + 96, 63]] as const;
  const pixels = () => Promise.all(sites.map(([x, y]) => pixel(page, x, y)));
  const red = sites.map(() => [255, 0, 0, 255]);
  const dark = sites.map(() => [32, 32, 32, 255]);
  await expect.poll(pixels).toEqual(red);

  await go(page, 'B2');
  await openBorder();
  await border.getByRole('group', { name: 'Presets' }).getByRole('button', { name: 'None', exact: true }).click();
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await go(page, 'A1');
  await expect.poll(pixels).toEqual(dark);
  await page.getByRole('grid').press('Control+z');
  await expect.poll(pixels).toEqual(red);
});
