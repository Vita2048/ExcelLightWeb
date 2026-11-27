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
async function choose(page: Page, trigger: string, name: string, toggle = false) {
  await page.getByRole('button', { name: trigger, exact: true }).click();
  await page.getByRole(toggle ? 'menuitemcheckbox' : 'menuitem', { name, exact: true }).click();
}

test('confirms discards from the Alignment tab and commits merge, font and number drafts in one undo', async ({ page }) => {
  await page.goto(path);
  await go(page, 'B2', '7');
  await go(page, 'C3', 'Discard this');
  await go(page, 'B2');
  const grid = page.getByRole('grid');
  await grid.press('Shift+ArrowRight');
  await grid.press('Shift+ArrowDown');
  await page.getByRole('button', { name: 'Format Cell Alignment', exact: true }).click();
  const format = page.getByRole('dialog', { name: 'Format Cells' });
  const align = format.getByRole('tabpanel', { name: 'Alignment' });
  await expect(align).toBeVisible();
  await align.getByRole('combobox', { name: 'Horizontal alignment' }).selectOption('right');
  await align.getByRole('spinbutton', { name: 'Indent' }).fill('3');
  const degrees = align.getByRole('spinbutton', { name: 'Degrees' });
  await align.locator('.orientation-dial').click({ position: { x: 80, y: 55 } });
  await expect(degrees).toHaveValue('27');
  await degrees.fill('-45');
  await align.getByRole('checkbox', { name: 'Merge cells' }).check();
  await format.getByRole('tab', { name: 'Font', exact: true }).click();
  await format.getByRole('listbox', { name: 'Font sizes' }).selectOption('18');
  await format.getByRole('tab', { name: 'Number', exact: true }).click();
  await format.getByLabel('Category').selectOption('Number');
  await format.getByRole('button', { name: 'OK', exact: true }).click();
  const warning = page.getByRole('dialog', { name: 'Merge cells', exact: true });
  await expect(warning).toContainText('discarded');
  await warning.getByRole('button', { name: 'Cancel' }).click();
  await expect(format).toBeVisible();
  await expect(page.getByRole('gridcell', { name: 'C3, Discard this' })).toBeAttached();
  await format.getByRole('button', { name: 'OK', exact: true }).click();
  await warning.getByRole('button', { name: 'OK', exact: true }).click();
  const anchor = page.getByRole('gridcell', { name: 'B2, 7.00', exact: true });
  await expect(anchor).toHaveAttribute('aria-rowspan', '2');
  await expect(anchor).toHaveAttribute('aria-colspan', '2');
  await expect(page.getByRole('gridcell', { name: /^C3,/ })).toHaveCount(0);
  await go(page, 'C3');
  await expect(grid).toHaveAttribute('aria-activedescendant', /-B2$/);
  await grid.press('ArrowRight');
  await expect(grid).toHaveAttribute('aria-activedescendant', /-D2$/);
  const bounds = (await grid.boundingBox())!;
  await page.mouse.click(bounds.x + 260, bounds.y + 90, { button: 'right' });
  await page.getByRole('group', { name: 'Structure commands' }).getByRole('button', { name: 'Format Cells…' }).click();
  await format.getByRole('tab', { name: 'Alignment', exact: true }).click();
  await expect(align.getByRole('combobox', { name: 'Horizontal alignment' })).toHaveValue('right');
  await expect(align.getByRole('checkbox', { name: 'Merge cells' })).toBeChecked();
  await format.getByRole('button', { name: 'Cancel' }).click();
  await go(page, 'C3');
  await grid.press('Enter');
  await expect(grid).toHaveAttribute('aria-activedescendant', /-B4$/);
  await grid.press('Control+z');
  await expect(page.getByRole('gridcell', { name: 'B2, 7', exact: true })).toBeAttached();
  await expect(page.getByRole('gridcell', { name: 'C3, Discard this' })).toBeAttached();
  await go(page, 'B2');
  await expect(page.getByRole('combobox', { name: 'Font size' })).toHaveValue('11');
});

test('merges across rows, suppresses internal gridlines, and clips rotated text in a zoomed merged rectangle', async ({ page }) => {
  await page.goto(path);
  await go(page, 'A1', 'Wide text in a merged cell');
  await go(page, 'A2', 'Second anchor');
  await go(page, 'A1');
  const grid = page.getByRole('grid');
  await grid.press('Shift+ArrowRight');
  await grid.press('Shift+ArrowDown');
  await choose(page, 'Merge options', 'Merge Across');
  for (const name of ['A1, Wide text in a merged cell', 'A2, Second anchor']) {
    await expect(page.getByRole('gridcell', { name, exact: true })).toHaveAttribute('aria-colspan', '2');
  }
  await choose(page, 'Merge options', 'Unmerge Cells');
  await expect(page.getByRole('gridcell', { name: 'B2, blank', exact: true })).toBeAttached();
  await go(page, 'A2');
  await grid.press('Delete');
  await go(page, 'A1');
  await grid.press('Shift+ArrowRight');
  await grid.press('Shift+ArrowDown');
  await choose(page, 'Merge options', 'Merge & Center');

  const canvas = page.locator('.worksheet-canvas');
  const sample = (x: number, y: number) => canvas.evaluate((element: HTMLCanvasElement, point) =>
    Array.from(element.getContext('2d')!.getImageData(point.x * devicePixelRatio, point.y * devicePixelRatio, 1, 1).data), { x, y });
  await go(page, 'C3');
  await expect.poll(() => sample(139, 34)).toEqual([32, 32, 32, 255]);
  await expect.poll(() => sample(70, 51)).toEqual([32, 32, 32, 255]);

  await go(page, 'A1');
  await choose(page, 'Orientation', 'Rotate Text Up', true);
  await page.getByRole('button', { name: 'Font Color', exact: true }).click();
  await page.getByRole('group', { name: 'Standard colors' }).getByRole('button', { name: 'Color #FF0000', exact: true }).click();
  await page.getByRole('button', { name: 'Zoom in' }).click();
  await page.getByRole('button', { name: 'Zoom in' }).click();
  await go(page, 'C3');
  const ink = () => canvas.evaluate((element: HTMLCanvasElement) => {
    const { width, height } = element;
    const image = element.getContext('2d')!.getImageData(0, 0, width, height).data;
    let left = width, right = -1, top = height, bottom = -1;
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
      const i = 4 * (y * width + x);
      if (image[i]! < 95 || image[i + 1]! > 40 || image[i + 2]! > 40) continue;
      left = Math.min(left, x); right = Math.max(right, x);
      top = Math.min(top, y); bottom = Math.max(bottom, y);
    }
    const ratio = devicePixelRatio;
    return { left: left / ratio, right: right / ratio, top: top / ratio, bottom: bottom / ratio };
  });
  await expect.poll(async () => {
    const { left, right, top, bottom } = await ink();
    return right >= left && bottom - top > right - left && top >= 30 && bottom < 95;
  }).toBe(true);
});

test('aligns stacked fixed-decimal glyphs at the top, bottom, and across the row height', async ({ page }) => {
  await page.goto(path);
  for (const address of ['A1', 'B1', 'C1']) await go(page, address, '10');
  await go(page, 'A1');
  const grid = page.getByRole('grid');
  await grid.press('Shift+ArrowRight');
  await grid.press('Shift+ArrowRight');
  await choose(page, 'Format', 'Row Height…');
  const height = page.getByRole('dialog', { name: 'Row Height', exact: true });
  await height.getByRole('spinbutton', { name: 'Size in pixels' }).fill('184');
  await height.getByRole('button', { name: 'OK' }).click();
  await page.getByRole('button', { name: /^Number Format/u }).click();
  await page.getByRole('menuitem', { name: /^Number(?:$| )/ }).click();
  await choose(page, 'Orientation', 'Vertical Text', true);
  await page.getByRole('button', { name: 'Font Color', exact: true }).click();
  await page.getByRole('group', { name: 'Standard colors' }).getByRole('button', { name: 'Color #FF0000', exact: true }).click();
  await go(page, 'A1');
  await page.getByRole('button', { name: 'Top Align', exact: true }).click();
  await go(page, 'B1');
  await page.getByRole('button', { name: 'Bottom Align', exact: true }).click();
  await go(page, 'C1');
  await page.getByRole('button', { name: 'Format Cell Alignment', exact: true }).click();
  const format = page.getByRole('dialog', { name: 'Format Cells' });
  await format.getByRole('combobox', { name: 'Vertical alignment' }).selectOption('distributed');
  await format.getByRole('button', { name: 'OK' }).click();
  await go(page, 'D1');
  for (const address of ['A1', 'B1', 'C1']) await expect(page.getByRole('gridcell', { name: address + ', 10.00' })).toBeAttached();

  const inkRows = () => page.locator('.worksheet-canvas').evaluate((canvas: HTMLCanvasElement) => {
    const ratio = devicePixelRatio;
    const width = 96 * ratio, height = 184 * ratio;
    const header = Number(document.querySelector('.worksheet-scroll')?.getAttribute('data-row-header') ?? 0);
    return [header, header + 96, header + 192].map(left => {
      const data = canvas.getContext('2d')!.getImageData(left * ratio, 26 * ratio, width, height).data;
      const spans: { top: number; bottom: number }[] = [];
      for (let row = 0; row < height; row++) {
        let visible = false;
        for (let col = 0; col < width; col++) {
          const index = (row * width + col) * 4;
          if (data[index]! > 110 && data[index + 1]! < 55) { visible = true; break; }
        }
        if (!visible) continue;
        const span = spans.at(-1);
        if (span?.bottom === row - 1) span.bottom = row;
        else spans.push({ top: row, bottom: row });
      }
      return spans.map(({ top, bottom }) => ({ top: top / ratio, bottom: bottom / ratio }));
    });
  });
  await expect.poll(async () => {
    const [top, bottom, distributed] = await inkRows();
    if (top!.length !== 5 || bottom!.length !== 5 || distributed!.length !== 5) return false;

    return top![0]!.top <= 5 &&
      bottom!.at(-1)!.bottom >= 178 &&
      distributed![0]!.top > top![0]!.top + 5 &&
      distributed!.at(-1)!.bottom - distributed![0]!.top > 135;
  }).toBe(true);
});

test('shows desktop ribbon artwork, keyboard flyouts, and the Insert and Delete radio dialogs', async ({ page }) => {
  await page.goto(path);
  const trigger = page.getByRole('button', { name: 'Insert', exact: true });
  await expect(trigger.locator('svg path').first()).toBeAttached();
  await trigger.press('ArrowDown');
  const menu = page.getByRole('menu', { name: 'Insert', exact: true });
  const first = menu.getByRole('menuitem', { name: 'Insert Cells…' });
  await expect(first).toBeFocused();
  await expect(first.locator('svg path').first()).toBeAttached();
  await page.keyboard.press('ArrowDown');
  await expect(menu.getByRole('menuitem', { name: 'Insert Sheet Rows' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(trigger).toBeFocused();
  const format = page.getByRole('button', { name: 'Format', exact: true });
  await format.click();
  const hide = page.getByRole('menuitem', { name: 'Hide & Unhide' });
  await hide.press('ArrowRight');
  await expect(page.getByRole('menu', { name: 'Hide & Unhide' }).getByRole('menuitem', { name: 'Hide Rows', exact: true })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(hide).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(format).toBeFocused();
  await choose(page, 'Insert', 'Insert Cells…');
  const insert = page.getByRole('dialog', { name: 'Insert', exact: true });
  await expect(insert.getByRole('radio')).toHaveCount(4);
  await expect(insert.getByRole('radio', { name: 'Shift cells right' })).toBeChecked();
  await insert.getByRole('button', { name: 'Close Insert' }).click();

  await choose(page, 'Delete', 'Delete Cells…');
  const remove = page.getByRole('dialog', { name: 'Delete', exact: true });
  await expect(remove.getByRole('radio', { name: 'Shift cells left' })).toBeChecked();
  await remove.getByRole('radio', { name: 'Shift cells up' }).check();
  await expect(remove.getByRole('radio', { name: 'Shift cells up' })).toBeChecked();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('grid')).toBeFocused();
});
