import { expect, test, type Page } from '@playwright/test';

const path = '/ExcelLightWeb/';

async function pixels(page: Page, points: [number, number][]) {
  return page.locator('.worksheet-canvas').evaluate((element: HTMLCanvasElement, samples) => {
    const context = element.getContext('2d')!;
    // A gridline ends at the CSS boundary and occupies its last device pixel.
    const devicePixel = (coordinate: number) => Math.ceil((coordinate + 1) * devicePixelRatio) - 1;
    return samples.map(([x, y]) =>
      Array.from(context.getImageData(devicePixel(x), devicePixel(y), 1, 1).data));
  }, points);
}

test('fixed worksheet geometry aligns headers and cells before and after fractional-cell scrolling', async ({ page }) => {
  await page.setViewportSize({ width: 880, height: 540 });
  await page.goto(path);
  const grid = page.getByRole('grid', { name: 'Inputs' });
  await expect(page.getByRole('gridcell', { name: 'A1, 8' })).toBeAttached();
  const line = [56, 56, 56, 255];
  const white = [255, 255, 255, 255];
  // Column C is 112px. Row 4 is 38px. Row 5 is hidden.
  await expect.poll(() => pixels(page, [[347, 4], [347, 109], [6, 141], [210, 141], [50, 32]]))
    .toEqual([line, line, line, line, white]);
  await grid.evaluate(node => node.scrollTo({ left: 37, top: 13 }));
  await expect.poll(() => pixels(page, [[310, 4], [310, 96], [6, 128], [210, 128]]))
    .toEqual([line, line, line, line]);
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('checkbox', { name: 'Gridlines' }).uncheck();
  await expect.poll(() => pixels(page, [[310, 4], [310, 96]]))
    .toEqual([line, [32, 32, 32, 255]]);
  await page.getByRole('checkbox', { name: 'Headings' }).uncheck();
  await expect.poll(() => pixels(page, [[1, 1]])).not.toEqual([[38, 38, 38, 255]]);
});

test('session snapshots supply formula source, visible values and selection statistics across sheets', async ({ page }) => {
  await page.goto(path);
  const inputs = page.getByRole('grid', { name: 'Inputs' });
  const nameBox = page.getByRole('textbox', { name: 'Name box' });
  const formulaBar = page.getByRole('textbox', { name: 'Formula bar' });
  await expect(formulaBar).toHaveValue('8');
  await inputs.focus();
  await inputs.press('Shift+ArrowDown');
  await inputs.press('Shift+ArrowRight');
  await expect(nameBox).toHaveValue('A1:B2');
  const stats = page.getByLabel('Selection statistics');
  await expect(stats).toContainText('Sum: 24');
  await expect(stats).toContainText('Average: 12');
  await expect(stats).toContainText('Count: 2');
  await nameBox.fill('B2');
  await nameBox.press('Enter');
  await expect(formulaBar).toHaveValue('=A1*2');
  await expect(page.getByRole('gridcell', { name: 'B2, 16' })).toHaveAttribute('aria-selected', 'true');
  await page.getByRole('button', { name: 'Next sheet' }).click();
  const totals = page.getByRole('grid', { name: 'Totals' });
  await expect(totals).toBeFocused();
  await expect(formulaBar).toHaveValue('Summary');
  await totals.press('Control+PageUp');
  await expect(inputs).toBeFocused();
  await expect(nameBox).toHaveValue('B2');
  await inputs.press('Control+End');
  await expect(nameBox).toHaveValue('Z1000');
  await expect(formulaBar).toHaveValue('Last cell');
});

test('draws and erases on a frozen cell at zoom after scrolling both axes', async ({ page }) => {
  await page.setViewportSize({ width: 1100, height: 700 });
  await page.goto(path + '?frozen');
  const grid = page.getByRole('grid', { name: 'Inputs' });
  for (let click = 0; click < 5; click++) await page.getByRole('button', { name: 'Zoom in' }).click();
  await expect(page.getByRole('slider', { name: 'Zoom' })).toHaveAttribute('aria-valuetext', '150%');
  const nameBox = page.getByRole('textbox', { name: 'Name box' });
  await nameBox.fill('B2');
  await nameBox.press('Enter');
  await grid.evaluate(node => node.scrollTo({ left: 190, top: 74 }));
  await expect.poll(() => grid.evaluate(node => [node.scrollLeft, node.scrollTop])).toEqual([190, 74]);
  const commands = page.getByRole('group', { name: 'Borders actions' });
  const open = () => page.getByRole('button', { name: 'Borders menu' }).click();
  await open();
  await commands.getByRole('button', { name: 'Border Line Color' }).click();
  await page.getByRole('group', { name: 'Border Line Color' }).getByRole('button', { name: 'Color #FF0000', exact: true }).click();
  await open();
  await commands.getByText('Line Style', { exact: true }).click();
  await commands.getByRole('button', { name: 'Thick', exact: true }).click();
  const box = (await grid.boundingBox())!;
  // The frozen divider covers row 1's bottom edge, so draw on its visible top edge instead.
  await page.mouse.click(box.x + 108, box.y + 39);
  await expect.poll(() => pixels(page, [[108, 39]])).toEqual([[255, 0, 0, 255]]);
  await open();
  await commands.getByRole('button', { name: 'Erase Border', exact: true }).click();
  await page.mouse.click(box.x + 108, box.y + 39);
  await expect.poll(() => pixels(page, [[108, 39]])).toEqual([[255, 255, 255, 255]]);
  await grid.press('Control+z');
  await expect.poll(() => pixels(page, [[108, 39]])).toEqual([[255, 0, 0, 255]]);
});

test('projects one merged fill through the four frozen panes and hides the internal gridlines', async ({ page }) => {
  await page.goto(path + '?frozen&merged');
  const grid = page.getByRole('grid', { name: 'Inputs' });
  const anchor = page.getByRole('gridcell', { name: 'A1, Anchor', exact: true });
  await expect(anchor).toHaveAttribute('aria-rowspan', '3');
  await expect(anchor).toHaveAttribute('aria-colspan', '3');
  await grid.evaluate(node => node.scrollTo({ left: 12, top: 7 }));
  await expect.poll(() => grid.evaluate(node => [node.scrollLeft, node.scrollTop])).toEqual([12, 7]);
  const ink = [21, 96, 130, 255];
  const divider = [160, 160, 160, 255];
  await expect.poll(() => pixels(page, [
    [120, 35], [160, 35], [120, 60], [160, 60],
    [139, 36], [120, 51], [223, 60], [120, 70], [360, 60],
  ])).toEqual([ink, ink, ink, ink, divider, divider, ink, ink, [32, 32, 32, 255]]);
});
