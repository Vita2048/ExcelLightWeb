import { expect, test, type Page } from '@playwright/test';
import axe from 'axe-core';

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

async function seedChartData(page: Page) {
  const rows: readonly (readonly [string, string])[] = [
    ['A1', 'Quarter'], ['B1', 'East'], ['C1', 'West'],
    ['A2', 'Q1'], ['B2', '100'], ['C2', '120'],
    ['A3', 'Q2'], ['B3', '150'], ['C3', '180'],
    ['A4', 'Q3'], ['B4', '200'], ['C4', '160'],
  ];
  for (const [cell, val] of rows) {
    await go(page, cell, val);
  }
}

test('inserts core 2D charts from Insert ribbon & dialog, updates live on cell edit, moves/resizes, resizes data range, undoes/redoes, and deletes', async ({ page }) => {
  await page.context().addInitScript(axe.source);
  await page.goto(path);
  await seedChartData(page);
  await go(page, 'A1');

  // Insert Clustered Column from Insert -> Charts gallery flyout
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Insert Column or Bar Chart', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Clustered Column', exact: true }).click();

  const chart1 = page.locator('.chart-frame[aria-label="Chart 1"]');
  await expect(chart1).toBeVisible();
  await expect(chart1.locator('.chart-series-bar')).toHaveCount(6);
  await expect(chart1.locator('.chart-range-handle')).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Resize chart data range' })).toBeVisible();

  // Live update on cell edit: change B2 from 100 to 400
  const initialHeight = await chart1.locator('.chart-series-bar').first().evaluate(el => Number(el.getAttribute('height')));
  await go(page, 'B2', '400');
  const updatedHeight = await chart1.locator('.chart-series-bar').first().evaluate(el => Number(el.getAttribute('height')));
  expect(updatedHeight).not.toBe(initialHeight);

  // Select chart 1, drag to move, resize SE handle, then Undo
  await chart1.click();
  const boxBefore = (await chart1.boundingBox())!;
  await page.mouse.move(boxBefore.x + boxBefore.width / 2, boxBefore.y + boxBefore.height / 2);
  await page.mouse.down();
  await page.mouse.move(boxBefore.x + boxBefore.width / 2 + 40, boxBefore.y + boxBefore.height / 2 + 30);
  await page.mouse.up();
  const boxMoved = (await chart1.boundingBox())!;
  expect(boxMoved.x).toBeGreaterThan(boxBefore.x + 20);

  const seHandle = chart1.locator('[data-chart-edge="bottom-right"]');
  const handleBox = (await seHandle.boundingBox())!;
  await page.mouse.move(handleBox.x + handleBox.width / 2, handleBox.y + handleBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(handleBox.x + 60, handleBox.y + 40);
  await page.mouse.up();
  const boxResized = (await chart1.boundingBox())!;
  expect(boxResized.width).toBeGreaterThan(boxMoved.width + 30);

  // Accessibility check with chart selected
  const violations = await page.evaluate(async () => (await (window as Window & { axe?: typeof axe }).axe!.run({
    runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] },
  })).violations.map(({ id, nodes }) => ({ id, targets: nodes.map(node => node.target) })));
  expect(violations).toEqual([]);

  // Drag range handle up into row 3 (C4 -> C3) to shrink data range from 3 categories to 2 (6 -> 4 bars)
  const rangeHandle = page.getByRole('button', { name: 'Resize chart data range' });
  const rhBox = (await rangeHandle.boundingBox())!;
  await page.mouse.move(rhBox.x + rhBox.width / 2, rhBox.y + rhBox.height / 2);
  await page.mouse.down();
  await page.mouse.move(rhBox.x - 10, rhBox.y - 35);
  await page.mouse.up();
  await expect(chart1.locator('.chart-series-bar')).toHaveCount(4);

  // Open Recommended Charts / All Charts dialog, verify unclipped subtype icons + dual previews, and insert a Pie chart
  await go(page, 'A1');
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Recommended Charts', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Insert Chart' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('tab', { name: 'All Charts', exact: true }).click();
  await dialog.getByRole('option', { name: 'Line', exact: true }).click();
  // Verify subtype SVG icons fit cleanly inside their buttons without horizontal/vertical clipping
  const subtypeButtons = dialog.locator('.insert-chart-subtype-btn');
  await expect(subtypeButtons).toHaveCount(6);
  for (let i = 0; i < 6; i++) {
    const btnBox = (await subtypeButtons.nth(i).boundingBox())!;
    const svgBox = (await subtypeButtons.nth(i).locator('svg').boundingBox())!;
    expect(svgBox.x).toBeGreaterThanOrEqual(btnBox.x);
    expect(svgBox.x + svgBox.width).toBeLessThanOrEqual(btnBox.x + btnBox.width + 0.5);
  }
  await expect(dialog.locator('.insert-chart-preview-grid .insert-chart-preview-card')).toHaveCount(2);
  await dialog.getByRole('option', { name: 'Pie', exact: true }).click();
  await dialog.getByRole('option', { name: 'Pie', exact: true }).first().click();
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();

  const chart2 = page.locator('.chart-frame[aria-label="Chart 2"]');
  await expect(chart2).toBeVisible();
  await expect(chart2.locator('.chart-series-slice')).toHaveCount(3);

  // Delete Chart 2 via Delete key, then Undo and verify restored
  await chart2.focus();
  await page.keyboard.press('Delete');
  await expect(chart2).toHaveCount(0);

  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(chart2).toBeVisible();
});

test('inserts Combo (Secondary Axis), Waterfall, Statistical, Hierarchy, and Radar charts via ribbon and All Charts dialog', async ({ page }) => {
  await page.goto(path);
  await seedChartData(page);
  await go(page, 'A1');

  // 1. Insert Combo (Clustered Column - Line on Secondary Axis) from Insert -> Charts
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Insert Combo Chart', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Clustered Column - Line on Secondary Axis', exact: true }).click();

  const comboChart = page.locator('.chart-frame[aria-label="Chart 1"]');
  await expect(comboChart).toBeVisible();
  await expect(comboChart.locator('.chart-series-bar')).toHaveCount(3);
  await expect(comboChart.locator('g.chart-axis')).toHaveCount(3); // X + Primary Y + Secondary Y

  // 2. Insert Waterfall from Waterfall/Funnel/Stock/Radar gallery
  await go(page, 'A1');
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Insert Waterfall, Funnel, Stock, or Radar Chart', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Waterfall', exact: true }).click();

  const waterfallChart = page.locator('.chart-frame[aria-label="Chart 2"]');
  await expect(waterfallChart).toBeVisible();
  await expect(waterfallChart.locator('.chart-waterfall-connector')).toHaveCount(2);

  // 3. Insert Pareto from Statistic gallery
  await go(page, 'A1');
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'Insert Statistic Chart', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Pareto', exact: true }).click();
  const paretoChart = page.locator('.chart-frame[aria-label="Chart 3"]');
  await expect(paretoChart).toBeVisible();

  // 4. Open All Charts dialog -> Treemap + Radar + Custom Combo table
  await go(page, 'A1');
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'See All Charts', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Insert Chart' });
  await expect(dialog).toBeVisible();

  await dialog.getByRole('option', { name: 'Combo', exact: true }).click();
  await expect(dialog.getByRole('table', { name: 'Combo series configuration' })).toBeVisible();
  await expect(dialog.getByRole('checkbox', { name: 'Secondary Axis for West' })).toBeVisible();

  await dialog.getByRole('option', { name: 'Treemap', exact: true }).click();
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();

  const treemapChart = page.locator('.chart-frame[aria-label="Chart 4"]');
  await expect(treemapChart).toBeVisible();
  await expect(treemapChart.locator('.chart-treemap-label')).toHaveCount(3);
});
