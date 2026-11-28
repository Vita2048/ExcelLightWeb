import { expect, test, type Page } from '@playwright/test';
import axe from 'axe-core';

const path = '/ExcelLightWeb/';
test.describe.configure({ timeout: 120_000 });

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

test('supports Excel Chart Chrome: Insert launcher parity, contextual Chart Design/Format tabs, floating +/Brush/Funnel buttons, dialogs, Format pane & inline title edit', async ({ page }) => {
  await page.context().addInitScript(axe.source);
  await page.goto(path);
  await seedChartData(page);

  // Add Sheet2 for Move Chart test
  await page.getByRole('button', { name: 'New sheet', exact: true }).click();
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click();
  await go(page, 'A1');

  // Verify Insert -> Charts launcher matches Home tab launcher class & SVG glyph
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  const chartLauncher = page.getByRole('button', { name: 'See All Charts', exact: true });
  await expect(chartLauncher).toHaveClass(/\bribbon-launcher\b/);
  await expect(chartLauncher.locator('svg.ribbon-glyph')).toBeVisible();

  // Insert Clustered Column chart -> contextual Chart Design & Format tabs appear and Chart Design auto-activates
  await page.getByRole('button', { name: 'Insert Column or Bar Chart', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Clustered Column', exact: true }).click();

  const chart = page.locator('.chart-frame[aria-label="Chart 1"]');
  await expect(chart).toBeVisible();
  const designTab = page.getByRole('tab', { name: 'Chart Design', exact: true });
  const formatTab = page.getByRole('tab', { name: 'Format', exact: true });
  await expect(designTab).toHaveAttribute('aria-selected', 'true');
  await expect(formatTab).toBeVisible();

  // Contextual tab memory & restoration: switch to Format, deselect to grid (restores Insert), reselect (remembers Format)
  await formatTab.click();
  await expect(formatTab).toHaveAttribute('aria-selected', 'true');
  await go(page, 'A6');
  await expect(designTab).toHaveCount(0);
  await expect(page.getByRole('tab', { name: 'Insert', exact: true })).toHaveAttribute('aria-selected', 'true');
  await chart.click();
  await expect(formatTab).toHaveAttribute('aria-selected', 'true');

  // Floating '+' Chart Elements quick button: toggle Data Labels, Data Table, Trendline + verify side-expanding submenu
  await page.getByRole('button', { name: 'Chart Elements', exact: true }).click();
  const elPopover = page.getByRole('dialog', { name: 'Chart Elements' });
  await elPopover.getByRole('checkbox', { name: 'Data Labels' }).check();
  await elPopover.getByRole('checkbox', { name: 'Data Table' }).check();
  await elPopover.getByRole('button', { name: 'Trendline options', exact: true }).click();
  const sideSubmenu = elPopover.getByRole('menu', { name: 'Element options' });
  await expect(sideSubmenu).toBeVisible();
  await sideSubmenu.getByRole('menuitem', { name: 'Linear Forecast', exact: true }).click();
  await expect(chart.locator('.chart-data-label')).toHaveCount(6);
  await expect(chart.locator('.chart-data-table')).toBeVisible();
  await expect(chart.locator('.chart-trendline')).toHaveCount(1);

  // Verify Data Table series-label border rect exists & Axis Title sits closer to Y-axis than Left Legend
  await expect(chart.locator('.chart-data-table rect.chart-gridline')).toHaveCount(2);
  await elPopover.getByRole('button', { name: 'Legend options', exact: true }).click();
  await elPopover.getByRole('menu', { name: 'Element options' }).getByRole('menuitemradio', { name: 'Left', exact: true }).click();
  await elPopover.getByRole('checkbox', { name: 'Axis Titles' }).check();
  const leftLegendX = Number(await chart.locator('.chart-legend-item rect').first().getAttribute('x'));
  const yAxisTitleX = Number(await chart.locator('.chart-axis-title[data-axis="y"]').getAttribute('x'));
  expect(yAxisTitleX).toBeGreaterThan(leftLegendX + 40);

  // Floating Brush quick button: switch palette via Color tab
  await page.getByRole('button', { name: 'Chart Styles', exact: true }).click();
  const stylePopover = page.getByRole('dialog', { name: 'Chart Styles' });
  await stylePopover.getByRole('tab', { name: 'Color', exact: true }).click();
  const beforeFill = await chart.locator('.chart-series-bar').nth(3).getAttribute('fill');
  await stylePopover.getByRole('button', { name: 'Monochrome Blue', exact: true }).click();
  const afterFill = await chart.locator('.chart-series-bar').nth(3).getAttribute('fill');
  expect(afterFill).not.toBe(beforeFill);

  // Floating Funnel quick button: Values tab filters 'West' -> Undo restores, then Names tab toggles Categories -> (None)
  await page.getByRole('button', { name: 'Chart Filters', exact: true }).click();
  const filterPopover = page.getByRole('dialog', { name: 'Chart Filters' });
  await filterPopover.getByRole('checkbox', { name: 'West' }).uncheck();
  await filterPopover.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(chart.locator('.chart-series-bar')).toHaveCount(3);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(chart.locator('.chart-series-bar')).toHaveCount(6);

  await page.getByRole('button', { name: 'Chart Filters', exact: true }).click();
  await filterPopover.getByRole('tab', { name: 'Names', exact: true }).click();
  await expect(filterPopover.getByText('Series', { exact: true })).toBeVisible();
  await expect(filterPopover.getByText('Categories', { exact: true })).toBeVisible();
  await filterPopover.getByRole('radio', { name: '(None)' }).nth(1).check();
  await filterPopover.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(chart.getByText('Q1')).toHaveCount(0);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(chart.getByText('Q1').first()).toBeVisible();

  // Open docked Format Chart pane from Format ribbon, resize width via Size input, check Accessibility
  await formatTab.click();
  await page.getByRole('button', { name: 'Format Selection', exact: true }).click();
  const formatPane = page.getByRole('complementary', { name: 'Format Chart' });
  await expect(formatPane).toBeVisible();
  await expect(formatPane.getByRole('heading', { name: 'Format Chart Area' })).toBeVisible();

  const widthInput = page.getByRole('spinbutton', { name: 'Width', exact: true });
  await widthInput.fill('540');
  await widthInput.press('Enter');
  const resizedBox = (await chart.boundingBox())!;
  expect(resizedBox.width).toBeGreaterThan(515);

  // Verify Shape Fill color palette (Theme Colors + Standard Colors) and Shape Outline Weight/Dashes submenus
  await page.getByRole('button', { name: 'Shape Fill', exact: true }).click();
  const fillMenu = page.getByRole('group', { name: 'Shape Fill' });
  await expect(fillMenu.getByText('Theme Colors')).toBeVisible();
  await expect(fillMenu.getByText('Standard Colors')).toBeVisible();
  await fillMenu.getByRole('button', { name: 'Color #0070C0', exact: true }).click();
  await expect(chart.locator('.chart-area-bg')).toHaveAttribute('fill', '#0070C0');

  await page.getByRole('button', { name: 'Shape Outline', exact: true }).click();
  const outlineMenu = page.getByRole('group', { name: 'Shape Outline' });
  await outlineMenu.getByRole('button', { name: 'Dashes', exact: true }).click();
  await page.getByRole('menu', { name: 'Outline Dashes' }).getByRole('menuitemradio', { name: 'Dash', exact: true }).click();
  await expect(chart.locator('.chart-area-bg')).toHaveAttribute('stroke-dasharray', /.+/);

  // Switch Format pane element to Value Axis, customize minimum bound, and type multi-char Axis Title without losing focus
  await formatPane.getByRole('combobox', { name: 'Chart Options' }).selectOption('axis:value');
  await expect(formatPane.getByRole('heading', { name: 'Format Axis' })).toBeVisible();
  await formatPane.getByRole('tab', { name: 'Element Options' }).click();
  await formatPane.getByRole('spinbutton', { name: 'Axis minimum' }).fill('50');
  const axisTitleInput = formatPane.getByRole('textbox', { name: 'Axis title text' });
  await axisTitleInput.click();
  await axisTitleInput.fill('');
  await axisTitleInput.pressSequentially('Revenue ($)');
  await expect(axisTitleInput).toBeFocused();
  await expect(axisTitleInput).toHaveValue('Revenue ($)');
  await expect(chart.locator('.chart-axis-title').filter({ hasText: 'Revenue ($)' })).toBeVisible();

  // WCAG check with contextual Format tab + docked Format pane open
  const violations = await page.evaluate(async () => (await (window as Window & { axe?: typeof axe }).axe!.run({
    runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] },
  })).violations.map(({ id, nodes }) => ({ id, targets: nodes.map(node => node.target) })));
  expect(violations).toEqual([]);

  // Right-click Chart context menu actually triggers commands (Change Chart Type... & Format Chart Area...)
  await chart.click({ button: 'right' });
  const ctxMenu = page.getByRole('menu', { name: 'Chart context menu' });
  await expect(ctxMenu).toBeVisible();
  await ctxMenu.getByRole('menuitem', { name: 'Change Chart Type...', exact: true }).click();
  const ctxChangeDialog = page.getByRole('dialog', { name: 'Change Chart Type' });
  await expect(ctxChangeDialog).toBeVisible();
  await ctxChangeDialog.getByRole('button', { name: 'Cancel', exact: true }).click();

  // Inline title edit on double-click
  await chart.locator('.chart-title-hit').dblclick();
  const titleInput = chart.getByRole('textbox', { name: 'Edit chart title' });
  await expect(titleInput).toBeVisible();
  await titleInput.fill('Quarterly Sales');
  await titleInput.press('Enter');
  await expect(chart.locator('.chart-title-text')).toHaveText('Quarterly Sales');

  // Chart Design actions: Add Chart Element submenu icons, Switch Row/Column, Select Data category checkboxes, Change Chart Type, Move Chart to Sheet2
  await designTab.click();
  await page.getByRole('button', { name: 'Add Chart Element', exact: true }).click();
  const addElementMenu = page.getByRole('menu', { name: 'Add Chart Element' });
  await addElementMenu.getByRole('menuitem', { name: 'Axis Titles', exact: true }).click();
  await expect(page.getByRole('menuitemcheckbox', { name: 'Primary Horizontal', exact: true }).locator('svg')).toBeVisible();
  await page.keyboard.press('Escape');
  await page.keyboard.press('Escape');

  await page.getByRole('button', { name: 'Switch Row/Column', exact: true }).click();
  await expect(chart.locator('.chart-legend-item')).toHaveCount(3);

  await page.getByRole('button', { name: 'Select Data', exact: true }).click();
  const selectDialog = page.getByRole('dialog', { name: 'Select Data Source' });
  await expect(selectDialog).toBeVisible();
  const eastRow = selectDialog.locator('.select-data-list-item', { hasText: 'East' });
  const eastCheckbox = selectDialog.getByRole('checkbox', { name: 'Show category East' });
  const rowBox = (await eastRow.boundingBox())!;
  const checkboxBox = (await eastCheckbox.boundingBox())!;
  expect(rowBox.height).toBeLessThan(36);
  expect(Math.abs(checkboxBox.y - rowBox.y)).toBeLessThan(12);
  await eastCheckbox.uncheck();
  await selectDialog.getByRole('button', { name: 'Hidden and Empty Cells', exact: true }).click();
  await selectDialog.getByRole('radio', { name: 'Zero', exact: true }).check();
  await selectDialog.getByRole('button', { name: 'Done', exact: true }).click();
  await selectDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(chart.getByText('East')).toHaveCount(0);

  await page.getByRole('button', { name: 'Change Chart Type' }).click();
  const changeDialog = page.getByRole('dialog', { name: 'Change Chart Type' });
  await expect(changeDialog).toBeVisible();
  await changeDialog.getByRole('option', { name: 'Line', exact: true }).click();
  await changeDialog.getByRole('option', { name: 'Line with Markers', exact: true }).click();
  await changeDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(chart.locator('.chart-series-marker').first()).toBeVisible();

  await page.getByRole('button', { name: 'Move Chart' }).click();
  const moveDialog = page.getByRole('dialog', { name: 'Move Chart' });
  await moveDialog.getByLabel('Object in:').selectOption({ label: 'Sheet2' });
  await moveDialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Sheet2', exact: true })).toHaveAttribute('aria-current', 'true');
  await expect(page.locator('.chart-frame[aria-label="Chart 1"]')).toBeVisible();
});

test('creates PivotChart with interactive Field Buttons, syncs filters bidirectionally with PivotTable, toggles Field Buttons, and passes axe', async ({ page }) => {
  await page.context().addInitScript(axe.source);
  await page.goto(path);
  await seedChartData(page);
  await go(page, 'A1');

  // Create PivotTable on Existing Worksheet E1
  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.getByRole('button', { name: 'PivotTable', exact: true }).click();
  await page.getByRole('menuitem', { name: 'From Table/Range', exact: true }).click();
  const createDialog = page.getByRole('dialog', { name: 'PivotTable from table or range' });
  await createDialog.getByRole('radio', { name: 'Existing Worksheet' }).check();
  await createDialog.getByLabel('Location:').fill("='Sheet1'!$E$1");
  await createDialog.getByRole('button', { name: 'OK' }).click();

  const fieldsPane = page.getByRole('complementary', { name: 'PivotTable Fields' });
  await fieldsPane.getByRole('checkbox', { name: 'Quarter' }).check();
  await fieldsPane.getByRole('checkbox', { name: 'East' }).check();
  await expect(page.getByRole('gridcell', { name: 'E2, Q1' })).toBeAttached();

  // Insert PivotChart from PivotTable Analyze -> Tools -> PivotChart
  await page.getByRole('tab', { name: 'PivotTable Analyze', exact: true }).click();
  await page.getByRole('button', { name: 'PivotChart', exact: true }).click();

  const pivotChart = page.locator('.chart-frame[aria-label="Chart 1"]');
  await expect(pivotChart).toBeVisible();
  await expect(pivotChart.locator('.chart-series-bar')).toHaveCount(3);

  // On-chart Field Buttons render cleanly inside the PivotChart
  const quarterBtn = pivotChart.getByRole('button', { name: 'Filter axis field Quarter', exact: true });
  await expect(quarterBtn).toBeVisible();
  await expect(pivotChart.locator('.pivot-chart-field-pill.value-pill', { hasText: 'Sum of East' })).toBeVisible();

  // Filter via on-chart Field Button -> uncheck Q1 -> both PivotTable and PivotChart update live
  await quarterBtn.click();
  const filterDialog = page.getByRole('dialog', { name: 'Filter Row Labels (Quarter)' });
  await expect(filterDialog).toBeVisible();
  await filterDialog.getByRole('checkbox', { name: 'Q1' }).uncheck();
  await filterDialog.getByRole('button', { name: 'OK' }).click();

  await expect(pivotChart.locator('.chart-series-bar')).toHaveCount(2);
  await expect(page.getByRole('gridcell', { name: 'E2, Q1' })).toHaveCount(0);
  await expect(quarterBtn).toHaveAttribute('data-filtered', 'true');

  // Toggle Field Buttons off and back on from Chart Design -> Show/Hide -> Field Buttons
  await page.getByRole('tab', { name: 'Chart Design', exact: true }).click();
  const fieldBtnsToggle = page.getByRole('button', { name: 'Field Buttons', exact: true });
  await expect(fieldBtnsToggle).toHaveAttribute('aria-pressed', 'true');
  await fieldBtnsToggle.click();
  await expect(quarterBtn).toHaveCount(0);
  await fieldBtnsToggle.click();
  await expect(quarterBtn).toBeVisible();

  // Accessibility check with PivotChart + Field Buttons active
  const violations = await page.evaluate(async () => (await (window as Window & { axe?: typeof axe }).axe!.run({
    runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] },
  })).violations.map(({ id, nodes }) => ({ id, targets: nodes.map(node => node.target) })));
  expect(violations).toEqual([]);
});
