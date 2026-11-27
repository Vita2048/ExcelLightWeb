import { expect, test, type Page } from '@playwright/test';
import axe from 'axe-core';
import { readFileSync } from 'node:fs';

interface RibbonManifest {
  readonly tabs: readonly string[];
  readonly panels: Readonly<Record<string, readonly string[]>>;
  readonly menus: readonly { readonly tab: string; readonly trigger: string; readonly items: readonly string[];
    readonly choices?: boolean; readonly textOnly?: boolean; readonly iconlessItems?: readonly string[] }[];
  readonly nested: readonly { readonly tab: string; readonly trigger: string; readonly item: string;
    readonly items: readonly string[]; readonly textOnly?: boolean; readonly iconlessItems?: readonly string[] }[];
}
const ribbon: RibbonManifest = JSON.parse(readFileSync(new URL('../fixtures/parity/ribbon-order.json', import.meta.url), 'utf8'));

const path = '/ExcelLightWeb/';
const tabs = ribbon.tabs;
const parentMenuSelector = '.ribbon-flyout[role="menu"]:not([aria-label="Hide & Unhide"]):not([aria-label="Statistical"]):not([aria-label="Information"])';

async function openTab(page: Page, tab: string) {
  await page.getByRole('tab', { name: tab, exact: true }).click();
  return page.locator('.ribbon-panel:not([hidden])');
}

function triggerButton(page: Page, trigger: string) {
  // The Number Format trigger prefixes its visible format sample in the accessible name.
  return trigger === 'Number Format'
    ? page.getByRole('button', { name: /^Number Format(?: |$)/u })
    : page.getByRole('button', { name: trigger, exact: true });
}

async function openMenu(page: Page, tab: string, trigger: string) {
  await openTab(page, tab);
  const button = triggerButton(page, trigger);
  await button.click();
  const menu = page.locator(parentMenuSelector).last();
  await expect(menu).toBeVisible();
  return { button, menu };
}

async function itemLabels(menu: ReturnType<Page['locator']>) {
  return menu.locator('[role="menuitem"], [role="menuitemcheckbox"]')
    .evaluateAll(nodes => nodes.map(node => (node.querySelector('.ribbon-flyout-label')?.firstChild?.textContent
      ?? node.textContent ?? '').trim()));
}

async function expectInsideViewport(page: Page, selector: string) {
  const box = (await page.locator(selector).boundingBox())!;
  const viewport = page.viewportSize()!;
  expect(box.x).toBeGreaterThanOrEqual(0);
  expect(box.y).toBeGreaterThanOrEqual(0);
  expect(box.x + box.width).toBeLessThanOrEqual(viewport.width + 0.5);
  expect(box.y + box.height).toBeLessThanOrEqual(viewport.height + 0.5);
}

test('ribbon tabs, groups and flyout commands follow the captured desktop order', async ({ page }) => {
  // Every tab, menu item and glyph is inspected in one journey; Firefox/WebKit need more than 30 s.
  test.setTimeout(120_000);
  await page.goto(path);
  const tree = page.getByRole('tablist', { name: 'Ribbon' });
  await expect(tree.getByRole('tab')).toHaveText([...tabs]);
  // The contextual tab only exists while the active cell is inside a table.
  await expect(page.getByRole('tab', { name: 'Table Design' })).toHaveCount(0);

  for (const [tab, groups] of Object.entries(ribbon.panels)) {
    const panel = await openTab(page, tab);
    await expect(panel.locator('> fieldset > legend')).toHaveText(groups);
    expect(await panel.locator('.ribbon-command .ribbon-glyph').count()).toBeGreaterThan(0);
    // Every flyout trigger is named, carries the disclosure chevron and opens a menu.
    for (const trigger of await panel.locator('button[aria-haspopup="menu"]').evaluateAll(nodes =>
      nodes.map(node => ({ label: node.getAttribute('aria-label') ?? '', chevron: !!node.querySelector('.ribbon-chevron') })))) {
      expect(trigger.label).not.toBe('');
      expect(trigger.chevron).toBe(true);
    }
  }
  // Split controls keep the desktop's primary command next to its options trigger.
  await openTab(page, 'Home');
  for (const [primary, options] of [['Paste', 'Paste options'], ['AutoSum', 'AutoSum options'],
    ['Merge & Center', 'Merge options'], ['Underline', 'Underline options'],
    ['Format Painter', 'Lock Format Painter']] as const) {
    const button = page.getByRole('button', { name: primary, exact: true }).first();
    expect(await button.evaluate((node, label) =>
      !!node.parentElement?.querySelector('button[aria-label="' + label + '"]'), options), primary).toBe(true);
  }

  for (const entry of ribbon.menus) {
    const { menu } = await openMenu(page, entry.tab, entry.trigger);
    expect(await itemLabels(menu)).toEqual(entry.items);
    // Checked-choice menus expose menuitemcheckbox rows; other rows show their desktop glyph.
    const rows = menu.locator(entry.choices ? '[role="menuitemcheckbox"]' : '[role="menuitem"], [role="menuitemcheckbox"]');
    for (let index = 0; index < await rows.count(); index++) {
      const row = rows.nth(index);
      const label = (await row.textContent())!.trim();
      const visual = await row.locator('.ribbon-flyout-glyph svg, .ribbon-flyout-check, .ribbon-submenu-chevron').count();
      if (entry.choices) continue;
      if (entry.textOnly) {
        if (label === 'Insert Function...') {
          expect(visual, entry.trigger + ' row ' + index).toBeGreaterThan(0);
        }
        continue;
      }
      if (entry.iconlessItems?.includes(label)) {
        expect(visual, entry.trigger + ' row ' + index).toBe(0);
        continue;
      }
      // The Number Format dialog action has no swatch, matching the desktop menu row.
      if (!label.startsWith('More Number Formats')) {
        expect(visual, entry.trigger + ' row ' + index).toBeGreaterThan(0);
      }
    }
    await expectInsideViewport(page, `${parentMenuSelector} >> nth=-1`);
    await page.keyboard.press('Escape');
    await expect(triggerButton(page, entry.trigger)).toBeFocused();
  }
});

test('nested ribbon flyout submenus open beside their parent and navigate by keyboard', async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto(path);
  for (const entry of ribbon.nested) {
    const { button, menu } = await openMenu(page, entry.tab, entry.trigger);
    const parent = (await menu.boundingBox())!;

    const triggerItem = menu.getByRole('menuitem', { name: entry.item, exact: true });
    await expect(triggerItem).toHaveAttribute('aria-haspopup', 'menu');
    await triggerItem.click();
    const child = page.getByRole('menu', { name: entry.item, exact: true });
    await expect(child).toBeVisible();
    expect(await itemLabels(child)).toEqual(entry.items);
    const rows = child.getByRole('menuitem');
    for (let index = 0; index < await rows.count(); index++) {
      const row = rows.nth(index);
      const label = (await row.textContent())!.trim();
      const visual = await row.locator('.ribbon-flyout-glyph svg, .ribbon-flyout-check, .ribbon-submenu-chevron').count();
      if (entry.textOnly) {
        if (label === 'Insert Function...') {
          expect(visual, entry.item + ' row ' + index).toBeGreaterThan(0);
        }
        continue;
      }
      if (entry.iconlessItems?.includes(label)) {
        expect(visual, entry.item + ' row ' + index).toBe(0);
      } else {
        expect(visual, entry.item + ' row ' + index).toBeGreaterThan(0);
      }
    }
    const submenu = (await child.boundingBox())!;
    const anchor = (await triggerItem.boundingBox())!;
    // The child opens beside the owning row, flipping left when the viewport has no room on the right.
    const beside = submenu.x >= anchor.x + anchor.width - 1 || submenu.x + submenu.width <= anchor.x + 1;
    expect(beside).toBe(true);
    expect(submenu.y).toBeLessThan(parent.y + parent.height);
    expect(submenu.y + submenu.height).toBeGreaterThan(parent.y);
    await expectInsideViewport(page, `.ribbon-flyout[aria-label="${entry.item}"]`);

    // Arrow keys move through the child menu; ArrowLeft closes only the submenu.
    await expect(child.getByRole('menuitem').first()).toBeFocused();
    await child.getByRole('menuitem').first().press('ArrowDown');
    await expect(child.getByRole('menuitem').nth(1)).toBeFocused();
    await page.keyboard.press('ArrowLeft');
    await expect(child).toHaveCount(0);
    await expect(triggerItem).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(menu).toHaveCount(0);
    await expect(button).toBeFocused();
  }
});

test('ribbon keeps the dark palette under a light OS preference and uses the system palette in forced colors', async ({ page }) => {
  await page.emulateMedia({ colorScheme: 'light' });
  await page.goto(path);
  const topBar = await page.evaluate(() => {
    const sample = document.createElement('span');
    sample.style.backgroundColor = getComputedStyle(document.documentElement).getPropertyValue('--top-bar').trim();
    document.body.append(sample);
    const color = getComputedStyle(sample).backgroundColor;
    sample.remove();
    return color;
  });
  await expect(page.locator('.ribbon')).toHaveCSS('background-color', topBar);
  // A light OS preference must not switch the app to a light theme.
  expect(topBar).toBe('rgb(0, 0, 0)');
  await expect(page.locator('.ribbon-panel:not([hidden])')).toHaveCSS('color', 'rgb(245, 245, 245)');

  await page.emulateMedia({ colorScheme: 'light', forcedColors: 'active' });
  expect(await page.evaluate(() => matchMedia('(forced-colors: active)').matches)).toBe(true);
  const canvas = await page.evaluate(() => {
    const sample = document.createElement('span');
    sample.style.backgroundColor = 'Canvas';
    document.body.append(sample);
    const color = getComputedStyle(sample).backgroundColor;
    sample.remove();
    return color;
  });
  await expect(page.locator('.ribbon')).toHaveCSS('background-color', canvas);
  await expect(page.locator('.ribbon-panel:not([hidden]) > fieldset > legend').first()).toBeVisible();
  const { menu } = await openMenu(page, 'Home', 'Clear');
  await expect(menu).toBeVisible();
  await expectInsideViewport(page, '.ribbon-flyout[role="menu"] >> nth=-1');
  await page.keyboard.press('Escape');
});

test('an open ribbon flyout passes axe in dark mode', async ({ page }) => {
  await page.context().addInitScript(axe.source);
  await page.goto(path);
  await openMenu(page, 'Home', 'Find & Select');
  const violations = await page.evaluate(async () => (await (window as Window & { axe?: typeof axe }).axe!.run({
    runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] },
  })).violations.map(({ id, nodes }) => ({ id, targets: nodes.map(node => node.target) })));
  expect(violations).toEqual([]);
});

test('Function Library icons are top-aligned and flyouts only scroll when exceeding viewport space', async ({ page }) => {
  await page.goto(path);
  const panel = await openTab(page, 'Formulas');
  const libraryFieldset = panel.locator('> fieldset').filter({ has: page.locator('> legend', { hasText: /^Function Library$/ }) });
  const iconTops = await libraryFieldset.locator('.ribbon-large .ribbon-glyph').evaluateAll(nodes =>
    nodes.map(node => node.getBoundingClientRect().top));
  expect(iconTops.length).toBeGreaterThanOrEqual(8);
  const baselineTop = iconTops[0]!;
  for (const top of iconTops) {
    expect(Math.abs(top - baselineTop)).toBeLessThanOrEqual(1);
  }

  const { menu: recent } = await openMenu(page, 'Formulas', 'Recently Used');
  const recentOverflow = await recent.evaluate(node => node.scrollHeight - node.clientHeight);
  expect(recentOverflow).toBeLessThanOrEqual(0);
  await page.keyboard.press('Escape');

  const { menu: lookup } = await openMenu(page, 'Formulas', 'Lookup & Reference');
  const lookupBox = (await lookup.boundingBox())!;
  expect(lookupBox.height).toBeGreaterThan(500);
  await expectInsideViewport(page, `${parentMenuSelector} >> nth=-1`);
  await page.keyboard.press('Escape');

  const { menu: math } = await openMenu(page, 'Formulas', 'Math & Trig');
  const scrollState = await math.evaluate(async node => {
    node.scrollTop = node.scrollHeight;
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    return { scrollTop: node.scrollTop, scrollHeight: node.scrollHeight, clientHeight: node.clientHeight };
  });
  expect(scrollState.scrollTop).toBeGreaterThan(100);
  expect(scrollState.scrollTop).toBeGreaterThanOrEqual(scrollState.scrollHeight - scrollState.clientHeight - 2);
  await expect(math.getByRole('menuitem', { name: 'Insert Function...', exact: true })).toBeInViewport();
  await page.keyboard.press('Escape');

  await page.setViewportSize({ width: 1280, height: 900 });
  const { menu: financial } = await openMenu(page, 'Formulas', 'Financial');
  const financialOverflow = await financial.evaluate(node => node.scrollHeight - node.clientHeight);
  expect(financialOverflow).toBeLessThanOrEqual(0);
  await expectInsideViewport(page, `${parentMenuSelector} >> nth=-1`);
  await page.keyboard.press('Escape');
});

test('contextual PivotTable Analyze and Design tabs appear only inside a PivotTable and restore the prior tab', async ({ page }) => {
  await page.context().addInitScript(axe.source);
  await page.goto(path);
  const tablist = page.getByRole('tablist', { name: 'Ribbon' });
  await expect(tablist.getByRole('tab')).toHaveText([...tabs]);

  const seed = async (addr: string, val: string) => {
    await page.getByRole('textbox', { name: 'Name box' }).fill(addr);
    await page.getByRole('textbox', { name: 'Name box' }).press('Enter');
    await page.getByRole('textbox', { name: 'Formula bar' }).fill(val);
    await page.getByRole('textbox', { name: 'Formula bar' }).press('Enter');
  };
  await seed('A1', 'Region');
  await seed('B1', 'Sales');
  await seed('A2', 'North');
  await seed('B2', '100');
  await page.getByRole('textbox', { name: 'Name box' }).fill('A1');
  await page.getByRole('textbox', { name: 'Name box' }).press('Enter');

  await openMenu(page, 'Insert', 'PivotTable');
  await page.getByRole('menuitem', { name: 'From Table/Range', exact: true }).click();
  const createDialog = page.getByRole('dialog', { name: 'PivotTable from table or range' });
  await createDialog.getByRole('radio', { name: 'Existing Worksheet' }).check();
  await createDialog.getByLabel('Location:').fill("='Sheet1'!$D$1");
  await createDialog.getByRole('button', { name: 'OK' }).click();

  await expect(tablist.getByRole('tab')).toHaveText([...tabs, 'PivotTable Analyze', 'Design']);
  const analyzeTab = page.getByRole('tab', { name: 'PivotTable Analyze', exact: true });
  await expect(analyzeTab).toHaveAttribute('aria-selected', 'true');
  await expect(analyzeTab).toHaveCSS('color', 'rgb(110, 231, 183)');
  const analyzePanel = page.locator('.ribbon-panel:not([hidden])');
  await expect(analyzePanel.locator('> fieldset > legend')).toHaveText(['PivotTable', 'Active Field', 'Data', 'Actions', 'Show']);

  const designTab = page.getByRole('tab', { name: 'Design', exact: true });
  await designTab.click();
  await expect(designTab).toHaveCSS('color', 'rgb(110, 231, 183)');
  const designPanel = page.locator('.ribbon-panel:not([hidden])');
  await expect(designPanel.locator('> fieldset > legend')).toHaveText(['Layout', 'PivotTable Style Options', 'PivotTable Styles']);
  const styleOptions = designPanel.getByRole('option');
  await expect(styleOptions).toHaveCount(4);
  const swatchHeaders = await designPanel.locator('.table-style-swatch-header').evaluateAll(nodes =>
    nodes.map(node => ({
      width: node.getBoundingClientRect().width,
      height: node.getBoundingClientRect().height,
      bg: getComputedStyle(node).backgroundColor,
    })));
  expect(swatchHeaders.every(item => item.width >= 40 && item.height >= 8)).toBe(true);
  expect(new Set(swatchHeaders.map(item => item.bg)).size).toBe(4);

  const designViolations = await page.evaluate(async () => (await (window as Window & { axe?: typeof axe }).axe!.run({
    runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'] },
  })).violations.map(({ id, nodes }) => ({ id, targets: nodes.map(node => node.target) })));
  expect(designViolations).toEqual([]);

  // Leaving the PivotTable hides contextual tabs and restores the previous non-contextual tab (Insert).
  await page.getByRole('textbox', { name: 'Name box' }).fill('A1');
  await page.getByRole('textbox', { name: 'Name box' }).press('Enter');
  await expect(tablist.getByRole('tab')).toHaveText([...tabs]);
  await expect(page.getByRole('tab', { name: 'Insert', exact: true })).toHaveAttribute('aria-selected', 'true');

  // Re-entering the same PivotTable restores the remembered contextual tab (Design).
  await page.getByRole('textbox', { name: 'Name box' }).fill('D1');
  await page.getByRole('textbox', { name: 'Name box' }).press('Enter');
  await expect(tablist.getByRole('tab')).toHaveText([...tabs, 'PivotTable Analyze', 'Design']);
  await expect(page.getByRole('tab', { name: 'Design', exact: true })).toHaveAttribute('aria-selected', 'true');

  // Table Design contextual tab shares the exact same green contextual tab accent color.
  await page.getByRole('textbox', { name: 'Name box' }).fill('F1:G2');
  await page.getByRole('textbox', { name: 'Name box' }).press('Enter');
  await seed('F1', 'Item');
  await seed('G1', 'Qty');
  await seed('F2', 'Pen');
  await seed('G2', '3');
  await page.getByRole('textbox', { name: 'Name box' }).fill('F1');
  await page.getByRole('textbox', { name: 'Name box' }).press('Enter');
  await openTab(page, 'Insert');
  await page.getByRole('button', { name: 'Table', exact: true }).click();
  await page.getByRole('dialog').getByRole('button', { name: 'OK' }).click();
  await expect(page.getByRole('tab', { name: 'Table Design', exact: true })).toHaveCSS('color', 'rgb(110, 231, 183)');
});
