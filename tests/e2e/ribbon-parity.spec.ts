import { expect, test, type Page } from '@playwright/test';
import axe from 'axe-core';
import { readFileSync } from 'node:fs';

interface RibbonManifest {
  readonly tabs: readonly string[];
  readonly panels: Readonly<Record<string, readonly string[]>>;
  readonly menus: readonly { readonly tab: string; readonly trigger: string; readonly items: readonly string[];
    readonly choices?: boolean; readonly textOnly?: boolean }[];
  readonly nested: readonly { readonly tab: string; readonly trigger: string; readonly item: string;
    readonly items: readonly string[] }[];
}
const ribbon: RibbonManifest = JSON.parse(readFileSync(new URL('../fixtures/parity/ribbon-order.json', import.meta.url), 'utf8'));

const path = '/ExcelLightWeb/';
const tabs = ribbon.tabs;

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
  const menu = page.locator('.ribbon-flyout[role="menu"]:not([aria-label="Hide & Unhide"])').last();
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
      if (entry.choices || entry.textOnly) continue;
      const visual = await row.locator('.ribbon-flyout-glyph svg, .ribbon-flyout-check, .ribbon-submenu-chevron').count();
      // The Number Format dialog action has no swatch, matching the desktop menu row.
      if (!(await row.textContent())!.startsWith('More Number Formats')) {
        expect(visual, entry.trigger + ' row ' + index).toBeGreaterThan(0);
      }
    }
    await expectInsideViewport(page, '.ribbon-flyout[role="menu"]:not([aria-label="Hide & Unhide"]) >> nth=-1');
    await page.keyboard.press('Escape');
    await expect(triggerButton(page, entry.trigger)).toBeFocused();
  }
});

test('nested Format > Hide & Unhide opens beside its parent and navigates by keyboard', async ({ page }) => {
  await page.goto(path);
  const { button, menu } = await openMenu(page, 'Home', 'Format');
  const parent = (await menu.boundingBox())!;

  const hide = menu.getByRole('menuitem', { name: 'Hide & Unhide', exact: true });
  await expect(hide).toHaveAttribute('aria-haspopup', 'menu');
  await hide.click();
  const child = page.getByRole('menu', { name: 'Hide & Unhide', exact: true });
  await expect(child).toBeVisible();
  expect(await itemLabels(child)).toEqual(ribbon.nested[0]!.items);
  const submenu = (await child.boundingBox())!;
  const anchor = (await hide.boundingBox())!;
  // The child opens beside the owning row, flipping left when the viewport has no room on the right.
  const beside = submenu.x >= anchor.x + anchor.width - 1 || submenu.x + submenu.width <= anchor.x + 1;
  expect(beside).toBe(true);
  expect(submenu.y).toBeLessThan(parent.y + parent.height);
  expect(submenu.y + submenu.height).toBeGreaterThan(parent.y);
  await expectInsideViewport(page, '.ribbon-flyout[aria-label="Hide & Unhide"]');

  // Arrow keys move through the child menu; ArrowLeft closes only the submenu.
  await expect(child.getByRole('menuitem').first()).toBeFocused();
  await child.getByRole('menuitem').first().press('ArrowDown');
  await expect(child.getByRole('menuitem').nth(1)).toBeFocused();
  await page.keyboard.press('ArrowLeft');
  await expect(child).toHaveCount(0);
  await expect(hide).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(menu).toHaveCount(0);
  await expect(button).toBeFocused();
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
