import { expect, test, type Page } from '@playwright/test';

const path = '/ExcelLightWeb/';

async function go(page: Page, address: string) {
  const name = page.getByRole('textbox', { name: 'Name box' });
  await name.fill(address);
  await name.press('Enter');
}

async function enter(page: Page, address: string, source: string) {
  await go(page, address);
  const editor = page.getByRole('textbox', { name: 'Formula bar' });
  await editor.fill(source);
  await editor.press('Enter');
}

function cell(page: Page, address: string, value: string) {
  return page.getByRole('gridcell', { name: address + ', ' + value, exact: true });
}

test('Manual preserves caches and exposes stale cells, Calculate Sheet and Now honor their scope, and Automatic resumes', async ({ page }) => {
  await page.goto(path);
  await enter(page, 'A1', '2');
  await enter(page, 'B1', '=A1+3');
  await expect(cell(page, 'B1', '5')).toBeAttached();
  await page.getByRole('button', { name: 'New sheet' }).click();
  await enter(page, 'A1', '=Sheet1!B1+5');
  await expect(cell(page, 'A1', '10')).toBeAttached();
  await enter(page, 'B1', '=20+1');
  await expect(cell(page, 'B1', '21')).toBeAttached();
  await page.getByRole('tab', { name: 'Formulas' }).click();
  const mode = page.getByRole('button', { name: 'Calculation Options' });
  const choice = (name: string) => page.getByRole('menuitemcheckbox', { name, exact: true });
  await mode.click();
  await expect(choice('Automatic')).toBeChecked();
  await choice('Manual').click();
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click();
  await enter(page, 'A1', '10');
  await expect(cell(page, 'B1', '5')).toHaveAccessibleDescription('Calculation needed');
  await expect(page.getByRole('status').filter({ hasText: 'Calculation needed' })).toBeVisible();

  await page.getByRole('button', { name: 'Calculate Sheet' }).click();
  await expect(cell(page, 'B1', '13')).toHaveAccessibleDescription('');
  await page.getByRole('button', { name: 'Sheet2', exact: true }).click();
  await expect(cell(page, 'A1', '10')).toHaveAccessibleDescription('Calculation needed');
  await expect(cell(page, 'B1', '21')).toHaveAccessibleDescription('');
  await page.getByRole('button', { name: 'Calculate Now' }).click();
  await expect(cell(page, 'A1', '18')).toHaveAccessibleDescription('');
  await mode.click();
  await expect(choice('Manual')).toBeChecked();
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Sheet1', exact: true }).click();
  await enter(page, 'A1', '11');
  await expect(cell(page, 'B1', '13')).toHaveAccessibleDescription('Calculation needed');
  await enter(page, 'B1', '=A1+4');
  await expect(cell(page, 'B1', '13')).toHaveAccessibleDescription('Calculation needed');
  await mode.click();
  await choice('Automatic').click();
  await expect(cell(page, 'B1', '15')).toHaveAccessibleDescription('');
  await page.getByRole('tab', { name: 'View' }).click();
  await page.getByRole('tab', { name: 'Formulas' }).click();
  await mode.click();
  await expect(choice('Automatic')).toBeChecked();
});

test.describe('localized formula commands', () => {
  test.use({ locale: 'cs-CZ' });

  test('Insert Function opens an editable draft and Show Formulas stays in sync with View', async ({ page }) => {
    await page.goto(path);
    await go(page, 'A1');
    await page.getByRole('tab', { name: 'Formulas' }).click();
    const picker = page.getByRole('button', { name: 'Insert Function', exact: true });
    await picker.click();
    const dialog = page.getByRole('dialog', { name: 'Insert Function' });
    await expect(dialog).toBeVisible();
    await dialog.getByLabel('Or select a category:').selectOption('All');
    const list = dialog.getByRole('listbox');
    await expect(list.getByRole('option')).toHaveCount(225);
    await dialog.getByLabel('Or select a category:').selectOption('Logical');
    await list.getByRole('option', { name: /IF\s+Logical/u }).click();
    await expect(dialog).toContainText('IF(logical_test; [value_if_true]; [value_if_false])');
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
    await expect(dialog).toHaveCount(0);

    const argsDialog = page.getByRole('dialog', { name: 'Function Arguments' });
    await expect(argsDialog).toBeVisible();
    await expect(argsDialog.locator('legend')).toHaveText('IF');
    await argsDialog.getByLabel('logical_test', { exact: true }).fill('TRUE');
    await argsDialog.getByLabel('value_if_true', { exact: true }).fill('"x,y;z"');
    await argsDialog.getByLabel('value_if_false', { exact: true }).fill('"other"');
    await expect(argsDialog.locator('.function-args-result')).toContainText('Formula result = x,y;z');
    await argsDialog.getByRole('button', { name: 'OK', exact: true }).click();
    await expect(argsDialog).toHaveCount(0);
    await expect(cell(page, 'A1', 'x,y;z')).toBeAttached();

    await go(page, 'B1');
    const editor = page.getByRole('textbox', { name: 'Formula bar' });
    const hint = page.getByRole('tooltip');
    await editor.click();
    await editor.fill('=IF()');
    await editor.press('ArrowLeft');
    await expect(hint.locator('strong')).toHaveText('logical_test');
    await page.keyboard.insertText('TRUE;');
    await expect(hint).toContainText('logical_test; [value_if_true]; [value_if_false]');
    await expect(hint.locator('strong')).toHaveText('[value_if_true]');
    await editor.press('Escape');

    await enter(page, 'C1', '10');
    await enter(page, 'C2', '20');
    await go(page, 'D1');
    await page.getByRole('button', { name: 'Math & Trig', exact: true }).click();
    await page.getByRole('menuitem', { name: 'SUM', exact: true }).click();
    await expect(argsDialog).toBeVisible();
    await expect(argsDialog.locator('legend')).toHaveText('SUM');
    await argsDialog.getByRole('button', { name: 'Pick range for number1', exact: true }).click();
    await expect(argsDialog).toHaveCount(0);
    const rangeInput = page.getByRole('textbox', { name: 'Selected worksheet range' });
    await rangeInput.fill('=Sheet1!$C$1:$C$2');
    await page.locator('.range-picker-floating-panel').getByRole('button', { name: 'OK', exact: true }).click();
    await expect(argsDialog).toBeVisible();
    await expect(argsDialog.getByLabel('number1', { exact: true })).toHaveValue('C1:C2');
    await expect(argsDialog.locator('.function-args-result')).toContainText('Formula result = 30');
    await argsDialog.getByRole('button', { name: 'OK', exact: true }).click();
    await expect(cell(page, 'D1', '30')).toBeAttached();

    await go(page, 'D1');
    const grid = page.getByRole('grid');
    await grid.focus();
    await grid.press('F2');
    const inlineD1 = page.getByRole('textbox', { name: 'Edit D1' });
    await expect(inlineD1).toBeFocused();
    await expect(grid).toHaveAttribute('data-formula-highlights', '1');
    await expect(page.locator('.inline-cell-editor-wrap .formula-highlight-token')).toHaveText(['C1:C2']);
    await inlineD1.press('Escape');
    await expect(grid).not.toHaveAttribute('data-formula-highlights');

    await enter(page, 'E1', '=SUM(A1:A3;B2:B4;C1:C2)');
    await go(page, 'E1');
    await grid.focus();
    await grid.press('F2');
    const tokens = page.locator('.inline-cell-editor-wrap .formula-highlight-token');
    await expect(tokens).toHaveText(['A1:A3', 'B2:B4', 'C1:C2']);
    const colors = await tokens.evaluateAll(nodes => nodes.map(n => getComputedStyle(n).color));
    expect(colors[0]).not.toBe(colors[1]);
    await expect(grid).toHaveAttribute('data-formula-highlights', '3');
    const inlineE1 = page.getByRole('textbox', { name: 'Edit E1' });
    const inlineBox = (await page.locator('.inline-cell-editor').boundingBox())!;
    expect(inlineBox.width).toBeGreaterThan(96);
    expect(await inlineE1.evaluate(node => node.scrollWidth <= node.clientWidth + 2)).toBe(true);
    await inlineE1.press('Escape');
    await expect(grid).not.toHaveAttribute('data-formula-highlights');

    await go(page, 'A1');
    const audit = page.getByRole('button', { name: 'Show Formulas' });
    await audit.click();
    await expect(audit).toHaveAttribute('aria-pressed', 'true');
    await expect(cell(page, 'A1', '=IF(TRUE;"x,y;z";"other")')).toBeAttached();
    await page.getByRole('tab', { name: 'View' }).click();
    const view = page.getByRole('checkbox', { name: 'Show Formulas' });
    await expect(view).toBeChecked();
    await view.uncheck();
    await expect(cell(page, 'A1', 'x,y;z')).toBeAttached();
    await page.getByRole('checkbox', { name: 'Formula Bar' }).uncheck();
    await page.getByRole('tab', { name: 'Formulas' }).click();
    await expect(audit).toHaveAttribute('aria-pressed', 'false');
    await goInlineTarget(page, 'A1');
    await page.getByRole('button', { name: 'Math & Trig', exact: true }).click();
    await page.getByRole('menuitem', { name: 'SUM', exact: true }).click();
    await expect(argsDialog).toBeVisible();
    await argsDialog.getByLabel('number1', { exact: true }).fill('1');
    await argsDialog.getByLabel('number2', { exact: true }).fill('2');
    await expect(argsDialog.getByLabel('number3', { exact: true })).toBeVisible();
    await expect(argsDialog.locator('.function-args-result')).toContainText('Formula result = 3');
    await argsDialog.getByRole('button', { name: 'Cancel', exact: true }).click();
    await expect(page.getByRole('grid')).toBeFocused();
    await expect(cell(page, 'A1', 'x,y;z')).toBeAttached();

    await page.getByRole('button', { name: 'Recently Used', exact: true }).click();
    const recentItems = page.getByRole('menu', { name: 'Recently Used' }).getByRole('menuitem');
    await expect(recentItems.nth(0)).toHaveText('SUM');
    await expect(recentItems.nth(1)).toHaveText('IF');
    await page.keyboard.press('Escape');

    await page.getByRole('button', { name: 'Financial', exact: true }).click();
    await page.getByRole('menuitem', { name: 'Insert Function...', exact: true }).click();
    await expect(dialog).toBeVisible();
    await expect(dialog.getByLabel('Or select a category:')).toHaveValue('Financial');
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click();
  });
});

async function goInlineTarget(page: Page, address: 'A1') {
  const grid = page.getByRole('grid');
  await grid.focus();
  await grid.press('Control+Home');
  await expect(grid).toHaveAttribute('aria-activedescendant', new RegExp('-' + address + '$', 'i'));
}

interface WorkerProbe {
  readonly urls: string[];
  readonly sent: number[];
  readonly finished: { elapsed: number; updates: number }[];
  readonly held: (() => void)[];
  replaying: boolean;
}
type ProbedWindow = Window & { probe: WorkerProbe };

test('the production worker calculates a 1,000-formula chain and drops a reply superseded by worksheet input', async ({ page, browser }, testInfo) => {
  await page.addInitScript(() => {
    const probe: WorkerProbe = { urls: [], sent: [], finished: [], held: [], replaying: false };
    (window as unknown as ProbedWindow).probe = probe;
    const BrowserWorker = window.Worker;
    const nativeWorkers = new WeakSet<Worker>();
    window.Worker = class extends BrowserWorker {
      constructor(url: URL | string, options?: WorkerOptions) {
        super(url, options);
        const nativeWorker = String(url).includes('/native-calculation.worker-');
        if (!nativeWorker) return;
        nativeWorkers.add(this);
        probe.urls.push(String(url));
        this.addEventListener('message', event => {
          if (probe.replaying || event.data.kind !== 'native-calculated') return;
          const start = probe.sent[probe.finished.length]!;
          probe.finished.push({ elapsed: performance.now() - start, updates: event.data.updates.length });
          event.stopImmediatePropagation();
          probe.held.push(() => {
            probe.replaying = true;
            this.dispatchEvent(new MessageEvent('message', { data: event.data }));
            probe.replaying = false;
          });
        });
      }

      postMessage(message: unknown) {
        if (nativeWorkers.has(this)) probe.sent.push(performance.now());
        super.postMessage(message);
      }
    };
  });
  await page.goto(path);
  const grid = page.getByRole('grid');
  const fixture = Array.from({ length: 1000 }, (_, index) => index === 0 ? '1\t=A1+1' : `\t=B${index}+1`).join('\n');
  await grid.evaluate((node, text) => {
    const clipboard = new DataTransfer();
    clipboard.setData('text/plain', text);
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboard });
    node.dispatchEvent(event);
  }, fixture);
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbedWindow).probe.sent.length)).toBe(1);
  await go(page, 'B1000');
  await expect(cell(page, 'B1000', '=B999+1')).toBeAttached();
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbedWindow).probe.held.length)).toBe(1);
  await expect(page.locator('[role="gridcell"][id$="-B1000"]')).toHaveAccessibleDescription('Calculating');
  await page.evaluate(() => (window as unknown as ProbedWindow).probe.held.shift()!());
  await expect(cell(page, 'B1000', '1001')).toBeAttached();
  await expect(page.locator('[role="gridcell"][id$="-B1000"]')).toHaveAccessibleDescription('');
  await enter(page, 'A1', '10');
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbedWindow).probe.sent.length)).toBe(2);
  await enter(page, 'A1', '20');
  await expect(cell(page, 'A1', '20')).toBeAttached();
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbedWindow).probe.held.length)).toBe(1);
  await page.evaluate(() => (window as unknown as ProbedWindow).probe.held.shift()!());
  await expect(cell(page, 'B1', '2')).toBeAttached();
  await expect.poll(() => page.evaluate(() => (window as unknown as ProbedWindow).probe.held.length)).toBe(1);
  await page.evaluate(() => (window as unknown as ProbedWindow).probe.held.shift()!());
  await expect(cell(page, 'B1', '21')).toBeAttached();
  await go(page, 'B1000');
  await expect(cell(page, 'B1000', '1020')).toBeAttached();
  const metrics = await page.evaluate(() => {
    const { urls, finished } = (window as unknown as ProbedWindow).probe;
    return { urls, finished };
  });
  expect(metrics.urls).toHaveLength(1);
  expect(metrics.urls[0]).toMatch(/\/ExcelLightWeb\/assets\/native-calculation\.worker-/u);
  expect(metrics.finished.map(result => result.updates)).toEqual([1000, 1000, 1000]);
  await testInfo.attach('native-worker-timings', {
    body: JSON.stringify({ browser: browser.version(), formulas: 1000, ...metrics }), contentType: 'application/json',
  });
});
