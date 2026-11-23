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
    const picker = page.getByRole('button', { name: 'Insert Function' });
    await picker.click();
    await expect(page.getByRole('menu', { name: 'Insert Function' }).getByRole('menuitem')).toHaveCount(31);
    await page.getByRole('menuitem', { name: 'IF(…)', exact: true }).click();
    const editor = page.getByRole('textbox', { name: 'Formula bar' });
    const hint = page.getByRole('tooltip');
    await expect(editor).toBeFocused();
    await expect(editor).toHaveValue('=IF()');
    expect(await editor.evaluate(node => (node as HTMLTextAreaElement).selectionStart)).toBe(4);
    await expect(hint.locator('strong')).toHaveText('logical_test');
    await page.keyboard.insertText('TRUE;');
    await expect(hint).toContainText('logical_test; [value_if_true]; [value_if_false]');
    await expect(hint.locator('strong')).toHaveText('[value_if_true]');
    await page.keyboard.insertText('"x,y;z";"other"');
    await expect(cell(page, 'A1', 'blank')).toBeAttached();
    await editor.press('Enter');
    await expect(cell(page, 'A1', 'x,y;z')).toBeAttached();

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
    await picker.click();
    await page.getByRole('menuitem', { name: 'SUM(…)', exact: true }).click();
    const inline = page.getByRole('textbox', { name: 'Edit A1' });
    await expect(inline).toBeFocused();
    await expect(inline).toHaveValue('=SUM()');
    expect(await inline.evaluate(node => (node as HTMLTextAreaElement).selectionStart)).toBe(5);
    await inline.press('Escape');
    await expect(page.getByRole('grid')).toBeFocused();
    await expect(cell(page, 'A1', 'x,y;z')).toBeAttached();
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
    window.Worker = class extends BrowserWorker {
      constructor(url: URL | string, options?: WorkerOptions) {
        super(url, options);
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
        probe.sent.push(performance.now());
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
