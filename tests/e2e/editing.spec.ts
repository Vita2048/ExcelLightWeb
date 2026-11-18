import { expect, test, type Page } from '@playwright/test';

const path = '/ExcelLightWeb/';

async function go(page: Page, address: string) {
  const nameBox = page.getByRole('textbox', { name: 'Name box' });
  await nameBox.fill(address);
  await nameBox.press('Enter');
  return page.getByRole('grid');
}

function cell(page: Page, address: string, value: string) {
  return page.getByRole('gridcell', { name: address + ', ' + value, exact: true });
}

async function pasteEvent(page: Page, text: string) {
  await page.getByRole('grid').evaluate((node, payload) => {
    const transfer = new DataTransfer();
    transfer.setData('text/plain', payload);
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: transfer });
    node.dispatchEvent(event);
  }, text);
}

async function copyEvent(page: Page, kind: 'copy' | 'cut') {
  return page.getByRole('grid').evaluate((node, kind) => {
    const clipboardData = new DataTransfer();
    const event = new ClipboardEvent(kind, { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    node.dispatchEvent(event);
    return Array.from(clipboardData.types, type => [type, clipboardData.getData(type)] as const);
  }, kind);
}

async function pasteTransfer(page: Page, fields: readonly (readonly [string, string])[]) {
  await page.getByRole('grid').evaluate((node, fields) => {
    const clipboardData = new DataTransfer();
    fields.forEach(([type, value]) => clipboardData.setData(type, value));
    const event = new ClipboardEvent('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    node.dispatchEvent(event);
  }, fields);
}

test('direct entry, formula bar, F2 and Escape share a draft and restore grid focus', async ({ page }) => {
  await page.goto(path);
  const grid = await go(page, 'A1');
  await grid.press('1');
  const inline = page.getByRole('textbox', { name: 'Edit A1' });
  await expect(inline).toBeFocused();
  await expect(inline).toHaveValue('1');
  await expect(inline).toHaveCSS('scrollbar-width', 'none');
  await page.keyboard.type('2');
  await inline.press('ArrowDown');
  await expect(cell(page, 'A1', '12')).toBeAttached();
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('A2');

  const formula = page.getByRole('textbox', { name: 'Formula bar' });
  await expect(formula).toHaveCSS('scrollbar-width', 'none');
  await formula.fill('=A1*2');
  await expect(cell(page, 'A2', 'blank')).toBeAttached();
  await formula.press('Tab');
  await expect(cell(page, 'A2', '24')).toBeAttached();
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('B2');

  await go(page, 'A2');
  await grid.press('F2');
  const second = page.getByRole('textbox', { name: 'Edit A2' });
  await expect(second).toHaveValue('=A1*2');
  await second.fill('=1');
  await expect(formula).toHaveValue('=1');
  await second.press('Escape');
  await expect(grid).toBeFocused();
  await expect(formula).toHaveValue('=A1*2');
  await expect(cell(page, 'A2', '24')).toBeAttached();
  await grid.press('Control+z');
  await expect(cell(page, 'A2', 'blank')).toBeAttached();
  await grid.press('Control+y');
  await expect(cell(page, 'A2', '24')).toBeAttached();
  const zoom = page.getByRole('slider', { name: 'Zoom' });
  await zoom.press('Home');
  await expect(zoom).toHaveAttribute('aria-valuetext', '10%');
  await go(page, 'A1');
  await grid.press('F2');
  await expect(inline).toBeFocused();
  expect((await inline.boundingBox())!.height).toBeGreaterThanOrEqual(24);
  await inline.press('Escape');
});

test('invalid draft blocks selection and commands, and a valid blur commits to the original cell', async ({ page }) => {
  await page.goto(path);
  const grid = await go(page, 'B2');
  const formula = page.getByRole('textbox', { name: 'Formula bar' });
  await formula.fill('=');
  const { x, y } = (await grid.boundingBox())!;
  await page.mouse.click(x + 50, y + 38);
  await expect(formula).toBeFocused();
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('B2');
  await expect(page.getByText(/formula must/u)).toBeVisible();
  await expect(page.getByRole('button', { name: 'Undo' })).toBeDisabled();
  await formula.fill('TRUE');
  await page.mouse.click(x + 50, y + 38);
  await expect(cell(page, 'B2', 'TRUE')).toBeAttached();
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('A1');
  await page.mouse.dblclick(x + 50, y + 38);
  await expect(page.getByRole('textbox', { name: 'Edit A1' })).toBeFocused();
});

test('composition gates Escape and Enter until preedit ends', async ({ page }) => {
  await page.goto(path);
  const grid = await go(page, 'A1');
  await grid.press('F2');
  const input = page.getByRole('textbox', { name: 'Edit A1' });
  await input.fill('漢字');
  await input.evaluate(node => {
    node.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '漢字' }));
    for (const key of ['Enter', 'Escape']) {
      node.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key, isComposing: true }));
    }
  });
  await expect(input).toHaveValue('漢字');
  await expect(cell(page, 'A1', 'blank')).toBeAttached();
  await input.evaluate(node => node.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '漢字' })));
  await input.press('Enter');
  await expect(cell(page, 'A1', '漢字')).toBeAttached();
  await expect(grid).toBeFocused();
});

test('clipboard event keeps overlapping moves atomic and multi-area clear leaves intervening cells', async ({ page }) => {
  await page.goto(path);
  const grid = await go(page, 'A1');
  await pasteEvent(page, '1\t2\t3\n=A1+$B$1\t=SUM(A1:C1)\t"001"');
  await expect(cell(page, 'B2', '6')).toBeAttached();
  await grid.press('Shift+ArrowRight');
  const fields = await copyEvent(page, 'cut');
  expect(fields).toContainEqual(['text/plain', '1\t2']);
  await expect(cell(page, 'A1', '1')).toBeAttached();
  await go(page, 'B1');
  await pasteTransfer(page, fields);
  await expect(cell(page, 'A1', 'blank')).toBeAttached();
  await expect(cell(page, 'C1', '2')).toBeAttached();
  await expect(cell(page, 'B2', '5')).toBeAttached();
  await page.getByRole('button', { name: 'Undo' }).click();
  await expect(cell(page, 'C1', '3')).toBeAttached();
  await expect(cell(page, 'B2', '6')).toBeAttached();
  await go(page, 'A1');
  await copyEvent(page, 'cut');
  await go(page, 'D1');
  // A text-only paste with identical text has no move marker.
  await pasteEvent(page, '1');
  await expect(cell(page, 'A1', '1')).toBeAttached();
  await expect(cell(page, 'D1', '1')).toBeAttached();
  await grid.press('Control+z');
  await go(page, 'A1');
  const { x, y } = (await grid.boundingBox())!;
  await page.keyboard.down('Control');
  await page.mouse.click(x + 257, y + 37);
  await page.keyboard.up('Control');
  await copyEvent(page, 'copy');
  await expect(page.getByText(/one contiguous selection/)).toBeVisible();
  await grid.press('Delete');
  await expect(cell(page, 'A1', 'blank')).toBeAttached();
  await expect(cell(page, 'B1', '2')).toBeAttached();
  await expect(cell(page, 'C1', 'blank')).toBeAttached();
  await grid.press('Control+z');
  await expect(cell(page, 'C1', '3')).toBeAttached();
});

test('ribbon Fill and AutoSum use calculated values and undo each compound edit once', async ({ page }) => {
  await page.goto(path);
  await go(page, 'A1');
  await pasteEvent(page, '4\t=A1*2\n6\t\n10\t');
  const grid = await go(page, 'B1');
  await grid.press('Shift+ArrowDown');
  await grid.press('Shift+ArrowDown');
  await page.getByRole('combobox', { name: 'Fill', exact: true }).selectOption({ label: 'Down' });
  await expect(grid).toBeFocused();
  await expect(cell(page, 'B3', '20')).toBeAttached();
  await page.getByRole('button', { name: 'AutoSum', exact: true }).click();
  await expect(cell(page, 'B4', '40')).toBeAttached();
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue('=SUM(B1:B3)');
  await grid.press('Control+z');
  await expect(cell(page, 'B4', 'blank')).toBeAttached();
  await grid.press('Control+z');
  await expect(cell(page, 'B2', 'blank')).toBeAttached();
  await expect(cell(page, 'B3', 'blank')).toBeAttached();
});

test('denied clipboard permission supports manual paste modes and in-page rich clipboard', async ({ page }) => {
  await page.addInitScript(() => {
    Object.defineProperty(navigator, 'clipboard', { value: {
      readText: async () => { throw new DOMException('denied', 'NotAllowedError'); },
      writeText: async () => { throw new DOMException('denied', 'NotAllowedError'); },
    } });
  });
  await page.goto(path);
  await go(page, 'A1');
  await page.getByRole('combobox', { name: 'Paste options' }).selectOption({ label: 'Paste Values' });
  const dialog = page.getByRole('dialog', { name: 'Paste Values' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('button')).toHaveText(['Paste', 'Cancel']);
  await expect(dialog.getByRole('textbox')).toBeFocused();
  await dialog.getByRole('textbox').fill('=1+2\t"007"');
  await dialog.getByRole('button', { name: 'Paste' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(page.getByRole('grid')).toBeFocused();
  await expect(cell(page, 'A1', '=1+2')).toBeAttached();
  await expect(cell(page, 'B1', '007')).toBeAttached();
  await go(page, 'A1');
  await page.getByRole('button', { name: 'Copy', exact: true }).click();
  await expect(page.getByText(/Copied to this workbook/)).toBeVisible();
  await go(page, 'C1');
  await page.getByRole('button', { name: 'Paste', exact: true }).click();
  await expect(cell(page, 'C1', '=1+2')).toBeAttached();
});

test('late ribbon clipboard reads cannot paste into the old or new selection', async ({ page }) => {
  await page.addInitScript(() => {
    Object.defineProperty(navigator, 'clipboard', { value: { readText: () => new Promise<string>(resolve => {
      document.addEventListener('resolve-test-clipboard', () => resolve('99'), { once: true });
      document.documentElement.dataset.clipboardStarted = 'true';
    }) } });
  });
  await page.goto(path);
  await go(page, 'A1');
  await page.getByRole('button', { name: 'Paste', exact: true }).click();
  await expect(page.locator('html')).toHaveAttribute('data-clipboard-started', 'true');
  await go(page, 'B1');
  await page.evaluate(() => document.dispatchEvent(new Event('resolve-test-clipboard')));
  await expect(cell(page, 'A1', 'blank')).toBeAttached();
  await expect(cell(page, 'B1', 'blank')).toBeAttached();
  await expect(page.getByRole('button', { name: 'Undo' })).toBeDisabled();
});
