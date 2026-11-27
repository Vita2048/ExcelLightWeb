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

test('F2 and double-click open the cell editor with the existing content selected, and typing replaces it', async ({ page }) => {
  await page.goto(path);
  const grid = await go(page, 'A1');
  await page.getByRole('textbox', { name: 'Formula bar' }).fill('existing');
  await page.getByRole('textbox', { name: 'Formula bar' }).press('Enter');
  await expect(cell(page, 'A1', 'existing')).toBeAttached();

  await go(page, 'A1');
  await grid.press('F2');
  const inline = page.getByRole('textbox', { name: 'Edit A1' });
  await expect(inline).toBeFocused();
  expect(await inline.evaluate(node => {
    const input = node as HTMLTextAreaElement;
    return [input.selectionStart, input.selectionEnd, input.value];
  })).toEqual([0, 8, 'existing']);
  await inline.press('Escape');
  await expect(grid).toBeFocused();

  const box = (await grid.boundingBox())!;
  const header = Number(await grid.getAttribute('data-row-header'));
  await page.mouse.dblclick(box.x + header + 12, box.y + 38);
  await expect(inline).toBeFocused();
  expect(await inline.evaluate(node => {
    const input = node as HTMLTextAreaElement;
    return [input.selectionStart, input.selectionEnd, input.value];
  })).toEqual([0, 8, 'existing']);
  await page.keyboard.type('replacement');
  await expect(inline).toHaveValue('replacement');
  await inline.press('Enter');
  await expect(cell(page, 'A1', 'replacement')).toBeAttached();
});

test('formula hints follow the active function and caret in both editors', async ({ page }) => {
  await page.goto(path);
  const grid = await go(page, 'A1');
  const editor = page.getByRole('textbox', { name: 'Formula bar' });
  const hint = page.getByRole('tooltip');
  await editor.fill('=IF(ROUND(2.5,');
  await expect(hint).toHaveText('ROUND(number, num_digits)');
  await expect(hint.locator('strong')).toHaveText('num_digits');
  await expect(editor).toHaveAccessibleDescription('ROUND(number, num_digits)');
  await editor.press('Home');
  await expect(hint).toHaveCount(0);
  await editor.press('End');
  await expect(hint).toBeVisible();
  await page.keyboard.insertText('0),');
  await expect(hint.locator('strong')).toHaveText('[value_if_true]');
  await page.keyboard.insertText('"x,y","z")');
  await expect(hint).toHaveCount(0);
  await editor.press('Enter');
  await expect(cell(page, 'A1', 'x,y')).toBeAttached();

  await go(page, 'B1');
  await grid.press('F2');
  const inline = page.getByRole('textbox', { name: 'Edit B1' });
  await inline.fill('=SUM({1,2},');
  await expect(hint).toHaveText('SUM(number1, [number2], …)');
  await expect(hint.locator('strong')).toHaveText('[number2]');
  await inline.press('Escape');
  await expect(hint).toHaveCount(0);
  await expect(grid).toBeFocused();
  await expect(cell(page, 'B1', 'blank')).toBeAttached();
});

test('native formulas recalculate dependencies and retain unsupported source in the worksheet', async ({ page }) => {
  await page.goto(path);
  const editor = page.getByRole('textbox', { name: 'Formula bar' });
  async function enter(address: string, source: string, value: string) {
    await go(page, address);
    await editor.fill(source);
    await editor.press('Enter');
    await expect(cell(page, address, value)).toBeAttached();
  }
  await enter('A1', '1', '1');
  await enter('A2', '=IF(A1=1,LEFT("hello",2),1/0)', 'he');
  await enter('A3', '=IFERROR(A2,"caught")', 'he');
  await enter('A1', '0', '0');
  await expect(cell(page, 'A2', '#DIV/0!')).toBeAttached();
  await expect(cell(page, 'A3', 'caught')).toBeAttached();
  await enter('B1', '=DAY(DATE(1900,2,29))', '29');
  await enter('B2', '=UNKNOWNFUNC(1)', '#NAME?');
  await enter('B3', '=SUM({1,2}+1)', '5');
  await enter('B4', '=SUBTOTAL(9,B1,B3)', '34');
  await go(page, 'B2');
  await expect(editor).toHaveValue('=UNKNOWNFUNC(1)');
});

test('invalid draft blocks selection and commands, and a valid blur commits to the original cell', async ({ page }) => {
  await page.goto(path);
  const grid = await go(page, 'B2');
  const formula = page.getByRole('textbox', { name: 'Formula bar' });
  await formula.fill('=');
  await formula.press('Home');
  const { x, y } = (await grid.boundingBox())!;
  const header = Number(await grid.getAttribute('data-row-header'));
  await page.mouse.click(x + header + 12, y + 38);
  await expect(formula).toBeFocused();
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('B2');
  await expect(page.getByText(/formula must/u)).toBeVisible();
  await expect(page.getByRole('button', { name: 'Undo' })).toBeDisabled();
  await formula.fill('TRUE');
  await page.mouse.click(x + header + 12, y + 38);
  await expect(cell(page, 'B2', 'TRUE')).toBeAttached();
  await expect(page.getByRole('textbox', { name: 'Name box' })).toHaveValue('A1');
  await page.mouse.dblclick(x + header + 12, y + 38);
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
  const header = Number(await grid.getAttribute('data-row-header'));
  const { x, y } = (await grid.boundingBox())!;
  await page.keyboard.down('Control');
  await page.mouse.click(x + header + 96 * 2 + 20, y + 37);
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
  await page.getByRole('button', { name: 'Fill', exact: true }).click();
  await page.getByRole('menuitem', { name: 'Down', exact: true }).click();
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
  await page.getByRole('button', { name: 'Paste options' }).click();
  await page.getByRole('menuitem', { name: 'Paste Values' }).click();
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
