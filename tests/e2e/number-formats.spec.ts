import { expect, test, type Page } from '@playwright/test';

const path = '/ExcelLightWeb/';
const formulaBar = (page: Page) => page.getByRole('textbox', { name: 'Formula bar' });
const cell = (page: Page, address: string, text: string) => page.getByRole('gridcell', { name: address + ', ' + text, exact: true });

async function go(page: Page, address: string) {
  const input = page.getByRole('textbox', { name: 'Name box' });
  await input.fill(address);
  await input.press('Enter');
}

async function enter(page: Page, address: string, value: string) {
  await go(page, address);
  await formulaBar(page).fill(value);
  await formulaBar(page).press('Enter');
}

async function openFormat(page: Page) {
  await page.getByRole('button', { name: 'Format Cells', exact: true }).click();
  const dialog = page.getByRole('dialog', { name: 'Format Cells' });
  await expect(dialog).toBeVisible();
  return dialog;
}

test('creates a custom red display and cancels or accepts mixed formats without changing their codes', async ({ page }) => {
  await page.goto(path);
  await enter(page, 'A1', '-1234.5');
  await enter(page, 'A2', '0.125');
  await go(page, 'A2');
  await page.getByRole('button', { name: 'Percentage', exact: true }).click();
  await expect(cell(page, 'A2', '12.50%')).toBeAttached();
  await go(page, 'A1');

  const dialog = await openFormat(page);
  await dialog.getByLabel('Category').selectOption('Custom');
  const code = dialog.getByRole('textbox', { name: 'Type', exact: true });
  await code.fill('0.00"');
  await expect(dialog.getByRole('alert')).toContainText('valid Excel number format');
  await expect(dialog.getByRole('button', { name: 'OK' })).toBeDisabled();
  const authored = '#,##0.00" kg";[Red](#,##0.00" kg")';
  await code.fill(authored);
  const sample = dialog.getByLabel('Number format sample');
  await expect(sample).toHaveText('(1,234.50 kg)');
  await expect(sample).toHaveCSS('color', 'rgb(255, 0, 0)');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(cell(page, 'A1', '(1,234.50 kg)')).toBeAttached();
  await expect(formulaBar(page)).toHaveValue('-1234.5');
  await expect.poll(() => page.locator('.worksheet-canvas').evaluate((node: HTMLCanvasElement) => {
    const ratio = devicePixelRatio;
    const { data } = node.getContext('2d')!.getImageData(46 * ratio, 27 * ratio, 92 * ratio, 22 * ratio);
    for (let offset = 0; offset < data.length; offset += 4) {
      if (data[offset]! > 150 && data[offset + 1]! < 30 && data[offset + 2]! < 30) return true;
    }
    return false;
  })).toBe(true);

  const grid = page.getByRole('grid');
  await grid.press('Shift+ArrowDown');
  await grid.press('Control+1');
  await expect(code).toHaveValue(authored);
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(cell(page, 'A2', '12.50%')).toBeAttached();
  await expect(grid).toBeFocused();
  await grid.press('Control+1');
  await code.fill('0.0000');
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
  await expect(cell(page, 'A1', '(1,234.50 kg)')).toBeAttached();

  await page.getByRole('button', { name: 'Increase Decimal' }).click();
  await expect(cell(page, 'A1', '(1,234.500 kg)')).toBeAttached();
  await expect(cell(page, 'A2', '12.500%')).toBeAttached();
  await page.getByRole('button', { name: 'Decrease Decimal' }).click();
  await expect(cell(page, 'A1', '(1,234.50 kg)')).toBeAttached();
  await grid.press('Control+z');
  await expect(cell(page, 'A1', '(1,234.500 kg)')).toBeAttached();
  await grid.press('Control+y');
  await expect(cell(page, 'A2', '12.50%')).toBeAttached();
});

test('supports Czech and US dates, the fictitious leap day, elapsed time and canonical formula editing', async ({ page }) => {
  await page.goto(path);
  await enter(page, 'A1', '60.5');
  await go(page, 'A1');
  const dialog = await openFormat(page);
  await dialog.getByLabel('Category').selectOption('Date');
  await dialog.getByLabel('Locale (location)').selectOption('cs-CZ');
  await expect(dialog.getByLabel('Number format sample')).toHaveText('29.02.1900');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(cell(page, 'A1', '29.02.1900')).toBeAttached();
  await expect(formulaBar(page)).toHaveValue('29.02.1900 12:00:00');

  const grid = page.getByRole('grid');
  await grid.press('F2');
  const inline = page.getByRole('textbox', { name: 'Edit A1' });
  await expect(inline).toHaveValue('29.02.1900 12:00:00');
  await inline.press('Enter');
  await grid.press('Control+z');
  await expect(cell(page, 'A1', '60.5')).toBeAttached();
  await grid.press('Control+y');
  await expect(cell(page, 'A1', '29.02.1900')).toBeAttached();
  await enter(page, 'A1', '01.03.1900 6:00:00');
  await expect(cell(page, 'A1', '01.03.1900')).toBeAttached();
  await enter(page, 'B1', '=A1');
  await expect(cell(page, 'B1', '61.25')).toBeAttached();

  await go(page, 'B1');
  await page.getByRole('combobox', { name: 'Number Format', exact: true }).selectOption('more');
  await dialog.getByLabel('Category').selectOption('Time');
  await dialog.getByLabel('Date/time type').selectOption('5');
  await expect(dialog.getByLabel('Number format sample')).toHaveText('1470:00:00');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(cell(page, 'B1', '1470:00:00')).toBeAttached();
  await expect(formulaBar(page)).toHaveValue('=A1');

  await go(page, 'A1');
  await grid.press('Shift+F10');
  await page.getByRole('button', { name: 'Format Cells…', exact: true }).click();
  await dialog.getByLabel('Locale (location)').selectOption('en-US');
  await expect(dialog.getByLabel('Number format sample')).toHaveText('3/1/1900');
  await dialog.getByRole('button', { name: 'OK' }).click();
  await expect(formulaBar(page)).toHaveValue('3/1/1900 6:00:00');
});

test.describe('Czech browser locale', () => {
  test.use({ locale: 'cs-CZ' });
  test('applies localized number options, currencies, fractions and a special mask', async ({ page }) => {
    await page.goto(path);
    await enter(page, 'A1', '123456789');
    await go(page, 'A1');
    const dialog = await openFormat(page);
    await dialog.getByLabel('Category').selectOption('Special');
    await expect(dialog.getByLabel('Locale (location)')).toHaveValue('cs-CZ');
    await dialog.getByLabel('Special type').selectOption({ label: 'telefonní číslo (dlouhé)' });
    await expect(dialog.getByLabel('Number format sample')).toHaveText('+420 123 456 789');
    await dialog.getByRole('button', { name: 'OK' }).click();
    await expect(cell(page, 'A1', '+420 123 456 789')).toBeAttached();

    await enter(page, 'A2', '-1234.125');
    await go(page, 'A2');
    await openFormat(page);
    await dialog.getByLabel('Category').selectOption('Number');
    await dialog.getByLabel('Decimal places').fill('3');
    await dialog.getByLabel('Use 1000 Separator').check();
    await dialog.getByLabel('Negative numbers').selectOption('3');
    await expect(dialog.getByLabel('Number format sample')).toHaveText('-1\u00a0234,125');
    await dialog.getByRole('button', { name: 'OK' }).click();
    await expect(cell(page, 'A2', '-1\u00a0234,125')).toBeAttached();
    await page.getByRole('combobox', { name: 'Currency', exact: true }).selectOption('GBP');
    await expect(cell(page, 'A2', '-£1,234.12')).toBeAttached();

    await enter(page, 'A3', '2.5');
    await go(page, 'A3');
    await openFormat(page);
    await dialog.getByLabel('Category').selectOption('Fraction');
    await dialog.getByLabel('Fraction type').selectOption({ label: 'As eighths (4/8)' });
    await expect(dialog.getByLabel('Number format sample')).toHaveText('2 4/8');
    await dialog.getByRole('button', { name: 'OK' }).click();
    await expect(cell(page, 'A3', '2 4/8')).toBeAttached();
  });
});
