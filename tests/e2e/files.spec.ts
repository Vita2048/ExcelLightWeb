import { expect, test, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const appPath = '/ExcelLightWeb/';
const nativeFixture = fileURLToPath(new URL('../fixtures/sales/Showcase.exlt', import.meta.url));

test.beforeEach(async ({ page }) => {
  page.on('dialog', dialog => { void dialog.accept(); });
});

test.afterEach(async ({ page }) => {
  if (!page.isClosed()) await page.close({ runBeforeUnload: false });
});

async function disableFileSystemPickers(page: Page) {
  await page.addInitScript(() => {
    Object.defineProperty(window, 'showOpenFilePicker', { configurable: true, value: undefined });
    Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: undefined });
  });
}

async function openFileBackstage(page: Page) {
  await page.getByRole('tab', { name: 'File' }).click();
  await expect(page.getByRole('dialog', { name: 'File Home' })).toBeVisible();
}

async function openShowcase(page: Page) {
  await openFileBackstage(page);
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Open', exact: true }).click();
  await (await chooser).setFiles(nativeFixture);
  await expect(page.locator('.workbook-title')).toHaveText('Showcase');
}

async function seedVersionOneDatabase(page: Page, settings: Record<string, unknown>) {
  await page.evaluate(async oldSettings => {
    const name = 'excellight-workspace';
    await new Promise<void>((resolve, reject) => {
      const remove = indexedDB.deleteDatabase(name);
      remove.onblocked = () => reject(new Error('Old workspace database stayed open.'));
      remove.onerror = () => reject(new Error('Could not reset the old workspace database: ' +
        (remove.error ? remove.error.name + ': ' + remove.error.message : 'IndexedDB supplied no error detail.')));
      remove.onsuccess = () => {
        const create = indexedDB.open(name, 1);
        create.onupgradeneeded = () => {
          const db = create.result;
          db.createObjectStore('workspaces', { keyPath: 'id' });
          db.createObjectStore('settings', { keyPath: 'id' });
          db.createObjectStore('recovery', { keyPath: 'id' });
          db.createObjectStore('locks', { keyPath: 'id' });
        };
        create.onerror = () => reject(create.error);
        create.onsuccess = () => {
          const db = create.result;
          const transaction = db.transaction(['settings', 'workspaces'], 'readwrite');
          transaction.objectStore('settings').put({ id: 'user', ...oldSettings });
          transaction.objectStore('workspaces').put({
            id: 'retained', name: 'Retained v1 workspace', package: new Uint8Array([1, 2, 3]).buffer,
            source: 'native', createdAt: 1, updatedAt: 1, lastUsedAt: 1,
          });
          transaction.oncomplete = () => { db.close(); resolve(); };
          transaction.onabort = transaction.onerror = () => reject(transaction.error);
        };
      };
    });
  }, settings);
}

async function seedVersionOneAndReload(page: Page, settings: Record<string, unknown>) {
  // Use a same-origin 404 page so the app cannot open and hold its database while seeding.
  await page.goto(appPath + 'test-only-seed');
  await seedVersionOneDatabase(page, settings);
  await page.goto(appPath);
  await expect(page.getByRole('grid')).toBeVisible();
}

async function readLocalCalculationMode(page: Page): Promise<string> {
  return page.evaluate(async () => new Promise<string>((resolve, reject) => {
    const open = indexedDB.open('excellight-workspace');
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const request = db.transaction('preferences', 'readonly').objectStore('preferences').get('user');
      request.onsuccess = () => { resolve(request.result.nativeCalculation); db.close(); };
      request.onerror = () => { reject(request.error); db.close(); };
    };
  }));
}

async function readLocalWorkbook(page: Page, name: string): Promise<boolean> {
  return page.evaluate(async expectedName => new Promise<boolean>((resolve, reject) => {
    const open = indexedDB.open('excellight-workspace');
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const request = db.transaction('workspaces', 'readonly').objectStore('workspaces').getAll();
      request.onsuccess = () => {
        resolve(request.result.some((record: { name: string; package: Blob | ArrayBuffer }) =>
          record.name === expectedName && (record.package instanceof Blob ? record.package.size : record.package.byteLength) > 0));
        db.close();
      };
      request.onerror = () => { reject(request.error); db.close(); };
    };
  }), name);
}

async function readLocalProvenance(page: Page, name: string): Promise<{ source: string; sourceFileName?: string } | undefined> {
  return page.evaluate(async expectedName => new Promise<{ source: string; sourceFileName?: string } | undefined>((resolve, reject) => {
    const open = indexedDB.open('excellight-workspace');
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const request = db.transaction('workspaces', 'readonly').objectStore('workspaces').getAll();
      request.onsuccess = () => {
        const record = request.result.find((item: { name: string }) => item.name === expectedName);
        resolve(record && { source: record.source, ...(record.sourceFileName && { sourceFileName: record.sourceFileName }) });
        db.close();
      };
      request.onerror = () => { reject(request.error); db.close(); };
    };
  }), name);
}

async function seedImportedSourceProvenance(page: Page, name: string, sourceFileName: string) {
  await page.evaluate(async ({ expectedName, importedFileName }) => new Promise<void>((resolve, reject) => {
    const open = indexedDB.open('excellight-workspace');
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const transaction = db.transaction('workspaces', 'readwrite');
      const store = transaction.objectStore('workspaces');
      const request = store.getAll();
      request.onsuccess = () => {
        const record = request.result.find((item: { name: string }) => item.name === expectedName);
        if (!record) { transaction.abort(); return; }
        store.put({ ...record, source: 'xlsx', sourceFileName: importedFileName });
      };
      transaction.oncomplete = () => { db.close(); resolve(); };
      transaction.onabort = transaction.onerror = () => { db.close(); reject(transaction.error ?? new Error('Workspace provenance update failed.')); };
    };
  }), { expectedName: name, importedFileName: sourceFileName });
}

async function readLocalRecoveryCount(page: Page): Promise<number> {
  return page.evaluate(async () => new Promise<number>((resolve, reject) => {
    const open = indexedDB.open('excellight-workspace');
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const request = db.transaction('recovery', 'readonly').objectStore('recovery').count();
      request.onsuccess = () => { resolve(request.result); db.close(); };
      request.onerror = () => { reject(request.error); db.close(); };
    };
  }));
}

test('uses the desktop File backstage and persists the Formulas calculation choice', async ({ page }) => {
  await page.goto(appPath);
  await expect(page.getByRole('grid')).toBeVisible();
  const quickAccess = page.getByRole('group', { name: 'Quick Access' });
  await expect(quickAccess.getByRole('button')).toHaveCount(3);
  await expect(quickAccess.getByRole('button').nth(0)).toHaveAccessibleName('Save');
  await expect(quickAccess.getByRole('button').nth(1)).toHaveAccessibleName('Undo');
  await expect(quickAccess.getByRole('button').nth(2)).toHaveAccessibleName('Redo');
  await openFileBackstage(page);
  const backstage = page.getByRole('dialog', { name: 'File Home' });
  await expect(backstage.locator('.file-backstage-nav')).toHaveCSS('width', '200px');
  await expect(page.getByRole('heading', { name: 'Home' })).toBeVisible();
  await expect(page.getByRole('heading', { name: 'New' })).toBeVisible();
  await expect(page.getByRole('heading', { name: 'Recent' })).toBeVisible();
  await expect(backstage.locator('.file-home-page')).toHaveCSS('padding-left', '32px');
  await expect(backstage.getByRole('button', { name: 'Blank workbook' })).toHaveCSS('min-height', '32px');
  await expect(page.getByText('No recent workbooks')).toBeVisible();
  await expect(backstage.locator('.file-backstage-nav button')).toHaveText([
    '‹ Back', 'Home', 'New', 'Open', 'Save', 'Save As', 'Export', 'Close', 'Options',
  ]);

  await page.getByRole('button', { name: 'Options', exact: true }).click();
  const options = page.getByRole('dialog', { name: 'Excel-Lite Options' });
  await expect(options).toBeVisible();
  expect((await options.boundingBox())?.width).toBe(520);
  await expect(options.locator('.options-categories')).toHaveCSS('width', '140px');
  await expect(options.getByRole('tab', { name: 'General' })).toHaveAttribute('aria-selected', 'true');
  await expect(options.getByLabel('Office Theme')).toHaveValue('dark');
  await expect(options.getByLabel('Office Theme')).toBeDisabled();
  await options.getByRole('tab', { name: 'Formulas' }).click();
  await expect(options.locator('.python-option-group input[name="python-calculation"]').first()).toBeDisabled();
  await expect(options.getByText(/Python calculation is not available yet/)).toBeVisible();
  await expect(options.locator('input[name="native-calculation"][value="automatic"]')).toBeChecked();
  await options.locator('input[name="native-calculation"][value="manual"]').check();
  await options.getByRole('button', { name: 'OK' }).click();
  await expect(options).not.toBeVisible();
  await expect.poll(() => readLocalCalculationMode(page)).toBe('manual');
  await page.reload();
  await expect(page.getByRole('grid')).toBeVisible();
  await openFileBackstage(page);
  await page.getByRole('button', { name: 'Options', exact: true }).click();
  await page.getByRole('dialog', { name: 'Excel-Lite Options' }).getByRole('tab', { name: 'Formulas' }).click();
  await expect(page.locator('input[name="native-calculation"][value="manual"]')).toBeChecked();
});

test('Quick Access Save names the browser copy and Recent can remove it without touching disk', async ({ page }) => {
  await disableFileSystemPickers(page);
  await page.goto(appPath);
  const quickAccess = page.getByRole('group', { name: 'Quick Access' });
  const save = quickAccess.getByRole('button', { name: 'Save', exact: true });
  await expect(save).toBeEnabled();
  const grid = page.getByRole('grid');
  await grid.focus();
  await page.keyboard.type('Quick Access save');
  await page.keyboard.press('Enter');
  await save.click();
  const nameDialog = page.getByRole('dialog', { name: 'Save to this browser' });
  await expect(nameDialog).toBeVisible();
  await expect(nameDialog.getByRole('textbox', { name: 'Workbook name' })).toHaveValue('Book1');
  await nameDialog.getByRole('textbox', { name: 'Workbook name' }).fill('Forecast');
  await nameDialog.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Forecast');
  await expect(page.getByRole('button', { name: 'Sheet1', exact: true })).toBeVisible();
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Saved to this browser');
  await expect.poll(() => readLocalWorkbook(page, 'Forecast')).toBe(true);
  await openFileBackstage(page);
  const remove = page.getByRole('button', { name: 'Remove Forecast from this browser' });
  await expect(remove).toBeDisabled();
  await page.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Book1');
  await expect(remove).toBeEnabled();
  await remove.click();
  const removal = page.getByRole('dialog', { name: 'Remove Forecast from this browser?' });
  await expect(removal.getByText(/does not delete an \.exlt file in Downloads/)).toBeVisible();
  await removal.getByRole('button', { name: 'Remove', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Open recent Forecast' })).toHaveCount(0);
  await expect.poll(() => readLocalWorkbook(page, 'Forecast')).toBe(false);
});

test('opens a native workbook, saves locally, lists Recent and exports without marking later edits saved', async ({ page }) => {
  await disableFileSystemPickers(page);
  await page.goto(appPath);
  await openShowcase(page);
  const grid = page.getByRole('grid');
  await grid.focus();
  await page.keyboard.type('Saved in browser');
  await page.keyboard.press('Enter');
  await openFileBackstage(page);
  await page.getByRole('dialog', { name: 'File Home' }).getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Showcase');
  await expect(page.locator('.status-bar [role="status"]')).toContainText('Saved to this browser');

  await openFileBackstage(page);
  await expect(page.getByRole('button', { name: 'Open recent Showcase' })).toBeVisible();
  await page.getByRole('button', { name: 'New', exact: true }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Book1');
  await openFileBackstage(page);
  await page.getByRole('button', { name: 'Open recent Showcase' }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Showcase');
  await expect(page.getByRole('gridcell', { name: 'A1, Saved in browser' })).toHaveAttribute('aria-label', 'A1, Saved in browser');

  await page.getByRole('grid').focus();
  await page.keyboard.type('Newer edit');
  await page.keyboard.press('Enter');
  await openFileBackstage(page);
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  expect((await download).suggestedFilename()).toMatch(/\.exlt$/i);
  await expect(page.locator('.workbook-title')).toHaveText('Showcase *');
  await expect(page.locator('.status-bar .edit-status')).toContainText('Export download started');
});

test('flushes an editor before New and honors Save, Don’t Save and Cancel', async ({ page }) => {
  await disableFileSystemPickers(page);
  await page.goto(appPath);
  const grid = page.getByRole('grid');
  await grid.focus();
  await page.keyboard.type('Draft text');
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'New', exact: true }).click();
  const prompt = page.getByRole('dialog', { name: 'Save changes to Book1?' });
  await expect(prompt).toBeVisible();
  await prompt.getByRole('button', { name: 'Cancel' }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Book1 *');
  await expect(page.getByRole('gridcell', { name: 'A1, Draft text' })).toHaveAttribute('aria-label', 'A1, Draft text');
  await page.getByRole('button', { name: 'New', exact: true }).click();
  await page.getByRole('dialog', { name: 'Save changes to Book1?' }).getByRole('button', { name: 'Save', exact: true }).click();
  const localSave = page.getByRole('dialog', { name: 'Save to this browser' });
  await expect(localSave).toBeVisible();
  await localSave.getByRole('textbox', { name: 'Workbook name' }).fill('Draft workbook');
  await localSave.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Book1');
  await expect.poll(() => readLocalWorkbook(page, 'Draft workbook')).toBe(true);

  await grid.focus();
  await page.keyboard.type('Discard me');
  await page.keyboard.press('Enter');
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'New', exact: true }).click();
  await page.getByRole('dialog', { name: 'Save changes to Book1?' }).getByRole('button', { name: "Don't Save" }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Book1');
  await expect(page.getByRole('gridcell', { name: 'A1, blank' })).toHaveAttribute('aria-label', 'A1, blank');
});

test('cancels local Save As without changing the draft and saves a named Recent copy', async ({ page }) => {
  await disableFileSystemPickers(page);
  await page.goto(appPath);
  const grid = page.getByRole('grid');
  await grid.focus();
  await page.keyboard.type('Save As draft');
  await page.keyboard.press('Enter');

  await openFileBackstage(page);
  await page.getByRole('button', { name: 'Save As', exact: true }).click();
  const saveAs = page.getByRole('dialog', { name: 'Save As' });
  await expect(saveAs).toBeVisible();
  await saveAs.getByRole('button', { name: 'Cancel', exact: true }).click();
  await expect(saveAs).not.toBeVisible();
  await expect(page.locator('.workbook-title')).toHaveText('Book1 *');
  await expect(page.getByRole('gridcell', { name: 'A1, Save As draft' })).toHaveAttribute('aria-label', 'A1, Save As draft');

  await page.getByRole('button', { name: 'Save As', exact: true }).click();
  const saveAsAgain = page.getByRole('dialog', { name: 'Save As' });
  await saveAsAgain.getByRole('textbox', { name: 'Workbook name' }).fill('Quarterly');
  await saveAsAgain.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Quarterly');
  await expect.poll(() => readLocalWorkbook(page, 'Quarterly')).toBe(true);
  await openFileBackstage(page);
  await expect(page.getByRole('button', { name: 'Open recent Quarterly' })).toBeVisible();
});

test('keeps imported-source provenance when a Recent workbook is locally saved', async ({ page }) => {
  await disableFileSystemPickers(page);
  await page.goto(appPath);
  await openShowcase(page);
  await seedImportedSourceProvenance(page, 'Showcase', 'annual-report.xlsx');
  await openFileBackstage(page);
  await page.getByRole('button', { name: 'Open recent Showcase' }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Showcase');
  await page.getByRole('grid').focus();
  await page.keyboard.type('Edited imported source');
  await page.keyboard.press('Enter');
  await openFileBackstage(page);
  await page.getByRole('dialog', { name: 'File Home' }).getByRole('button', { name: 'Save', exact: true }).click();
  await expect.poll(() => readLocalProvenance(page, 'Showcase')).toEqual({ source: 'xlsx', sourceFileName: 'annual-report.xlsx' });
});

test('offers explicit reload recovery for an unsaved native snapshot', async ({ page }) => {
  await disableFileSystemPickers(page);
  await page.goto(appPath);
  const grid = page.getByRole('grid');
  await grid.focus();
  await page.keyboard.type('Recovered value');
  await page.keyboard.press('Enter');
  await expect.poll(() => readLocalRecoveryCount(page)).toBe(1);
  await openFileBackstage(page);
  await expect(page.locator('.file-recovery-status')).toHaveText('Local recovery copy updated.');
  await page.getByRole('button', { name: /Back/ }).click();
  await page.reload();
  const recovery = page.getByRole('dialog', { name: 'Recover your workbook?' });
  await expect.poll(() => readLocalRecoveryCount(page)).toBe(1);
  await expect(recovery).toBeVisible();
  await recovery.getByRole('button', { name: 'Recover' }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Book1 *');
  await expect(page.getByRole('gridcell', { name: 'A1, Recovered value' })).toHaveAttribute('aria-label', 'A1, Recovered value');
});

test('does not open the same Recent workbook in two tabs and releases it through Close', async ({ page }) => {
  await page.context().addInitScript(() => {
    Object.defineProperty(window, 'showOpenFilePicker', { configurable: true, value: undefined });
    Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: undefined });
  });
  await page.goto(appPath);
  await openShowcase(page);
  const second = await page.context().newPage();
  await second.goto(appPath);
  await openFileBackstage(second);
  await second.getByRole('button', { name: 'Remove Showcase from this browser' }).click();
  const removal = second.getByRole('dialog', { name: 'Remove Showcase from this browser?' });
  await removal.getByRole('button', { name: 'Remove', exact: true }).click();
  await expect(second.getByRole('dialog', { name: 'Workbook already open' })).toBeVisible();
  await second.getByRole('dialog', { name: 'Workbook already open' }).getByRole('button', { name: 'Close', exact: true }).click();
  await expect(second.getByRole('button', { name: 'Open recent Showcase' })).toBeVisible();
  await second.getByRole('button', { name: 'Open recent Showcase' }).click();
  await expect(second.getByRole('dialog', { name: 'Workbook already open' })).toBeVisible();
  await expect(second.locator('.workbook-title')).toHaveText('Book1');
  await second.getByRole('dialog', { name: 'Workbook already open' }).getByRole('button', { name: 'Close', exact: true }).click();
  await page.bringToFront();
  await page.getByRole('tab', { name: 'File' }).click();
  await page.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Book1');
  await second.bringToFront();
  await second.getByRole('button', { name: 'Open recent Showcase' }).click();
  await expect(second.locator('.workbook-title')).toHaveText('Showcase');
});

test('an IndexedDB transaction failure preserves the dirty workbook and its values', async ({ page }) => {
  await disableFileSystemPickers(page);
  await page.goto(appPath);
  const grid = page.getByRole('grid');
  await grid.focus();
  await page.keyboard.type('Keep after quota failure');
  await page.keyboard.press('Enter');
  await page.evaluate(() => {
    const original = IDBObjectStore.prototype.put;
    IDBObjectStore.prototype.put = function (value: unknown, key?: IDBValidKey) {
      if (this.name === 'workspaces') {
        this.transaction.abort();
        throw new DOMException('Quota exceeded.', 'QuotaExceededError');
      }
      return original.call(this, value, key);
    };
  });
  await openFileBackstage(page);
  await page.getByRole('dialog', { name: 'File Home' }).getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.getByRole('dialog', { name: 'Save failed' })).toBeVisible();
  await page.getByRole('dialog', { name: 'Save failed' }).getByRole('button', { name: 'Close', exact: true }).click();
  await expect(page.locator('.workbook-title')).toHaveText('Book1 *');
  await expect(page.getByRole('gridcell', { name: 'A1, Keep after quota failure' })).toHaveAttribute('aria-label', 'A1, Keep after quota failure');
});

test('migrates v1 local settings and retains workspaces', async ({ page }) => {
  await seedVersionOneAndReload(page, { manual: true, pythonAutomatic: false, findHistory: ['reopened'] });
  await openFileBackstage(page);
  await expect(page.getByRole('button', { name: 'Open recent Retained v1 workspace' })).toBeVisible();
  await page.getByRole('button', { name: 'Options', exact: true }).click();
  await page.getByRole('dialog', { name: 'Excel-Lite Options' }).getByRole('tab', { name: 'Formulas' }).click();
  await expect(page.locator('input[name="native-calculation"][value="manual"]')).toBeChecked();
  await expect(page.locator('input[name="python-calculation"]:checked')).toBeChecked();
});

test('aborts an invalid settings migration without deleting existing local data', async ({ page }) => {
  await seedVersionOneAndReload(page, { manual: 'invalid' });
  await openFileBackstage(page);
  await expect(page.getByText(/Local calculation preference is invalid/)).toBeVisible();
  const retained = await page.evaluate(async () => new Promise<{ version: number; workspace: string; manual: unknown }>((resolve, reject) => {
    const open = indexedDB.open('excellight-workspace', 1);
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const transaction = db.transaction(['settings', 'workspaces'], 'readonly');
      const settings = transaction.objectStore('settings').get('user');
      const workspace = transaction.objectStore('workspaces').get('retained');
      let manual: unknown;
      let name = '';
      settings.onsuccess = () => { manual = settings.result.manual; };
      workspace.onsuccess = () => { name = workspace.result.name; };
      transaction.oncomplete = () => { resolve({ version: db.version, workspace: name, manual }); db.close(); };
      transaction.onerror = () => { reject(transaction.error); db.close(); };
    };
  }));
  expect(retained).toEqual({ version: 1, workspace: 'Retained v1 workspace', manual: 'invalid' });
});
