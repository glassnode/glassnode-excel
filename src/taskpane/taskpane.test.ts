/* global document */
import * as fs from 'fs';
import * as path from 'path';

// taskpane.ts registers Office.onReady on import; capture the callback so tests can run it
let onReady: () => void;
(global as any).Office = {onReady: (cb: () => void) => (onReady = cb)};

// eslint-disable-next-line @typescript-eslint/no-require-imports
const taskpane = require('./taskpane');

const storage = (global as any).OfficeRuntime.storage as Record<'getItem' | 'setItem' | 'removeItem', jest.Mock>;
const html = fs.readFileSync(path.join(__dirname, 'taskpane.html'), 'utf8');
const SAVED_KEY = 'abcdefgh-1234-saved-key-wxyz';

const input = () => document.getElementById('api-key-input') as HTMLInputElement;
const removeButton = () => document.getElementById('remove-api-key') as HTMLButtonElement;
const status = () => document.getElementById('api-status').textContent;

const flush = () => new Promise(resolve => setTimeout(resolve, 0));

describe('task pane API key settings', () => {
  beforeEach(() => {
    document.documentElement.innerHTML = html.replace(/<script[\s\S]*?<\/script>/g, '');
  });

  it('masks the API key field', () => {
    expect(input().type).toBe('password');
    expect(input().getAttribute('autocomplete')).toBe('off');
  });

  it('does not put a saved key back into the field, and enables Remove', async () => {
    storage.getItem.mockResolvedValue(SAVED_KEY);

    onReady();
    await flush();

    expect(input().value).toBe('');
    expect(document.body.innerHTML).not.toContain(SAVED_KEY);
    expect(status()).toContain('ending in wxyz');
    expect(removeButton().disabled).toBe(false);
  });

  it('disables Remove when no key is saved', async () => {
    storage.getItem.mockResolvedValue(null);

    onReady();
    await flush();

    expect(removeButton().disabled).toBe(true);
  });

  it('clears the field after saving a key', async () => {
    storage.setItem.mockResolvedValue(undefined);
    input().value = `  ${SAVED_KEY}  `;

    await taskpane.saveApiKey();

    expect(storage.setItem).toHaveBeenCalledWith('glassnodeApiKey', SAVED_KEY);
    expect(input().value).toBe('');
    expect(removeButton().disabled).toBe(false);
  });

  it('removes the saved key from storage', async () => {
    storage.removeItem.mockResolvedValue(undefined);
    removeButton().disabled = false;

    await taskpane.removeApiKey();

    expect(storage.removeItem).toHaveBeenCalledWith('glassnodeApiKey');
    expect(removeButton().disabled).toBe(true);
    expect(status()).toBe('API key removed');
  });

  it('only shows the last four characters of a long key', () => {
    expect(taskpane.describeSavedKey(SAVED_KEY)).toBe('API key saved ending in wxyz. Enter a new key to replace it.');
    expect(taskpane.describeSavedKey('short')).toBe('API key saved. Enter a new key to replace it.');
  });
});
