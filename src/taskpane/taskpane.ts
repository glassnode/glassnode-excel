/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document,  Office, OfficeRuntime */

// The initialize function must be run each time a new page is loaded
import {ADDIN_VERSION} from '../version';
import {isDevEnv} from '../functions/common/utils';

const API_KEY_STORAGE_KEY = 'glassnodeApiKey';
const COLOR_OK = '#107c10';
const COLOR_ERROR = '#d13438';

Office.onReady(() => {
  document.getElementById('sideload-msg').style.display = 'none';
  document.getElementById('app-body').style.display = 'flex';
  document.getElementById('save-api-key').onclick = saveApiKey;
  document.getElementById('remove-api-key').onclick = removeApiKey;

  // Update version badge if present
  const versionEl = document.getElementById('plugin-version');
  if (versionEl) {
    versionEl.textContent = 'v' + ADDIN_VERSION + (isDevEnv() ? '_DEV' : '');
  }

  // Show whether an API key is saved
  loadApiKey();
});

// Whether a key is saved (null: unknown, storage couldn't be read), and whether a storage
// operation is running. Only one runs at a time, so a late result can't overwrite a newer one.
let keySaved: boolean | null = null;
let busy = false;

function setStatus(message: string, color: string) {
  const statusDiv = document.getElementById('api-status');
  statusDiv.textContent = message;
  statusDiv.style.color = color;
}

function renderButtons() {
  (document.getElementById('save-api-key') as HTMLButtonElement).disabled = busy;
  // Unknown state still allows Remove, so a key that can't be read can still be deleted
  (document.getElementById('remove-api-key') as HTMLButtonElement).disabled = busy || keySaved === false;
}

async function exclusive(operation: () => Promise<void>) {
  if (busy) return;
  busy = true;
  renderButtons();
  try {
    await operation();
  } finally {
    busy = false;
    renderButtons();
  }
}

function errorMessage(error: unknown) {
  return error instanceof Error ? error.message : 'Unknown error';
}

/** Describes a saved key without revealing it: only its last four characters. */
export function describeSavedKey(apiKey: string): string {
  const ending = apiKey.length > 8 ? ` ending in ${apiKey.slice(-4)}` : '';
  return `API key saved${ending}. Enter a new key to replace it.`;
}

export function saveApiKey() {
  return exclusive(async () => {
    const apiKeyInput = document.getElementById('api-key-input') as HTMLInputElement;
    const apiKey = apiKeyInput.value.trim();

    if (!apiKey) {
      setStatus('Please enter an API key', COLOR_ERROR);
      return;
    }

    try {
      // Save using OfficeRuntime.storage (works across all contexts)
      await OfficeRuntime.storage.setItem(API_KEY_STORAGE_KEY, apiKey);
    } catch (error) {
      console.error('Error saving API key:', errorMessage(error));
      setStatus(`Error saving API key: ${errorMessage(error)}`, COLOR_ERROR);
      return;
    }

    // Don't keep the key in the page once it's saved (unless the field was edited meanwhile)
    if (apiKeyInput.value.trim() === apiKey) {
      apiKeyInput.value = '';
    }
    keySaved = true;
    setStatus(describeSavedKey(apiKey), COLOR_OK);
  });
}

export function removeApiKey() {
  return exclusive(async () => {
    try {
      await OfficeRuntime.storage.removeItem(API_KEY_STORAGE_KEY);
    } catch (error) {
      console.error('Error removing API key:', errorMessage(error));
      setStatus(`Error removing API key: ${errorMessage(error)}`, COLOR_ERROR);
      return;
    }

    keySaved = false;
    setStatus('API key removed', COLOR_OK);
  });
}

export function loadApiKey() {
  return exclusive(async () => {
    let apiKey: string | null;
    try {
      // Use OfficeRuntime.storage (works across all contexts)
      apiKey = await OfficeRuntime.storage.getItem(API_KEY_STORAGE_KEY);
    } catch (error) {
      console.error('Error loading API key:', errorMessage(error));
      keySaved = null;
      setStatus('Could not read the saved API key. Close and reopen this pane to try again.', COLOR_ERROR);
      return;
    }

    // The saved key is never put back into the input field
    keySaved = Boolean(apiKey);
    if (apiKey) {
      setStatus(describeSavedKey(apiKey), COLOR_OK);
    }
  });
}
