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

function setStatus(message: string, color: string) {
  const statusDiv = document.getElementById('api-status');
  statusDiv.textContent = message;
  statusDiv.style.color = color;
}

function setRemoveEnabled(enabled: boolean) {
  (document.getElementById('remove-api-key') as HTMLButtonElement).disabled = !enabled;
}

/** Describes a saved key without revealing it: only its last four characters. */
export function describeSavedKey(apiKey: string): string {
  const ending = apiKey.length > 8 ? ` ending in ${apiKey.slice(-4)}` : '';
  return `API key saved${ending}. Enter a new key to replace it.`;
}

export async function saveApiKey() {
  try {
    const apiKeyInput = document.getElementById('api-key-input') as HTMLInputElement;
    const apiKey = apiKeyInput.value.trim();

    if (!apiKey) {
      setStatus('Please enter an API key', COLOR_ERROR);
      return;
    }

    // Save using OfficeRuntime.storage (works across all contexts)
    await OfficeRuntime.storage.setItem(API_KEY_STORAGE_KEY, apiKey);

    // Don't keep the key in the page once it's saved
    apiKeyInput.value = '';
    setRemoveEnabled(true);
    setStatus(describeSavedKey(apiKey), COLOR_OK);
  } catch (error) {
    console.error('Error saving API key:', error);
    setStatus(`Error saving API key: ${error instanceof Error ? error.message : 'Unknown error'}`, COLOR_ERROR);
  }
}

export async function removeApiKey() {
  try {
    await OfficeRuntime.storage.removeItem(API_KEY_STORAGE_KEY);

    (document.getElementById('api-key-input') as HTMLInputElement).value = '';
    setRemoveEnabled(false);
    setStatus('API key removed', COLOR_OK);
  } catch (error) {
    console.error('Error removing API key:', error);
    setStatus(`Error removing API key: ${error instanceof Error ? error.message : 'Unknown error'}`, COLOR_ERROR);
  }
}

export async function loadApiKey() {
  try {
    // Use OfficeRuntime.storage (works across all contexts)
    const apiKey = await OfficeRuntime.storage.getItem(API_KEY_STORAGE_KEY);

    // The saved key is never put back into the input field
    setRemoveEnabled(Boolean(apiKey));
    if (apiKey) {
      setStatus(describeSavedKey(apiKey), COLOR_OK);
    }
  } catch (error) {
    console.error('Error loading API key:', error);
  }
}
