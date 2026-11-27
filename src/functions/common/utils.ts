/* global OfficeRuntime */

import {GlassnodeApiError, GlassnodeError, GlassnodeNetworkError, GlassnodeValidationError} from 'glassnode-api';

export const API_KEY_MISSING_ERROR = 'Error: API key not configured. Please set your API key in the task pane.';
export const API_KEY_STORAGE_ERROR = 'Error: could not read the API key from storage. Please try again.';

/**
 * Get the API key from OfficeRuntime.storage.
 * @returns The stored API key, or null if none is saved.
 * @throws If storage can't be read, so callers can tell a failure from "no key saved".
 */
export async function getApiKey(): Promise<string | null> {
  return await OfficeRuntime.storage.getItem('glassnodeApiKey');
}

/**
 * A loggable summary of an error. Never log a whole error object: a client error can carry the
 * failed request (`cause`), and with it the X-Api-Key header.
 */
export function describeError(error: unknown): Record<string, unknown> {
  if (error instanceof GlassnodeError) {
    return {name: error.name, status: error instanceof GlassnodeApiError ? error.status : undefined, message: error.message};
  }
  return {message: error instanceof Error ? error.message : String(error)};
}

/** The cell text for a failed API call. */
export function errorMessage(error: unknown, notFoundIsMetric = true): string {
  if (error instanceof GlassnodeApiError) {
    if (error.status === 404 && notFoundIsMetric) return 'Error: 404 metric not found - correct metric endpoint selected?';
    if (error.status === 429) return 'Error: 429 rate limit exceeded - too many requests to the Glassnode API';
    return 'Error: HTTP error! status: ' + (error.status || 'unknown');
  }
  if (error instanceof GlassnodeNetworkError) {
    return 'Error: HTTP error! status: unknown';
  }
  if (error instanceof GlassnodeValidationError) {
    return 'Error: Invalid response format';
  }
  return 'Error: ' + (error instanceof Error ? error.message : 'Unknown error');
}

export function isDevEnv() {
  return window?.location?.hostname === 'localhost';
}

/**
 * Helper function to get the appropriate API base URL based on environment
 * @returns The base API URL for the current environment
 */
export function getApiUrl(): string {
  const isDevelopment = isDevEnv();

  if (isDevelopment) {
    // The client needs an absolute URL; the dev server proxies this path to the API
    return `${window.location.origin}/api/glassnode`;
  } else {
    return 'https://api.glassnode.com';
  }
}

export function buildCacheId(filteredParams: object, metric: string) {
  const relevantParams = {...filteredParams, metric};
  const sorted = Object.keys(relevantParams)
    .sort()
    .reduce((obj, key) => {
      obj[key] = relevantParams[key];
      return obj;
    }, {});
  return `metrics-${JSON.stringify(sorted)}`;
}

export function buildHeaders(objectData: {t: number; o: Record<string, unknown>}[], metric: string) {
  const allKeys = new Set<string>();
  for (const item of objectData) {
    if (item.o) {
      Object.keys(item.o).forEach(k => allKeys.add(k));
    }
  }
  const keys = Array.from(allKeys).sort();
  const headers = ['Date', ...keys.map(it => `${metric}.${it}`)];
  return {headers, keys};
}
