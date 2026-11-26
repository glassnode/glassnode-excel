/* global OfficeRuntime */

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
 * A short, non-reversible tag for an API key, used to keep cached responses per key
 * so switching keys never serves data fetched with the previous one.
 */
export function apiKeyCacheTag(apiKey: string): string {
  // FNV-1a, 32 bit
  let hash = 0x811c9dc5;
  for (let i = 0; i < apiKey.length; i++) {
    hash ^= apiKey.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16);
}

/**
 * A loggable summary of an error. Never log a whole axios error: its request config
 * carries the X-Api-Key header.
 */
export function describeError(error: unknown): Record<string, unknown> {
  if (error && typeof error === 'object' && (error as {isAxiosError?: boolean}).isAxiosError) {
    const axiosError = error as {code?: string; message?: string; response?: {status?: number}};
    return {code: axiosError.code, status: axiosError.response?.status, message: axiosError.message};
  }
  return {message: error instanceof Error ? error.message : String(error)};
}

/**
 * Request headers that authenticate with the Glassnode API.
 * The key goes in a header rather than the query string, so it stays out of URLs.
 */
export function apiKeyHeaders(apiKey: string): Record<string, string> {
  return {'X-Api-Key': apiKey};
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
    return '/api/glassnode';
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
