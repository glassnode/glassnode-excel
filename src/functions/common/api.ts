import {GlassnodeAPI} from 'glassnode-api';
import {getApiUrl} from './utils';

/** Query parameter that tells the API this request comes from the add-in. */
const SOURCE_PARAM = 'source=excel-add-in';

/** Client settings. Tests lower the retry delay; the add-in never changes these. */
export const clientSettings = {maxRetries: 2, retryDelay: 1000};

/** At most this many API requests (retries included) run at the same time. */
const MAX_CONCURRENT_CALLS = 6;

/**
 * The client has no option for a `source` parameter, so it goes in here: every request URL gets
 * `source=excel-add-in` appended. Only the URL changes: `init` (the X-Api-Key header, redirect
 * mode, abort signal) goes to fetch as the client built it, and the key never enters the URL.
 * `fetch` is looked up on each call so a replaced global fetch is picked up.
 */
export function fetchWithSource(url: string | URL | Request, init?: RequestInit): Promise<Response> {
  if (typeof url !== 'string') {
    return globalThis.fetch(url, init);
  }
  // The client builds `endpoint?query`, with an empty query when there are no parameters
  const separator = url.endsWith('?') ? '' : url.includes('?') ? '&' : '?';
  return globalThis.fetch(url + separator + SOURCE_PARAM, init);
}

let client: GlassnodeAPI | null = null;
let clientKey: string | null = null;

/** The client for this API key. The key goes in the X-Api-Key header, never in the URL. */
export function getClient(apiKey: string): GlassnodeAPI {
  if (client === null || clientKey !== apiKey) {
    client = new GlassnodeAPI({
      apiKey,
      apiKeyLocation: 'header',
      // Must be absolute: the dev proxy path is resolved against the page origin
      apiUrl: getApiUrl(),
      fetch: fetchWithSource,
      maxRetries: clientSettings.maxRetries,
      retryDelay: clientSettings.retryDelay,
    });
    clientKey = apiKey;
  }
  return client;
}

// --- limiter: waits for a free slot, so retry waits inside a call keep their slot ---
let running = 0;
const waiting: Array<() => void> = [];

async function limited<T>(task: () => Promise<T>): Promise<T> {
  if (running >= MAX_CONCURRENT_CALLS) {
    await new Promise<void>(resolve => waiting.push(resolve));
  } else {
    running += 1;
  }
  try {
    return await task();
  } finally {
    // Hand the slot straight to the next waiter, or free it
    const next = waiting.shift();
    if (next) {
      next();
    } else {
      running -= 1;
    }
  }
}

// --- cache: in memory only (localStorage is unreliable in Excel for Windows) ---
type Entry = {promise: Promise<unknown>; expires: number};
const cache = new Map<string, Entry>();

/**
 * Runs `load` through the limiter and keeps its result for `ttl` ms. Concurrent calls with the
 * same id share one request. A failed request is never kept.
 */
export function cached<T>(id: string, ttl: number, load: () => Promise<T>): Promise<T> {
  const hit = cache.get(id);
  if (hit && hit.expires > Date.now()) {
    return hit.promise as Promise<T>;
  }
  const entry: Entry = {promise: undefined as unknown as Promise<unknown>, expires: Infinity};
  entry.promise = limited(load).then(
    value => {
      // Skip if the entry was cleared meanwhile (key change)
      if (cache.get(id) === entry) entry.expires = Date.now() + ttl;
      return value;
    },
    error => {
      if (cache.get(id) === entry) cache.delete(id);
      throw error;
    },
  );
  cache.set(id, entry);
  return entry.promise as Promise<T>;
}

// Cached responses belong to the API key they were fetched with. Each key change starts a new
// cache generation: IDs carry the generation, so a request still in flight under the old key
// writes where nothing reads, and the old entries are cleared.
let currentKey: string | null = null;
let generation = 0;

/** Returns the cache ID prefix for this key, starting a new generation if the key changed. */
export async function cachePrefixFor(apiKey: string): Promise<string> {
  const changed = currentKey !== null && currentKey !== apiKey;
  if (changed) {
    generation += 1;
    cache.clear();
  }
  currentKey = apiKey;
  return `g${generation}:`;
}

/** Drops all cached responses, the client and the key state. For tests. */
export function resetApi(): void {
  cache.clear();
  client = null;
  clientKey = null;
  currentKey = null;
  generation = 0;
}
