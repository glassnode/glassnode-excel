import {setupCache} from 'axios-cache-interceptor';
import axios from 'axios';

// due to issues with localstorage on Excel for Windows we use in-memory storage (default storage adapter)
export const apiClient = setupCache(axios, {
  interpretHeader: false, // Don't use HTTP cache headers
  methods: ['get'], // Only cache GET requests
  cachePredicate: {
    statusCheck: status => status >= 200 && status < 300,
  },
  // Needs to be disabled to prevent the pragma header being added by axios
  // read more here: https://axios-cache-interceptor.js.org/config/request-specifics#cache-cachetakeover
  // otherwise this at least fails in excel web, most likely also in other envs
  cacheTakeover: false,
});

// Cached responses belong to the API key they were fetched with. Each key change starts a new
// cache generation: IDs carry the generation, so a request still in flight under the old key
// writes where nothing reads, and the old entries are cleared.
let currentKey: string | null = null;
let generation = 0;

/** Returns the cache ID prefix for this key, starting a new generation if the key changed. */
export async function cachePrefixFor(apiKey: string): Promise<string> {
  // Update the state and take the prefix before any await, so overlapping calls with
  // different keys never get the same prefix
  const changed = currentKey !== null && currentKey !== apiKey;
  if (changed) {
    generation += 1;
  }
  currentKey = apiKey;
  const prefix = `g${generation}:`;
  if (changed) {
    await apiClient.storage.clear?.();
  }
  return prefix;
}
