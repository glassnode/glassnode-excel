/* global console */

import {API_KEY_MISSING_ERROR, API_KEY_STORAGE_ERROR, describeError, errorMessage, getApiKey} from './utils';
import {cached, cachePrefixFor, getClient} from './api';

export async function ASSETS(limit: number = null): Promise<string[][]> {
  if (limit !== null && limit !== undefined && !(Number.isInteger(limit) && limit >= 0)) {
    return [['Error: limit must be a whole number of 0 or more']];
  }
  try {
    // Get API key from settings
    let apiKey: string | null;
    try {
      apiKey = await getApiKey();
    } catch (error) {
      console.error('Error reading API key from storage:', describeError(error));
      return [[API_KEY_STORAGE_ERROR]];
    }
    if (!apiKey) {
      return [[API_KEY_MISSING_ERROR]];
    }

    const cacheId = (await cachePrefixFor(apiKey)) + 'assets_cache';
    const assets = await cached(cacheId, 60 * 60 * 1000 /* 1 hour */, () => getClient(apiKey).getAssetMetadata());

    // Extract IDs and limit the results
    const assetIds = assets
      .slice(0, limit ?? 50_000)
      .map(asset => asset.id)
      .filter(id => id); // Filter out any undefined/null/"" IDs

    // Return as 2D array for Excel (each ID in its own row)
    return assetIds.map(id => [id]);
  } catch (error) {
    console.error('Error fetching asset data:', describeError(error));

    return [[errorMessage(error, false)]];
  }
}
