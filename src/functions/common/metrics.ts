/* global console */

import axios from 'axios';
import {apiClient, cachePrefixFor} from './api';
import {API_KEY_MISSING_ERROR, API_KEY_STORAGE_ERROR, apiKeyHeaders, buildCacheId, buildHeaders, describeError, getApiKey, getApiUrl} from './utils';

import {dateToUnixSeconds, unixSecondsToYyyyMmDd} from './dateUtils';

// Parameters a formula may not set: the key only ever travels in the X-Api-Key header, and
// a, s, u and source come from METRIC's own arguments. i (resolution) stays settable.
const RESERVED_PARAMS = new Set(['api_key', 'a', 's', 'u', 'source', 'metric']);
const METRIC_SEGMENT = /^[A-Za-z0-9_-]+$/;

/** An API value as a cell value: numbers, strings and booleans as they are, anything else as text. */
function toCell(value: unknown): unknown {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') return value;
  return JSON.stringify(value);
}

/** A key of the object itself, never an inherited one such as constructor or __proto__. */
function ownValue(o: Record<string, unknown> | undefined, key: string): unknown {
  return o && Object.prototype.hasOwnProperty.call(o, key) ? o[key] : undefined;
}

export async function METRIC(
  asset: string,
  metric: string,
  startDate: Date,
  endDate: Date | null = null,
  parameter1: string | null = null,
  parameter2: string | null = null,
  parameter3: string | null = null,
  parameter4: string | null = null,
  pick: string | null = null,
): Promise<string[][]> {
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
      console.log('API key not found');
      return [[API_KEY_MISSING_ERROR]];
    }

    if (metric.includes('?') || metric.includes('#')) {
      return [['Error: the metric path cannot contain "?" or "#"; pass parameters as key=value arguments']];
    }

    // Only plain path segments (letters, digits, "_" and "-"): this blocks "..", ".", "%2e" and "\"
    // so a formula cannot point the request, and the API key with it, at another endpoint.
    const segments = metric.replace(/^\//, '').split('/');
    if (!segments.every(segment => METRIC_SEGMENT.test(segment))) {
      return [['Error: Invalid path, make sure to use API endpoint notation like /addresses/active_count']];
    }

    // Convert provided Date objects to unix timestamps (seconds)
    const startTimestamp = dateToUnixSeconds(startDate);
    const endTimestamp = endDate != null ? dateToUnixSeconds(endDate) : null;

    // Build URL parameters
    const params = {
      a: asset,
      i: '24h',
      s: startTimestamp.toString(),
    };

    if (endTimestamp !== null) {
      params['u'] = endTimestamp.toString();
    }

    // Process optional parameters
    const optionalParams = [parameter1, parameter2, parameter3, parameter4].filter(Boolean);

    for (const param of optionalParams) {
      if (param && typeof param === 'string') {
        // Split on the first "=" only, so a value may itself contain "="
        const eq = param.indexOf('=');
        const key = eq === -1 ? param.trim() : param.slice(0, eq).trim();
        const value = eq === -1 ? '' : param.slice(eq + 1).trim();
        const reserved = key.toLowerCase();
        if (reserved === 'api_key') {
          return [['Error: api_key cannot be passed as a parameter; set your API key in the task pane']];
        }
        if (RESERVED_PARAMS.has(reserved)) {
          return [[`Error: ${key} cannot be passed as a parameter; use METRIC's own arguments`]];
        }
        if (key && value) {
          params[key] = value;
        } else {
          // Log only the key: the value may be something the user didn't mean to share
          console.warn('Invalid parameter format, expected key=value:', key || '(empty)');
        }
      }
    }

    // Use proxy path for development, direct API for production
    const apiUrl = `${getApiUrl()}/v1/metrics/${segments.map(encodeURIComponent).join('/')}`;

    const cacheId = (await cachePrefixFor(apiKey)) + buildCacheId(params, metric);
    type ObjectResponseItem = {t: number; o: Record<string, unknown>};
    type ValueResponseItem = {t: number; v: unknown};
    type ResponseItem = ObjectResponseItem | ValueResponseItem;
    const response = await apiClient.get<ResponseItem[]>(apiUrl, {
      params: {
        ...params,
        source: 'excel-add-in',
      },
      headers: apiKeyHeaders(apiKey),
      cache: {
        ttl: 3 * 60 * 1000, // 3 min
      },
      id: cacheId,
    });

    // If endDate is not provided and no pick is requested, limit to the first data point
    const rawRows: ResponseItem[] = Array.isArray(response.data) ? response.data : [];
    const rows: ResponseItem[] = endTimestamp == null ? rawRows.slice(0, 1) : rawRows;

    if (rows.length === 0) {
      return [['No data available']];
    }

    const responseType: 'value' | 'object' | null = (() => {
      const first = rows[0];
      if ('v' in first) return 'value';
      if ('o' in first) return 'object';
      return null;
    })();

    if (responseType === null) {
      console.log('Invalid response format - Metric not supported:', rows);
      throw new Error('Invalid response format - Metric not supported');
    }

    if (responseType === 'value') {
      const valueData = rows as ValueResponseItem[];
      if (valueData.length === 1) {
        return [[toCell(valueData[0].v)]] as unknown as string[][];
      }
      // Return table format with headers
      const headers = ['Date', metric];

      const dataRows = valueData.map(item => [unixSecondsToYyyyMmDd(item.t), toCell(item.v)]);

      return [headers, ...dataRows] as unknown as string[][];
    }

    if (responseType === 'object') {
      const objectData = rows as ObjectResponseItem[];
      if (pick !== null) {
        const multi = objectData.length > 1;
        if (!multi) {
          // Single data point: return just the picked value
          const only = objectData[0];
          return [[toCell(ownValue(only.o, pick))]] as unknown as string[][];
        }

        // Multiple data points: return Date + picked column
        const headers = ['Date', `${metric}.${pick}`];
        const dataRows = objectData.map(item => [unixSecondsToYyyyMmDd(item.t), toCell(ownValue(item.o, pick))]);
        return [headers, ...dataRows] as unknown as string[][];
      }

      // Return table format with headers
      const {headers, keys} = buildHeaders(objectData, metric);
      const dataRows = objectData.map(item => [unixSecondsToYyyyMmDd(item.t), ...keys.map(k => toCell(ownValue(item.o, k)))]);
      return [headers, ...dataRows] as unknown as string[][];
    }
  } catch (error) {
    console.error('Error in METRIC function:', describeError(error));

    // Handle axios-specific errors
    if (axios.isAxiosError(error)) {
      if (error.response?.status === 404) {
        return [['Error: 404 metric not found - correct metric endpoint selected?']];
      }
      if (error.response?.status === 429) {
        return [['Error: 429 rate limit exceeded - too many requests to the Glassnode API']];
      }
      return [['Error: HTTP error! status: ' + (error.response?.status || 'unknown')]];
    }

    // Return error message in Excel-compatible format
    return [['Error: ' + (error instanceof Error ? error.message : 'Unknown error')]];
  }
}
