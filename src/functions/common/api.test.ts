import {cachePrefixFor, clientSettings, fetchWithSource, resetApi} from './api';
import {METRIC as METRIC_IMPL} from './metrics';
import {mockFetch} from './testFetch';

describe('cachePrefixFor', () => {
  it('keeps the prefix while the key stays the same', async () => {
    const first = await cachePrefixFor('key-a');
    expect(await cachePrefixFor('key-a')).toBe(first);
  });

  it('gives a new prefix when the key changes', async () => {
    const a = await cachePrefixFor('key-a');
    const b = await cachePrefixFor('key-b');
    expect(b).not.toBe(a);
  });

  it('never gives overlapping calls with different keys the same prefix', async () => {
    await cachePrefixFor('key-a');

    const [b, a] = await Promise.all([cachePrefixFor('key-b'), cachePrefixFor('key-a')]);

    expect(a).not.toBe(b);
  });
});

describe('client requests', () => {
  const path = '/api/glassnode/v1/metrics/addresses/active_count';
  let mock: ReturnType<typeof mockFetch>;

  beforeEach(() => {
    clientSettings.retryDelay = 1;
    mock = mockFetch();
    (global as any).OfficeRuntime.storage.getItem.mockReturnValue('secret-key-123');
    resetApi();
  });

  afterEach(() => {
    mock.restore();
  });

  it('adds source=excel-add-in to the URL and keeps the key in the X-Api-Key header only', async () => {
    mock.onGet(path).reply(200, [{t: 1640995200, v: 1}]);

    await METRIC_IMPL('BTC', '/addresses/active_count', new Date('2022-01-01T00:00:00Z'), null, 'e=binance');

    const request = mock.history.get[0];
    expect(request.params.source).toBe('excel-add-in');
    expect(request.params.e).toBe('binance');
    expect(request.headers['X-Api-Key']).toBe('secret-key-123');
    expect(request.fullUrl).not.toContain('secret-key-123');
    expect(request.fullUrl).not.toContain('api_key');
    // Redirects are not followed, so the key header cannot be forwarded elsewhere
    expect(request.init?.redirect).toBe('manual');
  });

  it('fetchWithSource passes headers and options through untouched', async () => {
    mock.onGet(path).reply(200, []);
    const init = {headers: {'X-Api-Key': 'k'}, redirect: 'manual' as const};

    await fetchWithSource('https://example.test' + path + '?a=BTC', init);
    await fetchWithSource('https://example.test' + path + '?', init);

    expect(mock.fetch.mock.calls[0]).toEqual(['https://example.test' + path + '?a=BTC&source=excel-add-in', init]);
    expect(mock.fetch.mock.calls[0][1]).toBe(init);
    expect(mock.fetch.mock.calls[1][0]).toBe('https://example.test' + path + '?source=excel-add-in');
  });

  it('blocks "." and ".." in the metric path before any request', async () => {
    for (const metric of ['/addresses/../x', '/addresses/./x', '/addresses/active.count']) {
      expect(await METRIC_IMPL('BTC', metric, new Date('2022-01-01T00:00:00Z'))).toEqual([
        ['Error: Invalid path, make sure to use API endpoint notation like /addresses/active_count'],
      ]);
    }
    expect(mock.fetch).not.toHaveBeenCalled();
  });

  it('reports a rejected parameter such as f=xml as an error, without a request', async () => {
    const result = await METRIC_IMPL('BTC', '/addresses/active_count', new Date('2022-01-01T00:00:00Z'), null, 'f=xml');

    expect(result[0][0]).toMatch(/^Error: .*f=/);
    expect(mock.fetch).not.toHaveBeenCalled();
  });

  it('maps a network failure to the unknown-status error', async () => {
    mock.onGet(path).networkError();

    expect(await METRIC_IMPL('BTC', '/addresses/active_count', new Date('2022-01-01T00:00:00Z'))).toEqual([['Error: HTTP error! status: unknown']]);
  });

  it('keeps a failed request out of the cache and shares one request between identical calls', async () => {
    mock.onGet(path).replyOnce(async () => [500]);
    mock.onGet(path).replyOnce(async () => [500]);
    mock.onGet(path).replyOnce(async () => [500]);
    mock.onGet(path).reply(200, [{t: 1640995200, v: 7}]);
    const start = new Date('2022-01-01T00:00:00Z');

    expect(await METRIC_IMPL('BTC', '/addresses/active_count', start)).toEqual([['Error: HTTP error! status: 500']]);
    const [a, b] = await Promise.all([METRIC_IMPL('BTC', '/addresses/active_count', start), METRIC_IMPL('BTC', '/addresses/active_count', start)]);

    expect(a).toEqual([[7]]);
    expect(b).toEqual([[7]]);
    expect(mock.history.get).toHaveLength(4);
  });

  it('never runs more than 6 requests at once', async () => {
    let active = 0;
    let peak = 0;
    let started = 0;
    const releases: Array<() => void> = [];
    mock.fetch.mockImplementation((async () => {
      active += 1;
      started += 1;
      peak = Math.max(peak, active);
      await new Promise<void>(resolve => releases.push(resolve));
      active -= 1;
      return {ok: true, status: 200, statusText: 'OK', headers: {get: () => null}, json: async () => [{t: 1, v: 1}], text: async () => ''};
    }) as any);

    const calls = Array.from({length: 10}, (_, i) => METRIC_IMPL('BTC', '/addresses/active_count', new Date('2022-01-01T00:00:00Z'), null, `e=ex${i}`));
    await new Promise(resolve => setTimeout(resolve, 20));
    expect(peak).toBe(6);

    while (started < 10 || active > 0) {
      releases.splice(0).forEach(release => release());
      await new Promise(resolve => setTimeout(resolve, 5));
    }
    const results = await Promise.all(calls);
    expect(results).toHaveLength(10);
    expect(peak).toBe(6);
  });
});
