/**
 * Test helper: replaces global fetch with a scripted API. The reply API follows the axios-mock-adapter style
 * (onGet / reply / replyOnce / networkError / history.get) so tests read the same as before.
 * Responses are plain objects: the jest jsdom environment has no Response.
 */
type Reply = [number, unknown?] | Promise<[number, unknown?]>;
type Handler = {match: string; once: boolean; used: boolean; respond: () => Reply | 'network-error'};
export type RecordedRequest = {url: string; fullUrl: string; params: Record<string, string>; headers: Record<string, string>; init?: RequestInit};

export function mockFetch() {
  const handlers: Handler[] = [];
  const history = {get: [] as RecordedRequest[]};

  const fake = jest.fn(async (input: string, init?: RequestInit) => {
    const parsed = new URL(input, 'http://localhost');
    history.get.push({
      url: parsed.pathname,
      fullUrl: input,
      params: Object.fromEntries(parsed.searchParams.entries()),
      headers: {...((init?.headers as Record<string, string>) ?? {})},
      init,
    });
    const handler = handlers.find(h => h.match === parsed.pathname && !(h.once && h.used));
    if (!handler) {
      return respond(404, undefined);
    }
    handler.used = true;
    const result = handler.respond();
    if (result === 'network-error') {
      throw new TypeError('Failed to fetch');
    }
    const [status, data] = await result;
    return respond(status, data);
  });
  const original = global.fetch;
  global.fetch = fake as unknown as typeof fetch;

  const on = (match: string) => ({
    reply: (status: number, data?: unknown) => void handlers.push({match, once: false, used: false, respond: () => [status, data]}),
    replyOnce: (fn: () => Promise<[number, unknown?]>) => void handlers.push({match, once: true, used: false, respond: fn}),
    networkError: () => void handlers.push({match, once: false, used: false, respond: () => 'network-error'}),
  });

  return {onGet: on, history, fetch: fake, restore: () => void (global.fetch = original)};
}

function respond(status: number, data: unknown) {
  const text = data === undefined ? '' : JSON.stringify(data);
  return {
    ok: status >= 200 && status < 300,
    status,
    statusText: String(status),
    headers: {get: () => null},
    json: async () => JSON.parse(text),
    text: async () => text,
  };
}
