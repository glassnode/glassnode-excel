import {cachePrefixFor} from './api';

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
