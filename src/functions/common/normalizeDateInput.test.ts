import {normalizeDateInput} from './normalizeDateInput';

describe('normalizeDateInput', () => {
  it('parses YYYY-MM-DD string', () => {
    const res = normalizeDateInput('2024-01-01');
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error('expected success');
    expect(res.date.toISOString().startsWith('2024-01-01')).toBe(true);
  });

  it('parses Excel serial number', () => {
    // 2024-01-01 = 45292
    const res = normalizeDateInput(45292);
    expect(res.ok).toBe(true);
    if (!res.ok) throw new Error('expected success');
    expect(res.date.toISOString().startsWith('2024-01-01')).toBe(true);
  });

  it('errors on invalid input', () => {
    const res = normalizeDateInput('not-a-date');
    expect(res.ok).toBe(false);
  });

  it.each(['2024-13-01', '2024-02-30', '2023-02-29', '2024-00-10'])('rejects the impossible date %s', input => {
    expect(normalizeDateInput(input).ok).toBe(false);
  });

  it('accepts a leap day', () => {
    const res = normalizeDateInput('2024-02-29');
    if (!res.ok) throw new Error('expected success');
    expect(res.date.toISOString().startsWith('2024-02-29')).toBe(true);
  });

  it.each([NaN, Infinity, -1, 0, 2958466])('rejects the serial number %p', input => {
    expect(normalizeDateInput(input).ok).toBe(false);
  });

  it('accepts the last date Excel supports', () => {
    const res = normalizeDateInput(2958465);
    if (!res.ok) throw new Error('expected success');
    expect(res.date.toISOString().startsWith('9999-12-31')).toBe(true);
  });

  it('accepts a serial number typed as text, and rejects other text', () => {
    expect(normalizeDateInput('45292').ok).toBe(true);
    expect(normalizeDateInput('45292abc').ok).toBe(false);
    expect(normalizeDateInput('1e5').ok).toBe(false);
  });
});
