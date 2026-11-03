import { findPeriod, periodLabel, selectablePeriods, type PayPeriod } from './payPeriods';

const pp = (year: number, number: number, from: string, to: string): PayPeriod => ({ year, number, from, to });
const p23 = pp(2026, 23, '2026-11-05', '2026-11-18');
const p24 = pp(2026, 24, '2026-11-19', '2026-12-02');
const p27_1 = pp(2027, 1, '2026-12-31', '2027-01-13');

describe('payPeriods', () => {
  it('labels a period', () => expect(periodLabel(p23)).toBe('2026 – PP 23 (Nov 5 – Nov 18)'));
  it('labels a period that crosses the year', () => expect(periodLabel(p27_1)).toBe('2027 – PP 1 (Dec 31 – Jan 13)'));
  it('hides future periods and sorts newest first', () =>
    expect(selectablePeriods([p23, p27_1, p24], '2026-11-20')).toEqual([p24, p23]));
  it('includes a period starting today', () => expect(selectablePeriods([p24], '2026-11-19')).toEqual([p24]));
  it('orders across the year boundary', () =>
    expect(selectablePeriods([p27_1, p24], '2027-01-02').map((p) => p.year)).toEqual([2027, 2026]));
  it('empty list stays empty', () => expect(selectablePeriods([], '2026-11-20')).toEqual([]));
  it('finds by dates', () => expect(findPeriod([p23, p24], '2026-11-05', '2026-11-18')).toBe(p23));
  it('no match is undefined', () => expect(findPeriod([p23], '2026-11-06', '2026-11-18')).toBeUndefined());
});
