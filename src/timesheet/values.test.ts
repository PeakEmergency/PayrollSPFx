import { normalizeDate, parseTime, timeToDayFraction, toExcelDate } from './values';

describe('parseTime', () => {
  it.each([
    ['8:30', '08:30'],
    ['08:30', '08:30'],
    ['0830', '08:30'],
    ['8:30 pm', '20:30'],
    ['12:15 am', '00:15'],
    ['25:00', null],
    ['abc', null],
    ['', null]
  ])('%p -> %p', (input, expected) => {
    expect(parseTime(input)).toBe(expected);
  });
});

describe('normalizeDate', () => {
  it.each([
    ['2026-10-15 00:00:00', '2026-10-15'],
    ['2026-10-15', '2026-10-15'],
    ['', null]
  ])('%p -> %p', (input, expected) => {
    expect(normalizeDate(input)).toBe(expected);
  });
});

it('toExcelDate is UTC midnight of the same day', () => {
  expect(toExcelDate('2026-10-15').toISOString()).toBe('2026-10-15T00:00:00.000Z');
});

it('timeToDayFraction', () => {
  expect(timeToDayFraction('06:00')).toBe(0.25);
});
