import { computeLine, computeTotals } from './calc';
import { Line } from './types';

const medic: Line = { date: '2026-10-15', days: 1, dayRate: 450, travel: 100, perDiem: 47 };

it('admin line', () => {
  const out = computeLine('Admin', { date: '2026-10-15', start: '08:00', end: '16:30', deductions: 0.5, rate: 20 });
  expect(out.hours).toBe(8);
  expect(out.amount).toBe(160);
});

it('instructor line', () => {
  const out = computeLine('Instructor', {
    date: '2026-10-15', days: 0.5, dayRate: 280, adminHours: 2, adminRate: 25, travel: 100, perDiem: 47
  });
  expect(out).toMatchObject({ instrTotal: 140, adminTotal: 50, otherTotal: 147, amount: 337 });
});

it('medic total is days x rate', () => {
  expect(computeLine('Medic', { date: '2026-11-05', days: 0.5, dayRate: 350 }).medicTotal).toBe(175);
  expect(computeLine('Medic', medic)).toMatchObject({ medicTotal: 450, otherTotal: 147, amount: 597 });
});

it('empty row computes to null', () => {
  expect(computeLine('Medic', {})).toMatchObject({ medicTotal: null, otherTotal: null, amount: null });
});

it('computeTotals', () => {
  expect(computeTotals('Medic', [medic, {}])).toEqual({ days: 1, travel: 100, perDiem: 47, amount: 597 });
});
