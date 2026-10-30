/* eslint-disable @rushstack/no-new-null -- null marks a blank computed cell, as in the template */
import { TEMPLATES } from './definitions';
import { Line, TimesheetType } from './types';
import { parseTime, timeToDayFraction } from './values';

const round2 = (n: number): number => Math.round(n * 100) / 100;

const num = (v: Line[string]): number => (typeof v === 'number' ? v : 0);

const has = (v: Line[string]): boolean => v !== null && v !== undefined && v !== '';

/** Hours between two times as typed; null unless both parse. Wraps past midnight when asked (Medic). */
function hoursBetween(start: Line[string], end: Line[string], wrap: boolean): number | null {
  const s = parseTime(String(start ?? ''));
  const e = parseTime(String(end ?? ''));
  if (!s || !e) return null;
  let diff = timeToDayFraction(e) - timeToDayFraction(s);
  if (wrap) diff = ((diff % 1) + 1) % 1;
  return diff * 24;
}

/** Returns the line with its computed columns filled. A row with no inputs computes to null. */
export function computeLine(type: TimesheetType, line: Line): Line {
  const out: Line = { ...line };
  const computedKeys = TEMPLATES[type].columns.filter((c) => c.kind === 'computed').map((c) => c.key);
  const isEmpty = TEMPLATES[type].columns.every((c) => c.kind === 'computed' || !has(line[c.key]));
  if (isEmpty) {
    computedKeys.forEach((k) => (out[k] = null));
    return out;
  }

  if (type === 'Admin') {
    const span = hoursBetween(line.start, line.end, false);
    const hours = span === null ? null : round2(span - num(line.deductions));
    out.hours = hours;
    out.amount = round2(num(hours) * num(line.rate));
  } else if (type === 'Instructor') {
    out.instrTotal = round2(num(line.days) * num(line.dayRate));
    // Template: =IF(OR(G="",H=""),"",G*H)
    out.adminTotal = has(line.adminHours) && has(line.adminRate) ? round2(num(line.adminHours) * num(line.adminRate)) : null;
    out.otherTotal = round2(num(line.travel) + num(line.perDiem));
    out.amount = round2(num(out.instrTotal) + num(out.adminTotal) + num(out.otherTotal));
  } else {
    const span = hoursBetween(line.start, line.end, true);
    out.hours = span === null ? null : round2(span);
    out.medicTotal = round2(num(line.dayRate));
    out.otherTotal = round2(num(line.travel) + num(line.perDiem));
    out.amount = round2(num(out.medicTotal) + num(out.otherTotal));
  }
  return out;
}

/** Sums each totals key over the computed lines. */
export function computeTotals(type: TimesheetType, lines: Line[]): Record<string, number> {
  const computed = lines.map((l) => computeLine(type, l));
  const totals: Record<string, number> = {};
  TEMPLATES[type].totals.forEach(({ key }) => {
    totals[key] = round2(computed.reduce((sum, l) => sum + num(l[key]), 0));
  });
  return totals;
}
