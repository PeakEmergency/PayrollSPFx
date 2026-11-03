import { TEMPLATES } from './definitions';
import { Line, TimesheetType } from './types';

/** Course types with their own instructor rate; keys match the Pay Rates columns without the `Rate` prefix. */
export type CourseType = 'AP' | 'ASU' | 'NUECAFull' | 'NUECABridging' | 'NUECI' | 'NUECB';
const COURSE_TYPES: CourseType[] = ['AP', 'ASU', 'NUECAFull', 'NUECABridging', 'NUECI', 'NUECB'];

/** The signed-in person's row in the Pay Rates list. */
export interface PayRates {
  admin?: number;
  medicPSL?: number;
  course: Partial<Record<CourseType, number>>;
}

/** The cell each form's rate goes in. */
export const RATE_KEY: Record<TimesheetType, string> = { Admin: 'rate', Instructor: 'dayRate', Medic: 'dayRate' };

export const COURSE_EXAMPLES = 'AP-I-0182, ASU-AB-998, NUEC-A-443F, NUEC-A-468B, NUEC-I-195 or NUEC-B-238';

/** Course # formats from Peak (2026-11-03). NUEC-A ends in F (full) or B (bridging). */
export function courseType(course: string): CourseType | undefined {
  const c = course.trim().toUpperCase();
  if (/^AP-/.test(c)) return 'AP';
  if (/^ASU-/.test(c)) return 'ASU';
  const nuecA = /^NUEC-A-\d+([FB])$/.exec(c);
  if (nuecA) return nuecA[1] === 'F' ? 'NUECAFull' : 'NUECABridging';
  if (/^NUEC-I-/.test(c)) return 'NUECI';
  if (/^NUEC-B-/.test(c)) return 'NUECB';
  return undefined;
}

const num = (v: unknown): number | undefined => (typeof v === 'number' ? v : undefined);

/** Pay Rates list row → PayRates. */
export function toPayRates(row: Record<string, unknown>): PayRates {
  const rates: PayRates = { course: {} };
  if (num(row.AdminRate) !== undefined) rates.admin = num(row.AdminRate);
  if (num(row.MedicPSL) !== undefined) rates.medicPSL = num(row.MedicPSL);
  COURSE_TYPES.forEach((t) => {
    const v = num(row[`Rate${t}`]);
    if (v !== undefined) rates.course[t] = v;
  });
  return rates;
}

/** Medic PSL level (8–11) → the dropdown amount whose label names that level. */
function medicAmount(psl: number): number | undefined {
  const col = TEMPLATES.Medic.columns.filter((c) => c.key === 'dayRate')[0];
  const i = (col.optionLabels || []).findIndex((l) => l.indexOf(`PSL ${psl})`) >= 0);
  return i >= 0 ? (col.options || [])[i] : undefined;
}

export function hasRate(type: TimesheetType, rates: PayRates | undefined): boolean {
  if (!rates) return false;
  if (type === 'Admin') return rates.admin !== undefined;
  if (type === 'Medic') return rates.medicPSL !== undefined && medicAmount(rates.medicPSL) !== undefined;
  return COURSE_TYPES.some((t) => rates.course[t] !== undefined);
}

function rateFor(type: TimesheetType, line: Line, rates: PayRates): number | undefined {
  if (type === 'Admin') return rates.admin;
  if (type === 'Medic') return rates.medicPSL === undefined ? undefined : medicAmount(rates.medicPSL);
  const t = courseType(String(line.course ?? ''));
  return t ? rates.course[t] : undefined;
}

const blank = (v: unknown): boolean => v === null || v === undefined || String(v).trim() === '';

/**
 * Fills the row's rate from Pay Rates once the row has other input and the rate is blank.
 * An instructor's rate follows the course type, so editing the course # resets it.
 * A rate the user picked is otherwise kept (a shift can be paid at another level).
 */
export function fillRate(type: TimesheetType, line: Line, rates: PayRates, editedKey: string): Line {
  const key = RATE_KEY[type];
  if (!Object.keys(line).some((k) => k !== key && !blank(line[k]))) return line;
  if (!blank(line[key]) && !(type === 'Instructor' && editedKey === 'course')) return line;
  const rate = rateFor(type, line, rates);
  return rate === undefined ? line : { ...line, [key]: rate };
}
