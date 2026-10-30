import { TEMPLATES } from './definitions';
import { TimesheetData } from './types';
import { normalizeDate, parseTime } from './values';

const has = (v: unknown): boolean => v !== null && v !== undefined && String(v).trim() !== '';

/** Returns user-facing messages; [] means the timesheet can be submitted. */
export function validate(data: TimesheetData): string[] {
  const errors: string[] = [];
  const def = TEMPLATES[data.type];
  const inputs = def.columns.filter((c) => c.kind !== 'computed');

  if (!data.name.trim()) errors.push('Your name is missing.');
  if (!data.from) errors.push('Pick the pay period From date.');
  if (!data.to) errors.push('Pick the pay period To date.');
  if (data.from && data.to && data.to < data.from) errors.push("The To date can't be before the From date.");

  let entered = 0;
  data.lines.forEach((line, i) => {
    if (!inputs.some((c) => has(line[c.key]))) return;
    entered++;
    const row = i + 1;
    if (!normalizeDate(String(line.date ?? ''))) errors.push(`Row ${row}: date missing.`);
    const times: Record<string, string | null> = {};
    inputs
      .filter((c) => c.kind === 'time' && has(line[c.key]))
      .forEach((c) => {
        const raw = String(line[c.key]);
        times[c.key] = parseTime(raw);
        if (!times[c.key]) errors.push(`Row ${row}: "${raw}" is not a valid time.`);
      });
    if (data.type === 'Admin' && times.start && times.end && times.end < times.start) {
      errors.push(`Row ${row}: end time is before start time.`);
    }
  });
  if (entered === 0) errors.push('Enter at least one row.');

  return errors;
}
