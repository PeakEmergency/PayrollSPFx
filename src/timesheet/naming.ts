/* eslint-disable @rushstack/no-new-null -- Graph returns null for unset givenName/surname */
import { TimesheetData } from './types';

/** 'LAST, FIRST' in uppercase. Without a surname, the last word of displayName is used. */
export function formatEmployeeName(givenName: string | null, surname: string | null, displayName: string): string {
  let first = (givenName || '').trim();
  let last = (surname || '').trim();
  if (!last) {
    const words = displayName.trim().split(/\s+/);
    last = words.pop() || '';
    first = first || words.join(' ');
  }
  return (first ? `${last}, ${first}` : last).toUpperCase();
}

/** 'yyyy-mm-dd' → 'dd/mm/yy'. */
export const ddmmyy = (iso: string | null): string => {
  const [y, m, d] = (iso || '').split('-');
  return `${d}/${m}/${(y || '').slice(2)}`;
};

export function buildFileName(data: TimesheetData): string {
  return `${data.type} Timesheet - ${data.name} - ${data.from}_${data.to}.xlsx`;
}

export function buildSubject(data: TimesheetData): string {
  return `${data.type} Timesheet from ${data.name} – ${ddmmyy(data.from)} to ${ddmmyy(data.to)}`;
}
