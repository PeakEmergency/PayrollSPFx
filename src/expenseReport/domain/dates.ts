function pad(value: number, width: number = 2): string {
  return ('0000' + value).slice(-width);
}

function isCalendarDay(year: number, month: number, day: number): boolean {
  if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1) return false;
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return day <= days[month - 1];
}

function isoParts(iso: string): { year: number; month: number; day: number } | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!match) return undefined;
  const year = Number(match[1]); const month = Number(match[2]); const day = Number(match[3]);
  return isCalendarDay(year, month, day) ? { year, month, day } : undefined;
}

/** Excel's documented/common 2029 cutoff is fixed here, independent of OS locale. */
export function parseDateInput(raw: string): string | undefined {
  const match = /^(\d{2})\/(\d{2})\/(\d{2}|\d{4})$/.exec(raw.trim());
  if (!match) return undefined;
  const day = Number(match[1]); const month = Number(match[2]);
  let year = Number(match[3]);
  if (match[3].length === 2) year += year <= 29 ? 2000 : 1900;
  return isCalendarDay(year, month, day) ? pad(year, 4) + '-' + pad(month) + '-' + pad(day) : undefined;
}

export function formatDateInput(iso: string): string {
  const parts = isoParts(iso);
  if (!parts) return '';
  // Four digits outside this window prevent 2030 becoming 1930 on reopen.
  const year = parts.year >= 1930 && parts.year <= 2029 ? pad(parts.year % 100) : pad(parts.year, 4);
  return pad(parts.day) + '/' + pad(parts.month) + '/' + year;
}

export function toCalendarDate(raw: string): Date | undefined {
  const iso = parseDateInput(raw);
  const parts = iso ? isoParts(iso) : undefined;
  if (!parts) return undefined;
  const date = new Date(0);
  date.setFullYear(parts.year, parts.month - 1, parts.day);
  date.setHours(12, 0, 0, 0);
  return date;
}

export function fromCalendarDate(date: Date): string {
  if (!Number.isFinite(date.getTime())) return '';
  const year = date.getFullYear(); const month = date.getMonth() + 1; const day = date.getDate();
  if (!isCalendarDay(year, month, day)) return '';
  return formatDateInput(pad(year, 4) + '-' + pad(month) + '-' + pad(day));
}

export function calendarStartDate(raw: string, other?: string): Date {
  return toCalendarDate(raw) || (other ? toCalendarDate(other) : undefined) || new Date();
}
