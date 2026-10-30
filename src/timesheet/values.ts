/* eslint-disable @rushstack/no-new-null -- null means "not a valid value" */
const pad = (n: number): string => (n < 10 ? '0' : '') + n;

/** Accepts 8:30, 08:30, 0830, 8:30 pm. Returns 'HH:MM' (24h) or null. */
export function parseTime(input: string): string | null {
  const m = /^(\d{1,2}):?(\d{2})\s*([ap]\.?m\.?)?$/i.exec((input || '').trim());
  if (!m) return null;
  let h = Number(m[1]);
  const min = Number(m[2]);
  if (min > 59) return null;
  if (m[3]) {
    if (h < 1 || h > 12) return null;
    h = (h % 12) + (m[3][0].toLowerCase() === 'p' ? 12 : 0);
  } else if (h > 23) {
    return null;
  }
  return `${pad(h)}:${pad(min)}`;
}

/** Grid date value ('yyyy-mm-dd' with optional time) → 'yyyy-mm-dd', or null. */
export function normalizeDate(input: string): string | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec((input || '').trim());
  if (!m) return null;
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  return d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3] ? m[0] : null;
}

/** UTC midnight, so ExcelJS writes the serial for the same calendar day in any time zone. */
export function toExcelDate(isoDate: string): Date {
  const [y, mo, d] = isoDate.split('-').map(Number);
  return new Date(Date.UTC(y, mo - 1, d));
}

export function timeToDayFraction(hhmm: string): number {
  const [h, m] = hhmm.split(':').map(Number);
  return (h * 60 + m) / 1440;
}
