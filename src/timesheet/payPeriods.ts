export interface PayPeriod {
  year: number;
  number: number;
  /** yyyy-mm-dd */
  from: string;
  /** yyyy-mm-dd */
  to: string;
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** 'yyyy-mm-dd' → 'Nov 5'. String split, not Date, to avoid time zone shifts. */
function shortDate(iso: string): string {
  const [, m, d] = iso.split('-').map(Number);
  return `${MONTHS[m - 1]} ${d}`;
}

export function periodLabel(p: PayPeriod): string {
  return `${p.year} – PP ${p.number} (${shortDate(p.from)} – ${shortDate(p.to)})`;
}

/** Periods that have started by today, newest first. */
export function selectablePeriods(all: PayPeriod[], todayIso: string): PayPeriod[] {
  return all.filter((p) => p.from <= todayIso).sort((a, b) => (a.from < b.from ? 1 : a.from > b.from ? -1 : 0));
}

export function findPeriod(all: PayPeriod[], fromIso: string, toIso: string): PayPeriod | undefined {
  return all.filter((p) => p.from === fromIso && p.to === toIso)[0];
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}/;

/**
 * Maps a Pay Period Schedule list row to a PayPeriod. Date-only fields come back as an ISO
 * timestamp; the first 10 characters are the calendar day (Peak's sites are west of UTC).
 */
export function toPayPeriod(row: { Year: number; PeriodNumber: number; FromDate: string; ToDate: string }): PayPeriod {
  const ok =
    isFinite(Number(row.Year)) &&
    isFinite(Number(row.PeriodNumber)) &&
    ISO_DATE.test(String(row.FromDate)) &&
    ISO_DATE.test(String(row.ToDate));
  if (!ok) throw new Error(`Bad pay period row in the schedule list: ${JSON.stringify(row)}`);
  return { year: Number(row.Year), number: Number(row.PeriodNumber), from: row.FromDate.slice(0, 10), to: row.ToDate.slice(0, 10) };
}

/** Local calendar date as 'yyyy-mm-dd' (toISOString would give the UTC date). */
export function todayIso(): string {
  const d = new Date();
  const pad = (n: number): string => (n < 10 ? '0' : '') + n;
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** 'yyyy-mm-dd' → 'dd/mm/yy' by string slicing; '' stays ''. */
export function displayDate(iso: string): string {
  return iso ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(2, 4)}` : '';
}
