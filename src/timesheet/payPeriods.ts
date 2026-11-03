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
