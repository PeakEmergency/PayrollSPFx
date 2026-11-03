/* eslint-disable @rushstack/no-new-null -- null explicitly represents unavailable calculated totals in the shared domain contract. */
import Decimal from 'decimal.js-light';
import type { ExpenseReport, ReportTotals } from './models';

/** Plain signed decimals only: no implicit currency, grouping, or exponent coercion. */
export function isNumericInput(raw: string, allowBlank: boolean = true): boolean {
  const value = raw.trim();
  return (allowBlank && value === '') || /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(value);
}

function decimalText(raw: string): string {
  return raw.trim().replace(/^\+/, '') || '0';
}

function decimalType(values: string[]): typeof Decimal {
  // Multiplication/addition are the only operations. Counting every input digit
  // provides sufficient working precision even for unusually precise decimals.
  const precision = values.reduce((count, value) => count + value.replace(/\D/g, '').length, 32);
  return Decimal.clone({ precision, rounding: Decimal.ROUND_HALF_UP });
}

export function calculateReport(report: ExpenseReport): ReportTotals {
  const DecimalValue = decimalType(report.lines.reduce<string[]>((values, line) =>
    values.concat(line.kilometres, line.rate, line.gst, line.cost), []));
  const parse = (raw: string, allowBlank: boolean = true): Decimal | null =>
    isNumericInput(raw, allowBlank) ? new DecimalValue(decimalText(raw)) : null;
  const parsed = report.lines.map(line => {
    const kilometres = parse(line.kilometres);
    const rate = parse(line.rate, false);
    const gst = parse(line.gst);
    const cost = parse(line.cost);
    const mileage = kilometres !== null && rate !== null ? kilometres.times(rate) : null;
    const amount = mileage !== null && gst !== null && cost !== null ? mileage.plus(gst).plus(cost) : null;
    return { kilometres, mileage, gst, cost, amount };
  });
  const sum = (values: (Decimal | null)[]): string | null => {
    let total = new DecimalValue(0);
    for (const value of values) {
      if (value === null) return null;
      total = total.plus(value);
    }
    return total.toFixed();
  };
  return {
    lines: parsed.map(line => ({ mileage: line.mileage === null ? null : line.mileage.toFixed(), amount: line.amount === null ? null : line.amount.toFixed() })),
    kilometres: sum(parsed.map(line => line.kilometres)),
    mileage: sum(parsed.map(line => line.mileage)),
    gst: sum(parsed.map(line => line.gst)),
    cost: sum(parsed.map(line => line.cost)),
    amount: sum(parsed.map(line => line.amount))
  };
}

function groupFixed(value: string): string {
  const parts = value.split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}

export function formatCurrency(value: string | null): string {
  if (value === null || !isNumericInput(value, false)) return '—';
  const DecimalValue = decimalType([value]);
  const rounded = new DecimalValue(decimalText(value)).toFixed(2, Decimal.ROUND_HALF_UP);
  return rounded.charAt(0) === '-' ? '-$' + groupFixed(rounded.slice(1)) : '$' + groupFixed(rounded);
}

export function formatQuantity(value: string | null): string {
  if (value === null || !isNumericInput(value, false)) return '—';
  const DecimalValue = decimalType([value]);
  return groupFixed(new DecimalValue(decimalText(value)).toFixed());
}
