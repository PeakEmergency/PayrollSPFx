import { calculateReport, formatCurrency } from './calculations';
import { parseDateInput } from './dates';
import type { ExpenseReport } from './models';

const escapeHtml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const ENTRY_KEYS = ['date', 'startCity', 'endCity', 'kilometres', 'description', 'gst', 'cost'] as const;

/** True when any row has something typed in it; the prefilled rate does not count. */
export function hasEntries(report: ExpenseReport): boolean {
  return report.lines.some(line => ENTRY_KEYS.some(key => line[key].trim() !== ''));
}

export function expenseFileName(report: ExpenseReport, id: number): string {
  return `Expense Report - ${report.instructorName.trim()} - ${parseDateInput(report.fromDate)}_${parseDateInput(report.toDate)} - ${id}.xlsx`;
}

export function expenseSubject(report: ExpenseReport): string {
  return `Expense Report from ${report.instructorName.trim()} – ${report.fromDate.trim()} to ${report.toDate.trim()}`;
}

export function expenseBody(report: ExpenseReport, senderEmail: string, receiptCount: number): string {
  return [
    `Expense report submitted by <b>${escapeHtml(report.instructorName.trim())}</b> (${escapeHtml(senderEmail)}).`,
    `Pay period: ${report.fromDate.trim()} to ${report.toDate.trim()}`,
    `Total: ${formatCurrency(calculateReport(report).amount)}`,
    `Receipts: ${receiptCount}`
  ].join('<br>');
}
