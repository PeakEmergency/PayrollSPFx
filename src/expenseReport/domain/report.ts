import type { ExpenseReport, ValidationIssue } from './models';
import { parseDateInput } from './dates';
import { isNumericInput } from './calculations';

export type { ValidationIssue } from './models';
export const EXPENSE_LINE_COUNT: number = 13;

let identitySequence = 0;
function newIdentity(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return crypto.randomUUID();
  identitySequence += 1;
  return Date.now().toString(36) + '-' + identitySequence.toString(36) + '-' + Math.random().toString(36).slice(2);
}

export function createReport(rate: string = '0.48', user: { key: string; name: string } = { key: '', name: '' }): ExpenseReport {
  return {
    schemaVersion: 1, clientId: newIdentity(), instructorId: user.key, instructorName: user.name, fromDate: '', toDate: '',
    lines: Array.from({ length: EXPENSE_LINE_COUNT }, () => ({
      id: newIdentity(), date: '', startCity: '', endCity: '', kilometres: '', rate, description: '', gst: '', cost: ''
    }))
  };
}

export function validateReport(report: ExpenseReport): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  if (!report.instructorName.trim()) {
    issues.push({ field: 'instructorId', message: "Your name couldn't be loaded. Refresh the page." });
  }
  const from = parseDateInput(report.fromDate);
  const to = parseDateInput(report.toDate);
  if (!from) issues.push({ field: 'fromDate', message: 'Enter a valid pay period FROM date (DD/MM/YY).' });
  if (!to) issues.push({ field: 'toDate', message: 'Enter a valid pay period TO date (DD/MM/YY).' });
  if (from && to && to < from) {
    issues.push({ field: 'toDate', message: "The TO date can't be before the FROM date (" + report.fromDate.trim() + ').' });
  }
  if (report.lines.length !== EXPENSE_LINE_COUNT) issues.push({ field: 'lines', message: 'The report must retain all 13 expense rows.' });
  const identities: { [id: string]: boolean } = Object.create(null) as { [id: string]: boolean };
  report.lines.forEach((line, index) => {
    const field = 'lines.' + index + '.';
    if (!line.id.trim() || identities[line.id]) issues.push({ field: field + 'id', message: 'Each expense row must have a unique identity.' });
    identities[line.id] = true;
    if (line.date.trim() && !parseDateInput(line.date)) issues.push({ field: field + 'date', message: 'Enter a valid expense date (DD/MM/YY) or leave it empty.' });
    (['kilometres', 'rate', 'gst', 'cost'] as const).forEach(key => {
      if (!isNumericInput(line[key], key !== 'rate')) {
        issues.push({ field: field + key, message: key === 'rate' ? 'The captured mileage rate must be a number.' : 'Enter a number or leave this amount empty.' });
      }
    });
  });
  return issues;
}
