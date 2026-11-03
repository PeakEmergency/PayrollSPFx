import type { ExpenseLine, ExpenseReport } from '../domain/models';
import { formatDateInput, parseDateInput } from '../domain/dates';
import { ReportServiceError } from './IReportService';

const reportKeys = ['schemaVersion', 'clientId', 'instructorId', 'instructorName', 'fromDate', 'toDate', 'lines'];
const lineKeys = ['id', 'date', 'startCity', 'endCity', 'kilometres', 'rate', 'description', 'gst', 'cost'];
const decimalPattern = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/;
function invalid(detail: string): never {
  throw new ReportServiceError(`This report contains invalid or unsupported saved data: ${detail}.`, 'invalidData');
}
function object(value: unknown, field: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) invalid(field);
  return value as Record<string, unknown>;
}
function exactKeys(value: Record<string, unknown>, keys: string[], field: string): void {
  if (Object.keys(value).length !== keys.length || keys.some(key => !Object.prototype.hasOwnProperty.call(value, key))) invalid(`${field} fields`);
}
function text(value: unknown, field: string, required = false, max?: number): string {
  if (typeof value !== 'string' || (required && !value.trim()) || (max !== undefined && value.length > max)) invalid(field);
  return value;
}
function numberText(value: unknown, field: string, required = false): string {
  const raw = text(value, field, required);
  if (raw.trim() && !decimalPattern.test(raw.trim())) invalid(field);
  return raw;
}
function date(value: unknown, field: string, stored: boolean, required: boolean): string {
  const raw = text(value, field, required);
  if (!raw || (!stored && !raw.trim())) return '';
  if (stored) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(raw) || !formatDateInput(raw) || parseDateInput(formatDateInput(raw)) !== raw) invalid(field);
    return raw;
  }
  const iso = parseDateInput(raw);
  if (!iso) invalid(field);
  return iso;
}
/** Canonical payload excludes SharePoint identity/version metadata and calculated values. */
function normalize(value: unknown, stored: boolean): ExpenseReport {
  const input = object(value, 'report');
  if (stored) exactKeys(input, reportKeys, 'report');
  if (input.schemaVersion !== 1) invalid('schema version');
  const fromDate = date(input.fromDate, 'From date', stored, true);
  const toDate = date(input.toDate, 'To date', stored, true);
  if (toDate < fromDate) invalid('pay period range');
  if (!Array.isArray(input.lines) || input.lines.length !== 13) invalid('exactly 13 row slots are required');
  const ids = new Set<string>();
  const lines: ExpenseLine[] = input.lines.map((entry: unknown, index: number) => {
    const row = object(entry, `row ${index + 1}`);
    exactKeys(row, lineKeys, `row ${index + 1}`);
    const id = text(row.id, 'row identity', true);
    if (ids.has(id)) invalid('duplicate row identity');
    ids.add(id);
    return { id, date: date(row.date, `row ${index + 1} date`, stored, false),
      startCity: text(row.startCity, 'start city'), endCity: text(row.endCity, 'end city'),
      kilometres: numberText(row.kilometres, 'kilometres'), rate: numberText(row.rate, 'rate', true),
      description: text(row.description, 'description'), gst: numberText(row.gst, 'GST'), cost: numberText(row.cost, 'cost') };
  });
  const instructorName = text(input.instructorName, 'instructor name', true, 255);
  if (instructorName.trim().toUpperCase() === 'PICK FROM DROPDOWN') invalid('instructor selection');
  return { schemaVersion: 1, clientId: text(input.clientId, 'client identity', true, 255),
    instructorId: text(input.instructorId, 'instructor key', true, 255), instructorName, fromDate, toDate, lines };
}
export function encodeReport(report: ExpenseReport): string {
  const payload = JSON.stringify(normalize(report, false));
  // SharePoint plain multiline text columns support at most 63,999 characters.
  if (payload.length > 63999) invalid('report exceeds the SharePoint 63,999-character payload capacity');
  return payload;
}
export function decodeReport(payload: string): ExpenseReport {
  if (typeof payload !== 'string' || payload.length > 63999) invalid('payload');
  let parsed: unknown;
  try { parsed = JSON.parse(payload); } catch { invalid('JSON payload'); }
  const report = normalize(parsed, true);
  return { ...report, fromDate: formatDateInput(report.fromDate), toDate: formatDateInput(report.toDate),
    lines: report.lines.map(line => ({ ...line, date: line.date ? formatDateInput(line.date) : '' })) };
}
