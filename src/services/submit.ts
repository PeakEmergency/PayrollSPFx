import { fillTemplate } from '../excel/fillTemplate';
import { computeLine, computeTotals } from '../timesheet/calc';
import { TEMPLATES } from '../timesheet/definitions';
import { buildFileName, buildSubject, ddmmyy } from '../timesheet/naming';
import { TimesheetData, TimesheetType } from '../timesheet/types';
import { validate } from '../timesheet/validate';
import { recipients } from './copyTo';

export type SubmissionStatus = 'Pending' | 'Queued' | 'Sent' | 'Email failed';

/** List item. emailTo/emailSubject/emailBody are read by the Power Automate flow that sends the email. */
export interface SubmissionRecord {
  title: string;
  type: TimesheetType;
  employeeId: number;
  from: string;
  to: string;
  totalAmount: number;
  lineItemsJson: string;
  emailTo: string;
  emailSubject: string;
  /** HTML. */
  emailBody: string;
}

export interface SubmissionStore {
  create(r: SubmissionRecord): Promise<number>;
  attach(id: number, fileName: string, content: ArrayBuffer): Promise<void>;
  setStatus(id: number, status: SubmissionStatus, detail?: string): Promise<void>;
  deleteAttachment(id: number, fileName: string): Promise<void>;
  setEmailBody(id: number, body: string): Promise<void>;
}

/** A submitted expense report's attachments: the filled Excel form and the receipts. */
export interface ExpenseFile {
  name: string;
  load(): Promise<ArrayBuffer>;
}

export interface ExpenseSource {
  /** Every submitted (Queued, Sent or Email failed) expense report for this person and pay period. */
  find(name: string, from: string, to: string): Promise<{ id: number; files: ExpenseFile[] }[]>;
}

/** Over this the timesheet email could fail outright, so it goes without the expense reports. */
export const MAX_EXPENSE_BYTES = 20 * 1024 * 1024;

export interface TemplateSource {
  load(type: TimesheetType): Promise<ArrayBuffer>;
}

export interface SubmitDeps {
  store: SubmissionStore;
  templates: TemplateSource;
  emailTo: string;
  employeeId: number;
  /** Signed-in user's email, named in the body: the flow sends from the shared schedule mailbox. */
  senderEmail: string;
  /** Expense reports for the same person and pay period are copied onto the timesheet item. */
  expenses?: ExpenseSource;
}

/** warning: the timesheet was sent, but without its expense reports. */
export type SubmitResult = { ok: true; warning?: string } | { ok: false; errors: string[] };

const isEntered = (type: TimesheetType, line: TimesheetData['lines'][number]): boolean =>
  TEMPLATES[type].columns.some((c) => c.kind !== 'computed' && line[c.key] !== null && line[c.key] !== undefined && String(line[c.key]).trim() !== '');

const escapeHtml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function buildBody(data: TimesheetData, senderEmail: string, total: number, expenseReports: number): string {
  return [
    `${data.type} timesheet submitted by <b>${escapeHtml(data.name)}</b> (${escapeHtml(senderEmail)}).`,
    `Pay period: ${ddmmyy(data.from)} to ${ddmmyy(data.to)}`,
    `Total: $${total.toFixed(2)}`,
    ...(expenseReports ? [`Expense reports attached: ${expenseReports}`] : [])
  ].join('<br>');
}

const message = (err: unknown): string => (err instanceof Error ? err.message : String(err));
const notIncluded = (why: string): string => `Your timesheet was sent, but your expense report could not be included: ${why}`;

type LoadedExpense = { id: number; files: { name: string; bytes: ArrayBuffer }[] };

/** Reads the files up front, so the body can count the reports and a read failure never touches the list. */
async function loadExpenses(source: ExpenseSource, data: TimesheetData): Promise<{ reports: LoadedExpense[]; warning?: string }> {
  try {
    const reports: LoadedExpense[] = [];
    for (const r of await source.find(data.name, data.from!, data.to!)) {
      reports.push({ id: r.id, files: await Promise.all(r.files.map(async (f) => ({ name: f.name, bytes: await f.load() }))) });
    }
    const size = reports.reduce((sum, r) => r.files.reduce((s, f) => s + f.bytes.byteLength, sum), 0);
    if (size > MAX_EXPENSE_BYTES) {
      return { reports: [], warning: notIncluded(`its files are too large to email (${Math.ceil(size / 1048576)} MB). Accounts get it in its own email.`) };
    }
    return { reports };
  } catch (err) {
    return { reports: [], warning: notIncluded(message(err)) };
  }
}

/** All or nothing: on a failure the copies already added are removed, so accounts never get half a report. */
async function copyExpenses(store: SubmissionStore, id: number, timesheetFile: string, reports: LoadedExpense[]): Promise<string | undefined> {
  const used = new Set([timesheetFile.toLowerCase()]);
  const added: string[] = [];
  try {
    for (const r of reports) {
      for (const f of r.files) {
        // An item can't hold two attachments with one name (two reports may share a receipt name).
        const name = used.has(f.name.toLowerCase()) ? `${r.id} - ${f.name}` : f.name;
        used.add(name.toLowerCase());
        await store.attach(id, name, f.bytes);
        added.push(name);
      }
    }
    return undefined;
  } catch (err) {
    for (const name of added) await store.deleteAttachment(id, name).catch(() => undefined);
    return notIncluded(message(err));
  }
}

/**
 * validate → read matching expense reports → list item (Pending) → fill template → attach → copy expense files → Queued.
 * A Power Automate flow emails items once they are Queued and sets Sent / Email failed.
 * Failures here mark the item 'Email failed'.
 */
export async function submitTimesheet(data: TimesheetData, deps: SubmitDeps): Promise<SubmitResult> {
  const errors = validate(data);
  if (errors.length) return { ok: false, errors };

  const def = TEMPLATES[data.type];
  const total = computeTotals(data.type, data.lines)[def.grandTotalKey];
  const fileName = buildFileName(data);
  const expenses = deps.expenses ? await loadExpenses(deps.expenses, data) : { reports: [] as LoadedExpense[], warning: undefined };
  const id = await deps.store.create({
    title: fileName,
    type: data.type,
    employeeId: deps.employeeId,
    from: data.from!,
    to: data.to!,
    totalAmount: total,
    lineItemsJson: JSON.stringify(data.lines.filter((l) => isEntered(data.type, l)).map((l) => computeLine(data.type, l))),
    emailTo: recipients(deps.emailTo, data.copyTo),
    emailSubject: buildSubject(data),
    emailBody: buildBody(data, deps.senderEmail, total, expenses.reports.length)
  });

  try {
    const content = await fillTemplate(await deps.templates.load(data.type), data);
    await deps.store.attach(id, fileName, content);
    let warning = expenses.warning;
    if (expenses.reports.length) {
      warning = await copyExpenses(deps.store, id, fileName, expenses.reports);
      // Best effort: a wrong count in the body must not stop the timesheet.
      if (warning) await deps.store.setEmailBody(id, buildBody(data, deps.senderEmail, total, 0)).catch(() => undefined);
    }
    await deps.store.setStatus(id, 'Queued');
    return warning ? { ok: true, warning } : { ok: true };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    try {
      await deps.store.setStatus(id, 'Email failed', message);
    } catch {
      // Keep the original error for the user; the item stays Pending.
    }
    return { ok: false, errors: [`Your timesheet was saved but could not be queued for email: ${message}`] };
  }
}
