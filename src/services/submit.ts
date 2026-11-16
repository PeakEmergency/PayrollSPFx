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
}

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
}

export type SubmitResult = { ok: true } | { ok: false; errors: string[] };

const isEntered = (type: TimesheetType, line: TimesheetData['lines'][number]): boolean =>
  TEMPLATES[type].columns.some((c) => c.kind !== 'computed' && line[c.key] !== null && line[c.key] !== undefined && String(line[c.key]).trim() !== '');

const escapeHtml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function buildBody(data: TimesheetData, senderEmail: string, total: number): string {
  return [
    `${data.type} timesheet submitted by <b>${escapeHtml(data.name)}</b> (${escapeHtml(senderEmail)}).`,
    `Pay period: ${ddmmyy(data.from)} to ${ddmmyy(data.to)}`,
    `Total: $${total.toFixed(2)}`
  ].join('<br>');
}

/**
 * validate → list item (Pending) → fill template → attach → Queued.
 * A Power Automate flow emails items once they are Queued and sets Sent / Email failed.
 * Failures here mark the item 'Email failed'.
 */
export async function submitTimesheet(data: TimesheetData, deps: SubmitDeps): Promise<SubmitResult> {
  const errors = validate(data);
  if (errors.length) return { ok: false, errors };

  const def = TEMPLATES[data.type];
  const total = computeTotals(data.type, data.lines)[def.grandTotalKey];
  const fileName = buildFileName(data);
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
    emailBody: buildBody(data, deps.senderEmail, total)
  });

  try {
    const content = await fillTemplate(await deps.templates.load(data.type), data);
    await deps.store.attach(id, fileName, content);
    await deps.store.setStatus(id, 'Queued');
    return { ok: true };
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
