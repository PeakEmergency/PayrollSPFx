import { fillTemplate } from '../excel/fillTemplate';
import { computeLine, computeTotals } from '../timesheet/calc';
import { TEMPLATES } from '../timesheet/definitions';
import { buildFileName, buildSubject, ddmmyy } from '../timesheet/naming';
import { TimesheetData, TimesheetType } from '../timesheet/types';
import { validate } from '../timesheet/validate';
import { recipients } from './copyTo';

/** Draft: saved to finish later; the email flow only acts on Queued. */
export type SubmissionStatus = 'Draft' | 'Pending' | 'Queued' | 'Sent' | 'Email failed';

/** List item. emailTo/emailSubject/emailBody are read by the Power Automate flow that sends the email. */
export interface SubmissionRecord {
  title: string;
  type: TimesheetType;
  employeeId: number;
  /** null only on a draft. */
  // eslint-disable-next-line @rushstack/no-new-null
  from: string | null;
  // eslint-disable-next-line @rushstack/no-new-null
  to: string | null;
  totalAmount: number;
  lineItemsJson: string;
  emailTo: string;
  emailSubject: string;
  /** HTML. */
  emailBody: string;
}

/** A saved, unsubmitted timesheet. */
export interface Draft {
  id: number;
  modified: string;
  data: TimesheetData;
}

export interface SubmissionStore {
  create(r: SubmissionRecord, status: SubmissionStatus): Promise<number>;
  /** Overwrites an existing item (a draft) with r and status. */
  update(id: number, r: SubmissionRecord, status: SubmissionStatus): Promise<void>;
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
  /** The draft being finished: submitting turns that item into the submission. */
  draftId?: number;
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
  const record: SubmissionRecord = {
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
  };
  let id: number;
  if (deps.draftId === undefined) id = await deps.store.create(record, 'Pending');
  else {
    id = deps.draftId;
    await deps.store.update(id, record, 'Pending');
  }

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

/**
 * Saves the timesheet as it stands, valid or not, to finish later. The whole TimesheetData goes in
 * LineItems so it reopens exactly; submitting replaces it with the computed lines. Returns the item id.
 */
export async function saveDraft(data: TimesheetData, deps: Pick<SubmitDeps, 'store' | 'employeeId' | 'draftId'>): Promise<number> {
  const record: SubmissionRecord = {
    title: `Draft - ${data.type} Timesheet - ${data.name || 'no name'}`,
    type: data.type,
    employeeId: deps.employeeId,
    from: data.from,
    to: data.to,
    totalAmount: computeTotals(data.type, data.lines)[TEMPLATES[data.type].grandTotalKey],
    lineItemsJson: JSON.stringify(data),
    emailTo: '',
    emailSubject: '',
    emailBody: ''
  };
  if (deps.draftId === undefined) return deps.store.create(record, 'Draft');
  await deps.store.update(deps.draftId, record, 'Draft');
  return deps.draftId;
}
