import { fillTemplate } from '../excel/fillTemplate';
import { computeLine, computeTotals } from '../timesheet/calc';
import { TEMPLATES } from '../timesheet/definitions';
import { buildFileName, buildSubject, ddmmyy } from '../timesheet/naming';
import { TimesheetData, TimesheetType } from '../timesheet/types';
import { validate } from '../timesheet/validate';

export type SubmissionStatus = 'Pending' | 'Sent' | 'Email failed';

export interface SubmissionRecord {
  title: string;
  type: TimesheetType;
  employeeId: number;
  from: string;
  to: string;
  totalAmount: number;
  lineItemsJson: string;
}

export interface SubmissionStore {
  create(r: SubmissionRecord): Promise<number>;
  attach(id: number, fileName: string, content: ArrayBuffer): Promise<void>;
  setStatus(id: number, status: SubmissionStatus, detail?: string): Promise<void>;
}

export interface TemplateSource {
  load(type: TimesheetType): Promise<ArrayBuffer>;
}

export interface MailMessage {
  to: string;
  subject: string;
  bodyText: string;
  fileName: string;
  content: ArrayBuffer;
}

export interface Mailer {
  send(m: MailMessage): Promise<void>;
}

export interface SubmitDeps {
  store: SubmissionStore;
  templates: TemplateSource;
  mailer: Mailer;
  emailTo: string;
  employeeId: number;
}

export type SubmitResult = { ok: true } | { ok: false; errors: string[] };

const isEntered = (type: TimesheetType, line: TimesheetData['lines'][number]): boolean =>
  TEMPLATES[type].columns.some((c) => c.kind !== 'computed' && line[c.key] !== null && line[c.key] !== undefined && String(line[c.key]).trim() !== '');

/** validate → list item (Pending) → fill template → attach → email → Sent. Failures after the item exists mark it 'Email failed'. */
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
    lineItemsJson: JSON.stringify(data.lines.filter((l) => isEntered(data.type, l)).map((l) => computeLine(data.type, l)))
  });

  try {
    const content = await fillTemplate(await deps.templates.load(data.type), data);
    await deps.store.attach(id, fileName, content);
    await deps.mailer.send({
      to: deps.emailTo,
      subject: buildSubject(data),
      bodyText: `${data.type} timesheet for ${data.name}, ${ddmmyy(data.from)} to ${ddmmyy(data.to)}. Total: $${total.toFixed(2)}.`,
      fileName,
      content
    });
    await deps.store.setStatus(id, 'Sent');
    return { ok: true };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    await deps.store.setStatus(id, 'Email failed', message);
    return { ok: false, errors: [`Your timesheet was saved but the email was not sent: ${message}`] };
  }
}
