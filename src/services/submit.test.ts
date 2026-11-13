import * as fs from 'fs';
import * as path from 'path';
import { ExpenseSource, MAX_EXPENSE_BYTES, SubmissionRecord, SubmissionStatus, SubmitDeps, submitTimesheet } from './submit';
import { TimesheetData, TimesheetType } from '../timesheet/types';

const medicFixture = fs.readFileSync(path.resolve(__dirname, '../../Templates/xlsx/Medic Timesheet Template.xlsx'));

const data: TimesheetData = {
  type: 'Medic',
  name: 'NIKODYM, CAROLYN',
  from: '2026-10-01',
  to: '2026-10-15',
  lines: [{ date: '2026-10-15', client: 'acme', days: 1, dayRate: 450, travel: 100, perDiem: 47 }, {}]
};

interface Fakes {
  deps: SubmitDeps;
  calls: string[];
  created: SubmissionRecord[];
  statuses: [number, SubmissionStatus, string | undefined][];
  attached: string[];
  body?: string;
}

function fakes(opts: { attachError?: Error; loadError?: Error } = {}): Fakes {
  const f: Fakes = { calls: [], created: [], statuses: [], attached: [], deps: undefined! };
  f.deps = {
    emailTo: 'accounting@peakemergency.com',
    employeeId: 7,
    senderEmail: 'carolyn@peakemergency.com',
    store: {
      create: async (r) => { f.calls.push('create'); f.created.push(r); return 42; },
      attach: async (_id, name) => { f.calls.push('attach'); if (opts.attachError) throw opts.attachError; f.attached.push(name); },
      setStatus: async (id, status, detail) => { f.calls.push(`setStatus(${status})`); f.statuses.push([id, status, detail]); },
      deleteAttachment: async (_id, name) => { f.calls.push(`delete(${name})`); f.attached = f.attached.filter((n) => n !== name); },
      setEmailBody: async (_id, body) => { f.calls.push('setEmailBody'); f.body = body; }
    },
    templates: {
      load: async (_type: TimesheetType) => {
        f.calls.push('load');
        if (opts.loadError) throw opts.loadError;
        return new Uint8Array(medicFixture).buffer;
      }
    }
  };
  return f;
}

it('happy path queues the item for the email flow', async () => {
  const f = fakes();
  expect(await submitTimesheet(data, f.deps)).toEqual({ ok: true });
  // Queued only after the file is attached: the flow triggers on Status = Queued.
  expect(f.calls).toEqual(['create', 'load', 'attach', 'setStatus(Queued)']);
  expect(f.created[0]).toMatchObject({
    title: 'Medic Timesheet - NIKODYM, CAROLYN - 2026-10-01_2026-10-15.xlsx',
    totalAmount: 597,
    employeeId: 7,
    type: 'Medic',
    from: '2026-10-01',
    to: '2026-10-15',
    emailTo: 'accounting@peakemergency.com',
    emailSubject: 'Medic Timesheet from NIKODYM, CAROLYN – 01/10/26 to 15/10/26'
  });
  expect(JSON.parse(f.created[0].lineItemsJson)).toHaveLength(1);
});

it('email body names the real sender with their email', () => {
  const f = fakes();
  return submitTimesheet(data, f.deps).then(() => {
    const body = f.created[0].emailBody;
    expect(body).toContain('NIKODYM, CAROLYN');
    expect(body).toContain('carolyn@peakemergency.com');
    expect(body).toContain('01/10/26 to 15/10/26');
    expect(body).toContain('$597.00');
  });
});

it('email body escapes HTML in names', async () => {
  const f = fakes();
  await submitTimesheet({ ...data, name: 'SMITH & SONS, <B>' }, f.deps);
  expect(f.created[0].emailBody).toContain('SMITH &amp; SONS, &lt;B&gt;');
});

it('invalid data never touches SharePoint', async () => {
  const f = fakes();
  const result = await submitTimesheet({ ...data, name: '' }, f.deps);
  expect(result.ok).toBe(false);
  expect(f.calls).toEqual([]);
});

it('attach failure marks item Email failed and is never queued', async () => {
  const f = fakes({ attachError: new Error('Attachment too large') });
  const result = await submitTimesheet(data, f.deps);
  expect(f.statuses).toEqual([[42, 'Email failed', 'Attachment too large']]);
  expect(result.ok).toBe(false);
  expect(!result.ok && result.errors[0]).toContain('Attachment too large');
});

it('template load failure marks item Email failed', async () => {
  const f = fakes({ loadError: new Error('Template not found') });
  const result = await submitTimesheet(data, f.deps);
  expect(f.statuses).toEqual([[42, 'Email failed', 'Template not found']]);
  expect(f.calls).not.toContain('attach');
  expect(result.ok).toBe(false);
});

it('status update failure in the error path keeps the original error', async () => {
  const f = fakes({ loadError: new Error('Template not found') });
  f.deps.store.setStatus = async () => { throw new Error('List throttled'); };
  const result = await submitTimesheet(data, f.deps);
  expect(!result.ok && result.errors[0]).toContain('Template not found');
});

it('sends a copy to the address the submitter asked for', async () => {
  const f = fakes();
  expect(await submitTimesheet({ ...data, copyTo: 'carolyn@peakemergency.com' }, f.deps)).toEqual({ ok: true });
  expect(f.created[0].emailTo).toBe('accounting@peakemergency.com;carolyn@peakemergency.com');
});

it('refuses a bad copy address before creating anything', async () => {
  const f = fakes();
  const result = await submitTimesheet({ ...data, copyTo: 'nope' }, f.deps);
  expect(result.ok).toBe(false);
  expect(f.calls).toEqual([]);
});

describe('expense reports for the same person and pay period', () => {
  const file = (name: string, size = 10): { name: string; load: () => Promise<ArrayBuffer> } => ({ name, load: async () => new ArrayBuffer(size) });
  const source = (reports: Awaited<ReturnType<ExpenseSource['find']>>, seen: string[] = []): ExpenseSource => ({
    find: async (name, from, to) => { seen.push(`${name}|${from}|${to}`); return reports; }
  });
  const timesheetFile = 'Medic Timesheet - NIKODYM, CAROLYN - 2026-10-01_2026-10-15.xlsx';

  it('none found: sends the timesheet alone with no warning', async () => {
    const f = fakes();
    f.deps.expenses = source([]);
    expect(await submitTimesheet(data, f.deps)).toEqual({ ok: true });
    expect(f.attached).toEqual([timesheetFile]);
    expect(f.created[0].emailBody).not.toContain('Expense reports');
  });

  it('copies every file of every match before queuing, and counts the reports in the body', async () => {
    const f = fakes();
    const seen: string[] = [];
    f.deps.expenses = source([
      { id: 5, files: [file('Expense Report - NIKODYM, CAROLYN - 01-10-26_15-10-26 - 5.xlsx'), file('receipt.jpg')] },
      { id: 9, files: [file('receipt.jpg')] }
    ], seen);
    expect(await submitTimesheet(data, f.deps)).toEqual({ ok: true });
    expect(seen).toEqual(['NIKODYM, CAROLYN|2026-10-01|2026-10-15']);
    // The second receipt.jpg gets its report id so the names don't clash.
    expect(f.attached).toEqual([timesheetFile, 'Expense Report - NIKODYM, CAROLYN - 01-10-26_15-10-26 - 5.xlsx', 'receipt.jpg', '9 - receipt.jpg']);
    expect(f.calls[f.calls.length - 1]).toBe('setStatus(Queued)');
    expect(f.created[0].emailBody).toContain('Expense reports attached: 2');
  });

  it('lookup failure still sends the timesheet, with a warning', async () => {
    const f = fakes();
    f.deps.expenses = { find: async () => { throw new Error('Loading expense reports failed (403)'); } };
    const result = await submitTimesheet(data, f.deps);
    expect(result).toEqual({ ok: true, warning: expect.stringContaining('Loading expense reports failed (403)') });
    expect(f.attached).toEqual([timesheetFile]);
    expect(f.statuses.map((s) => s[1])).toEqual(['Queued']);
  });

  it('files over the size cap are left off, with a warning', async () => {
    const f = fakes();
    f.deps.expenses = source([{ id: 5, files: [file('big.pdf', MAX_EXPENSE_BYTES + 1)] }]);
    const result = await submitTimesheet(data, f.deps);
    expect(result).toEqual({ ok: true, warning: expect.stringContaining('too large') });
    expect(f.attached).toEqual([timesheetFile]);
    expect(f.created[0].emailBody).not.toContain('Expense reports');
  });

  it('a copy failing partway removes the copies already added and fixes the body', async () => {
    const f = fakes();
    f.deps.expenses = source([{ id: 5, files: [file('a.jpg'), file('b.jpg')] }]);
    const attach = f.deps.store.attach;
    f.deps.store.attach = async (id, name, content) => { if (name === 'b.jpg') throw new Error('Attaching the file failed (500)'); return attach(id, name, content); };
    const result = await submitTimesheet(data, f.deps);
    expect(result).toEqual({ ok: true, warning: expect.stringContaining('Attaching the file failed (500)') });
    expect(f.attached).toEqual([timesheetFile]);
    expect(f.calls).toContain('delete(a.jpg)');
    expect(f.body).not.toContain('Expense reports');
    expect(f.statuses.map((s) => s[1])).toEqual(['Queued']);
  });
});
