import * as fs from 'fs';
import * as path from 'path';
import { SubmissionRecord, SubmissionStatus, SubmitDeps, submitTimesheet } from './submit';
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
}

function fakes(opts: { attachError?: Error; loadError?: Error } = {}): Fakes {
  const f: Fakes = { calls: [], created: [], statuses: [], deps: undefined! };
  f.deps = {
    emailTo: 'accounting@peakemergency.com',
    employeeId: 7,
    senderEmail: 'carolyn@peakemergency.com',
    store: {
      create: async (r) => { f.calls.push('create'); f.created.push(r); return 42; },
      attach: async () => { f.calls.push('attach'); if (opts.attachError) throw opts.attachError; },
      setStatus: async (id, status, detail) => { f.calls.push(`setStatus(${status})`); f.statuses.push([id, status, detail]); }
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
