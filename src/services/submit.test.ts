import * as fs from 'fs';
import * as path from 'path';
import { MailMessage, SubmissionRecord, SubmissionStatus, SubmitDeps, submitTimesheet } from './submit';
import { TimesheetData, TimesheetType } from '../timesheet/types';

const medicFixture = fs.readFileSync(path.resolve(__dirname, '../../Templates/xlsx/Medic Timesheet Template.xlsx'));

const data: TimesheetData = {
  type: 'Medic',
  name: 'NIKODYM, CAROLYN',
  from: '2026-10-01',
  to: '2026-10-15',
  lines: [{ date: '2026-10-15', client: 'acme', start: '19:00', end: '07:00', dayRate: 450, travel: 100, perDiem: 47 }, {}]
};

interface Fakes {
  deps: SubmitDeps;
  calls: string[];
  created: SubmissionRecord[];
  sent: MailMessage[];
  statuses: [number, SubmissionStatus, string | undefined][];
}

function fakes(opts: { sendError?: Error; loadError?: Error } = {}): Fakes {
  const f: Fakes = { calls: [], created: [], sent: [], statuses: [], deps: undefined! };
  f.deps = {
    emailTo: 'account@peakemergency.com',
    employeeId: 7,
    store: {
      create: async (r) => { f.calls.push('create'); f.created.push(r); return 42; },
      attach: async () => { f.calls.push('attach'); },
      setStatus: async (id, status, detail) => { f.calls.push(`setStatus(${status})`); f.statuses.push([id, status, detail]); }
    },
    templates: {
      load: async (_type: TimesheetType) => {
        f.calls.push('load');
        if (opts.loadError) throw opts.loadError;
        return new Uint8Array(medicFixture).buffer;
      }
    },
    mailer: {
      send: async (m) => { f.calls.push('send'); f.sent.push(m); if (opts.sendError) throw opts.sendError; }
    }
  };
  return f;
}

it('happy path', async () => {
  const f = fakes();
  expect(await submitTimesheet(data, f.deps)).toEqual({ ok: true });
  expect(f.calls).toEqual(['create', 'load', 'attach', 'send', 'setStatus(Sent)']);
  expect(f.created[0]).toMatchObject({ totalAmount: 597, employeeId: 7, type: 'Medic', from: '2026-10-01', to: '2026-10-15' });
  expect(JSON.parse(f.created[0].lineItemsJson)).toHaveLength(1);
  expect(f.sent[0].to).toBe('account@peakemergency.com');
  expect(f.sent[0].fileName).toBe('Medic Timesheet - NIKODYM, CAROLYN - 2026-10-01_2026-10-15.xlsx');
  expect(f.sent[0].bodyText).toBe('Medic timesheet for NIKODYM, CAROLYN, 01/10/26 to 15/10/26. Total: $597.00.');
});

it('invalid data never touches SharePoint', async () => {
  const f = fakes();
  const result = await submitTimesheet({ ...data, name: '' }, f.deps);
  expect(result.ok).toBe(false);
  expect(f.calls).toEqual([]);
});

it('send failure marks item Email failed and keeps file', async () => {
  const f = fakes({ sendError: new Error('Mailbox unavailable') });
  const result = await submitTimesheet(data, f.deps);
  expect(f.calls).toContain('attach');
  expect(f.statuses).toEqual([[42, 'Email failed', 'Mailbox unavailable']]);
  expect(result.ok).toBe(false);
  expect(!result.ok && result.errors[0]).toContain('Mailbox unavailable');
});

it('template load failure marks item Email failed', async () => {
  const f = fakes({ loadError: new Error('Template not found') });
  const result = await submitTimesheet(data, f.deps);
  expect(f.statuses).toEqual([[42, 'Email failed', 'Template not found']]);
  expect(f.calls).not.toContain('send');
  expect(result.ok).toBe(false);
});
