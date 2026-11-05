import type { ExpenseReport } from './domain/models';
import { expenseFileName } from './domain/submission';
import { MemoryReportService } from './services/MemoryReportService';
import { submitReport, type SubmitReportDeps } from './services/submitReport';

function draft(): ExpenseReport {
  return { schemaVersion: 1, clientId: 'client-one', instructorId: 'instructor-01',
    instructorName: 'Test Instructor', fromDate: '01/11/26', toDate: '15/11/26',
    lines: Array.from({ length: 13 }, (_, i) => ({ id: `line-${i}`, date: '', startCity: '', endCity: '',
      kilometres: '', rate: '0.48', description: '', gst: '', cost: '' })) };
}
function setup(overrides: Partial<SubmitReportDeps> = {}): { service: MemoryReportService; deps: SubmitReportDeps } {
  const service = new MemoryReportService();
  return { service, deps: { service, loadTemplate: async () => new ArrayBuffer(1), emailTo: 'accounts@example.com',
    senderEmail: 'me@example.com', receiptCount: 0, fill: async () => new Uint8Array([1]).buffer, ...overrides } };
}

describe('submitReport', () => {
  it('queues a draft', async () => {
    const { service, deps } = setup();
    const saved = await service.saveReport(draft());
    const setStatus = jest.spyOn(service, 'setStatus');
    await submitReport(saved, deps);
    const { report, receipts } = await service.loadReport(saved.id!);
    expect(report.status).toBe('Queued');
    expect(receipts.map(r => r.name)).toContain(expenseFileName(saved, saved.id!));
    expect(setStatus.mock.calls[0]).toEqual([saved.id, 'Pending', saved.etag]);
  });

  it('finishes a Pending report without its ETag', async () => {
    const { service, deps } = setup();
    const saved = await service.saveReport(draft());
    await service.setStatus(saved.id!, 'Pending');
    const setStatus = jest.spyOn(service, 'setStatus');
    await submitReport({ ...saved, status: 'Pending', etag: undefined }, deps);
    expect((await service.loadReport(saved.id!)).report.status).toBe('Queued');
    expect(setStatus).not.toHaveBeenCalledWith(saved.id, 'Pending', expect.anything());
    expect(setStatus).not.toHaveBeenCalledWith(saved.id, 'Pending');
  });

  it('returns to draft when the email columns fail', async () => {
    const { service, deps } = setup();
    const saved = await service.saveReport(draft());
    jest.spyOn(service, 'queueEmail').mockRejectedValue(new Error('columns missing'));
    await expect(submitReport(saved, deps)).rejects.toThrow(/^Submitting failed; the report is still a draft\./);
    const { report, receipts } = await service.loadReport(saved.id!);
    expect(report.status).toBe('');
    expect(receipts).toHaveLength(0);
  });

  it('returns to draft when the template fails to load', async () => {
    const { service, deps } = setup({ loadTemplate: async () => { throw new Error('no template'); } });
    const saved = await service.saveReport(draft());
    const del = jest.spyOn(service, 'deleteAttachment');
    await expect(submitReport(saved, deps)).rejects.toThrow('no template');
    expect((await service.loadReport(saved.id!)).report.status).toBe('');
    expect(del).not.toHaveBeenCalled();
  });

  it('refuses a submitted report', async () => {
    const { service, deps } = setup();
    const saved = await service.saveReport(draft());
    const setStatus = jest.spyOn(service, 'setStatus');
    await expect(submitReport({ ...saved, status: 'Sent' }, deps)).rejects.toMatchObject({ kind: 'conflict' });
    expect(setStatus).not.toHaveBeenCalled();
  });
});
