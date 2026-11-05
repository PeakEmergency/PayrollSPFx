import { fillExpenseReport } from '../../excel/fillExpenseReport';
import type { ExpenseReport } from '../domain/models';
import { expenseBody, expenseFileName, expenseSubject } from '../domain/submission';
import { ReportServiceError, type IReportService } from './IReportService';

export interface SubmitReportDeps {
  service: IReportService;
  loadTemplate(): Promise<ArrayBuffer>;
  emailTo: string;
  senderEmail: string;
  receiptCount: number;
  fill?: typeof fillExpenseReport; // tests replace it; defaults to fillExpenseReport
}

/** report must be saved (id and etag set) with status '' or 'Pending'. */
export async function submitReport(report: ExpenseReport, deps: SubmitReportDeps): Promise<void> {
  const { service } = deps;
  const id = report.id as number;
  if (report.status && report.status !== 'Pending') throw new ReportServiceError('This report has already been submitted.', 'conflict');
  if (!report.status) await service.setStatus(id, 'Pending', report.etag);
  const fileName = expenseFileName(report, id);
  try {
    const bytes = await (deps.fill || fillExpenseReport)(await deps.loadTemplate(), report);
    await service.attachFile(id, fileName, bytes);
    await service.queueEmail(id, { to: deps.emailTo, subject: expenseSubject(report), body: expenseBody(report, deps.senderEmail, deps.receiptCount) });
  } catch (error) {
    // Best effort: the original error matters more than a failed cleanup. Always delete: an upload may have committed even if its response was lost.
    await service.deleteAttachment(id, fileName).catch(() => undefined);
    await service.setStatus(id, '').catch(() => undefined);
    const message = `Submitting failed; the report is still a draft. ${error instanceof Error ? error.message : String(error)}`;
    throw error instanceof ReportServiceError ? new ReportServiceError(message, error.kind, error.status) : new Error(message);
  }
}
