import type { ExpenseReport, ReceiptInfo, ReportSummary } from '../domain/models';

export interface IReportService {
  listReports(ownerEmail: string): Promise<ReportSummary[]>;
  loadReport(id: number): Promise<{ report: ExpenseReport; receipts: ReceiptInfo[] }>;
  saveReport(report: ExpenseReport): Promise<ExpenseReport>;
  uploadReceipt(reportId: number, file: File): Promise<ReceiptInfo>;
}

export type ReportServiceErrorKind = 'configuration' | 'permission' | 'conflict' | 'notFound' | 'invalidData' | 'network' | 'server';
export class ReportServiceError extends Error {
  public constructor(message: string, public readonly kind: ReportServiceErrorKind, public readonly status?: number) {
    super(message);
    // SPFx compiles to ES5, where extending Error loses the prototype and breaks instanceof.
    Object.setPrototypeOf(this, ReportServiceError.prototype);
    this.name = 'ReportServiceError';
  }
}
