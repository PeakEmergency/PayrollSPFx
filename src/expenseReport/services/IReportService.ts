import type { ExpenseReport, ReceiptInfo, ReportStatus, ReportSummary } from '../domain/models';

export interface IReportService {
  listReports(ownerEmail: string): Promise<ReportSummary[]>;
  loadReport(id: number): Promise<{ report: ExpenseReport; receipts: ReceiptInfo[] }>;
  saveReport(report: ExpenseReport): Promise<ExpenseReport>;
  uploadReceipt(reportId: number, file: File): Promise<ReceiptInfo>;
  setStatus(reportId: number, status: ReportStatus, etag?: string): Promise<void>;
  /** Replaces an existing attachment of the same name. */
  attachFile(reportId: number, fileName: string, content: ArrayBuffer): Promise<void>;
  /** A missing file is not an error. */
  deleteAttachment(reportId: number, fileName: string): Promise<void>;
  queueEmail(reportId: number, email: { to: string; subject: string; body: string }): Promise<void>;
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
