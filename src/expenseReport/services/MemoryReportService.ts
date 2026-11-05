import type { ExpenseReport, ReceiptInfo, ReportStatus, ReportSummary } from '../domain/models';
import { ReportServiceError, type IReportService } from './IReportService';
import { encodeReport, decodeReport } from './reportCodec';
import { equalReceiptBytes, validateReceiptName } from './SharePointReportService';
const SUBMITTED = 'This report has been submitted and can no longer be changed. Start a new report for more expenses.';
type SavedReport = { payload: string; status: ReportStatus; email?: { to: string; subject: string; body: string }; version: number; modified: string; receipts: { info: ReceiptInfo; bytes: ArrayBuffer }[] };
export class MemoryReportService implements IReportService {
  private readonly reports = new Map<number, SavedReport>();
  private nextId = 1;
  private saved(id: number): SavedReport {
    const saved = this.reports.get(id);
    if (!saved) throw new ReportServiceError('The preview report was not found.', 'notFound', 404);
    return saved;
  }
  private model(id: number, saved: SavedReport): ExpenseReport { return { ...decodeReport(saved.payload), id, etag: `"${saved.version}"`, status: saved.status }; }
  public async listReports(): Promise<ReportSummary[]> {
    return Array.from(this.reports, ([id, saved]) => {
      const report = this.model(id, saved);
      return { id, title: `${report.instructorName} ${report.fromDate} – ${report.toDate}`, instructorName: report.instructorName,
        fromDate: report.fromDate, toDate: report.toDate, modified: saved.modified, status: saved.status };
    }).sort((a, b) => b.modified.localeCompare(a.modified) || b.id - a.id);
  }
  public async loadReport(id: number): Promise<{ report: ExpenseReport; receipts: ReceiptInfo[] }> {
    const saved = this.saved(id);
    return { report: this.model(id, saved), receipts: saved.receipts.map(receipt => ({ ...receipt.info })) };
  }
  public async saveReport(report: ExpenseReport): Promise<ExpenseReport> {
    const payload = encodeReport(report);
    if (report.id !== undefined) {
      const saved = this.saved(report.id);
      if (saved.status) throw new ReportServiceError(SUBMITTED, 'conflict');
      if (report.etag !== `"${saved.version}"`) throw new ReportServiceError('This preview report has changed. Reopen it before saving.', 'conflict', 412);
      if (decodeReport(saved.payload).clientId !== report.clientId) throw new ReportServiceError('The report identity cannot change.', 'invalidData');
      saved.payload = payload; saved.version++; saved.modified = new Date().toISOString();
      return this.model(report.id, saved);
    }
    for (const [id, saved] of this.reports) {
      if (decodeReport(saved.payload).clientId === report.clientId) {
        if (saved.status) throw new ReportServiceError(SUBMITTED, 'conflict');
        if (saved.payload !== payload) throw new ReportServiceError('This client key already identifies a different saved report. Reopen it before saving.', 'conflict');
        return this.model(id, saved);
      }
    }
    const id = this.nextId++;
    const saved = { payload, status: '' as ReportStatus, version: 1, modified: new Date().toISOString(), receipts: [] };
    this.reports.set(id, saved);
    return this.model(id, saved);
  }
  public async uploadReceipt(id: number, file: File): Promise<ReceiptInfo> {
    const saved = this.saved(id);
    validateReceiptName(file.name);
    const existing = saved.receipts.find(receipt => receipt.info.name.toLowerCase() === file.name.toLowerCase());
    if (existing) {
      if (!(await equalReceiptBytes(file, existing.bytes))) throw new ReportServiceError('A different preview receipt has this name. Rename the file before retrying.', 'conflict');
      return { ...existing.info };
    }
    const bytes = await file.arrayBuffer();
    let binary = '';
    for (const byte of new Uint8Array(bytes)) binary += String.fromCharCode(byte);
    const info = { name: file.name, url: `data:application/octet-stream;base64,${btoa(binary)}` };
    saved.receipts.push({ info, bytes });
    return { ...info };
  }
  public async setStatus(id: number, status: ReportStatus, etag?: string): Promise<void> {
    const saved = this.saved(id);
    if (etag !== undefined && etag !== `"${saved.version}"`) throw new ReportServiceError('This preview report has changed. Reopen it before saving.', 'conflict', 412);
    saved.status = status; saved.version++;
  }
  public async deleteAttachment(id: number, fileName: string): Promise<void> {
    const saved = this.saved(id);
    saved.receipts = saved.receipts.filter(receipt => receipt.info.name !== fileName);
  }
  public async attachFile(id: number, fileName: string, content: ArrayBuffer): Promise<void> {
    await this.deleteAttachment(id, fileName);
    let binary = '';
    for (const byte of new Uint8Array(content)) binary += String.fromCharCode(byte);
    this.saved(id).receipts.push({ info: { name: fileName, url: `data:application/octet-stream;base64,${btoa(binary)}` }, bytes: content });
  }
  public async queueEmail(id: number, email: { to: string; subject: string; body: string }): Promise<void> {
    const saved = this.saved(id);
    saved.email = email; saved.status = 'Queued'; saved.version++;
  }
}
