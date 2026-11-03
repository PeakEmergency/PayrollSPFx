import type { ExpenseReport, Instructor, ReceiptInfo, ReportSummary } from '../domain/models';
import { ReportServiceError, type IReportService } from './IReportService';
import { encodeReport, decodeReport } from './reportCodec';
import { equalReceiptBytes, validateReceiptName } from './SharePointReportService';
type SavedReport = { payload: string; version: number; modified: string; receipts: { info: ReceiptInfo; bytes: ArrayBuffer }[] };
export class MemoryReportService implements IReportService {
  private readonly reports = new Map<number, SavedReport>();
  private readonly instructors: Instructor[];
  private nextId = 1;
  public constructor(instructors: Instructor[]) { this.instructors = instructors.map(instructor => ({ ...instructor })); }
  public async getInstructors(): Promise<Instructor[]> { return this.instructors.map(instructor => ({ ...instructor })); }
  private saved(id: number): SavedReport {
    const saved = this.reports.get(id);
    if (!saved) throw new ReportServiceError('The preview report was not found.', 'notFound', 404);
    return saved;
  }
  private model(id: number, saved: SavedReport): ExpenseReport { return { ...decodeReport(saved.payload), id, etag: `"${saved.version}"` }; }
  public async listReports(): Promise<ReportSummary[]> {
    return Array.from(this.reports, ([id, saved]) => {
      const report = this.model(id, saved);
      return { id, title: `${report.instructorName} ${report.fromDate} – ${report.toDate}`, instructorName: report.instructorName,
        fromDate: report.fromDate, toDate: report.toDate, modified: saved.modified };
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
      if (report.etag !== `"${saved.version}"`) throw new ReportServiceError('This preview report has changed. Reopen it before saving.', 'conflict', 412);
      if (decodeReport(saved.payload).clientId !== report.clientId) throw new ReportServiceError('The report identity cannot change.', 'invalidData');
      saved.payload = payload; saved.version++; saved.modified = new Date().toISOString();
      return this.model(report.id, saved);
    }
    for (const [id, saved] of this.reports) {
      if (decodeReport(saved.payload).clientId === report.clientId) {
        if (saved.payload !== payload) throw new ReportServiceError('This client key already identifies a different saved report. Reopen it before saving.', 'conflict');
        return this.model(id, saved);
      }
    }
    const id = this.nextId++;
    const saved = { payload, version: 1, modified: new Date().toISOString(), receipts: [] };
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
}
