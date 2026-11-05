import type { ExpenseReport, ReceiptInfo, ReportStatus, ReportSummary } from '../domain/models';
import { formatDateInput, parseDateInput } from '../domain/dates';
import { ReportServiceError, type IReportService } from './IReportService';
import { decodeReport, encodeReport } from './reportCodec';
export interface HttpTransport { request(url: string, init: RequestInit): Promise<Response>; }
export interface SharePointReportServiceConfig { siteUrl: string; reportsListTitle: string; }
type JsonObject = Record<string, unknown>;
function invalid(message: string): never { throw new ReportServiceError(message, 'invalidData'); }
function object(value: unknown): JsonObject {
  if (!value || typeof value !== 'object' || Array.isArray(value)) invalid('SharePoint returned malformed data.');
  return value as JsonObject;
}
function unwrap(value: unknown): JsonObject {
  const data = object(value);
  return data.d !== undefined ? object(data.d) : data;
}
function string(value: unknown, name: string): string {
  if (typeof value !== 'string') invalid(`SharePoint returned an invalid ${name} field. Check the configured list schema.`);
  return value;
}
function id(value: unknown): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 1) invalid('SharePoint returned an invalid report ID.');
  return value;
}
const STATUSES: string[] = ['Pending', 'Queued', 'Sent', 'Email failed'];
const SUBMITTED = 'This report has been submitted and can no longer be changed. Start a new report for more expenses.';
function status(value: unknown): ReportStatus {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value !== 'string' || !STATUSES.includes(value)) invalid('SharePoint returned an invalid Status field. Check the configured list schema.');
  return value as ReportStatus;
}
function literal(value: string): string { return encodeURIComponent(value.replace(/'/g, "''")).replace(/'/g, '%27'); }
function isoDate(value: unknown): string {
  const raw = string(value, 'date');
  const day = raw.slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day) || !formatDateInput(day) || parseDateInput(formatDateInput(day)) !== day) invalid('SharePoint returned an invalid date.');
  return day;
}
export function validateReceiptName(name: string): void {
  const hasControlCharacter = Array.from(name).some(character => character.charCodeAt(0) < 32 || character.charCodeAt(0) === 127);
  if (!name.trim() || hasControlCharacter || /["*:<>?\/\\|]/.test(name) || /[. ]$/.test(name) || name.length > 255) {
    throw new ReportServiceError('Choose a receipt file with a valid SharePoint file name (maximum 255 characters).', 'invalidData');
  }
}
export async function equalReceiptBytes(file: File, existing: ArrayBuffer): Promise<boolean> {
  const incoming = new Uint8Array(await file.arrayBuffer());
  const saved = new Uint8Array(existing);
  return incoming.length === saved.length && incoming.every((byte, index) => byte === saved[index]);
}
export class SharePointReportService implements IReportService {
  private readonly siteUrl: string;
  private readonly reportsUrl: string;
  public constructor(private readonly transport: HttpTransport, config: SharePointReportServiceConfig) {
    let site: URL;
    try { site = new URL(config.siteUrl); } catch { throw new ReportServiceError('Configure an absolute SharePoint site URL.', 'configuration'); }
    if (site.protocol !== 'https:' || site.search || site.hash || site.username || site.password || !config.reportsListTitle.trim()) {
      throw new ReportServiceError('Configure an HTTPS SharePoint site and a nonblank list title.', 'configuration');
    }
    this.siteUrl = site.href.replace(/\/$/, '');
    const list = (title: string): string => `${this.siteUrl}/_api/web/lists/getbytitle('${literal(title)}')`;
    this.reportsUrl = list(config.reportsListTitle);
  }
  private async request(url: string, init: RequestInit = {}): Promise<Response> {
    const headers = new Headers(init.headers);
    headers.set('Accept', 'application/json;odata=nometadata');
    // SPHttpClient defaults to OData v4; the REST payloads here explicitly use v3.
    headers.set('odata-version', '');
    let response: Response;
    try { response = await this.transport.request(url, { ...init, headers }); }
    catch (error) {
      if (error instanceof ReportServiceError) throw error;
      throw new ReportServiceError(`SharePoint could not be reached. Your input is retained. ${error instanceof Error ? error.message : ''}`, 'network');
    }
    if (response.ok) return response;
    let detail = '';
    try {
      const body = object(await response.json());
      const error = object(body.error || body['odata.error']);
      const message = error.message;
      detail = typeof message === 'string' ? message : typeof message === 'object' && message ? String(object(message).value || '') : '';
    } catch { /* Status remains authoritative when the error body is not JSON. */ }
    const status = response.status;
    const kind = status === 401 || status === 403 ? 'permission' : status === 409 || status === 412 ? 'conflict' : status === 404 ? 'notFound' : status === 400 ? 'configuration' : 'server';
    const advice = kind === 'permission' ? 'You do not have permission for this SharePoint operation.'
      : kind === 'conflict' ? 'This report or receipt has changed. Reopen the saved report before overwriting anything.'
      : kind === 'notFound' || kind === 'configuration' ? 'Check the configured site, list names, fields, and report ID.' : 'SharePoint could not complete the operation.';
    throw new ReportServiceError(`${advice} (HTTP ${status})${detail ? ` ${detail}` : ''}`, kind, status);
  }
  private async json(url: string, init?: RequestInit): Promise<{ data: JsonObject; response: Response }> {
    const response = await this.request(url, init);
    let data: unknown;
    try { data = await response.json(); } catch { invalid('SharePoint returned unreadable JSON.'); }
    return { data: unwrap(data), response };
  }
  private async collection(url: string): Promise<JsonObject[]> {
    const rows: JsonObject[] = [];
    const seen = new Set<string>();
    let next: string | undefined = url;
    while (next) {
      if (seen.has(next)) invalid('SharePoint returned a repeated pagination link.');
      seen.add(next);
      const { data } = await this.json(next);
      const values = data.value !== undefined ? data.value : data.results;
      if (!Array.isArray(values)) invalid('SharePoint returned a malformed list response.');
      rows.push(...values.map(object));
      const link: unknown = data['@odata.nextLink'] || data['odata.nextLink'] || data.__next;
      next = undefined;
      if (link !== undefined) {
        const resolved = new URL(string(link, 'pagination'), this.siteUrl + '/');
        // A transport carries the user's credentials: never follow another site's link.
        if (resolved.origin !== new URL(this.siteUrl).origin || !resolved.href.startsWith(this.siteUrl + '/_api/')) invalid('SharePoint returned an unsafe pagination link.');
        next = resolved.href;
      }
    }
    return rows;
  }
  public async listReports(ownerEmail: string): Promise<ReportSummary[]> {
    const rows = await this.collection(`${this.reportsUrl}/items?$expand=Author&$filter=Author/EMail%20eq%20'${literal(ownerEmail)}'&$select=Id,Title,InstructorName,FromDate,ToDate,Modified,Status,Author/EMail&$orderby=Modified desc&$top=5000`);
    return rows.map(row => ({ id: id(row.Id), title: string(row.Title, 'Title'), instructorName: string(row.InstructorName, 'InstructorName'),
      fromDate: formatDateInput(isoDate(row.FromDate)), toDate: formatDateInput(isoDate(row.ToDate)), modified: string(row.Modified, 'Modified'), status: status(row.Status) }));
  }
  private itemUrl(reportId: number): string { return `${this.reportsUrl}/items(${id(reportId)})`; }
  private decodeItem(data: JsonObject, headerEtag?: string | null): ExpenseReport {
    const report = decodeReport(string(data.Payload, 'Payload'));
    const metadata = data.__metadata === undefined ? undefined : object(data.__metadata);
    const etag = headerEtag || data['@odata.etag'] || data['odata.etag'] || metadata?.etag;
    if (typeof etag !== 'string' || !etag || etag === '*') invalid('SharePoint omitted the report version (ETag); concurrent edits cannot be protected.');
    if (data.SchemaVersion !== 1 || data.ClientId !== report.clientId || data.InstructorKey !== report.instructorId || data.InstructorName !== report.instructorName || isoDate(data.FromDate) !== parseDateInput(report.fromDate) || isoDate(data.ToDate) !== parseDateInput(report.toDate)) {
      invalid('The report headers do not match its saved payload. Reconcile the stored item before editing.');
    }
    return { ...report, id: id(data.Id), etag, status: status(data.Status) };
  }
  private async readReport(reportId: number): Promise<ExpenseReport> {
    // Without $select SharePoint omits ClientId from the item, which fails the header check.
    const { data, response } = await this.json(`${this.itemUrl(reportId)}?$select=Id,ClientId,InstructorKey,InstructorName,FromDate,ToDate,SchemaVersion,Status,Payload`);
    return this.decodeItem(data, response.headers.get('ETag'));
  }
  private async verifySavedReport(reportId: number, payload: string): Promise<ExpenseReport> {
    const saved = await this.readReport(reportId);
    if (saved.id !== reportId || encodeReport(saved) !== payload) {
      throw new ReportServiceError('The report changed after your save. Your local entries are retained; reopen the saved report to review the other changes.', 'conflict', 412);
    }
    return saved;
  }
  private receipt(row: JsonObject): ReceiptInfo {
    const name = string(row.FileName, 'receipt name');
    const path = string(row.ServerRelativeUrl, 'receipt URL');
    if (!path.startsWith('/') || path.startsWith('//')) invalid('SharePoint returned an invalid server-relative receipt path.');
    // ServerRelativeUrl is a decoded SharePoint path. Escape each segment before
    // parsing so literal # and % characters remain part of the file name.
    const url = new URL(path.split('/').map(segment => encodeURIComponent(segment)).join('/'), this.siteUrl);
    if (url.origin !== new URL(this.siteUrl).origin || !url.pathname.startsWith(new URL(this.siteUrl).pathname.replace(/\/$/, '') + '/')) invalid('SharePoint returned an unsafe receipt URL.');
    return { name, url: url.href };
  }
  private async receipts(reportId: number): Promise<ReceiptInfo[]> {
    return (await this.collection(`${this.itemUrl(reportId)}/AttachmentFiles?$select=FileName,ServerRelativeUrl`)).map(row => this.receipt(row));
  }
  public async loadReport(reportId: number): Promise<{ report: ExpenseReport; receipts: ReceiptInfo[] }> {
    const report = await this.readReport(reportId);
    const receipts = await this.receipts(reportId);
    return { report, receipts };
  }
  private async findClient(clientId: string, payload: string): Promise<ExpenseReport | undefined> {
    const rows = await this.collection(`${this.reportsUrl}/items?$filter=ClientId%20eq%20'${literal(clientId)}'&$top=2`);
    if (rows.length > 1) invalid('Two reports share one ClientId. Delete the extra item in the ExpenseReports list before saving.');
    if (!rows.length) return undefined;
    // nometadata collection responses may omit item ETags. Read the item so the
    // authoritative HTTP version header is available before returning identity.
    const existing = await this.readReport(id(rows[0].Id));
    if (existing.status) throw new ReportServiceError(SUBMITTED, 'conflict');
    if (encodeReport(existing) !== payload) throw new ReportServiceError('A report with this client key is already saved with different content. Reopen it to continue safely.', 'conflict');
    return existing;
  }
  public async saveReport(report: ExpenseReport): Promise<ExpenseReport> {
    const payload = encodeReport(report);
    const stored = object(JSON.parse(payload));
    const body = JSON.stringify({ Title: `${report.instructorName} ${stored.fromDate} – ${stored.toDate}`.slice(0, 255),
      ClientId: report.clientId, InstructorKey: report.instructorId, InstructorName: report.instructorName,
      FromDate: stored.fromDate, ToDate: stored.toDate, SchemaVersion: 1, Payload: payload });
    if (report.id !== undefined) {
      if (!report.etag || report.etag === '*') throw new ReportServiceError('Reopen this report to obtain its current version before saving.', 'conflict');
      const current = await this.readReport(report.id);
      if (current.status) throw new ReportServiceError(SUBMITTED, 'conflict');
      if (current.clientId !== report.clientId) throw new ReportServiceError('The stable report client identity cannot change.', 'invalidData');
      await this.request(this.itemUrl(report.id), { method: 'POST', headers: { 'Content-Type': 'application/json;odata=nometadata', 'X-HTTP-Method': 'MERGE', 'If-Match': report.etag }, body });
      return this.verifySavedReport(report.id, payload);
    }
    const existing = await this.findClient(report.clientId, payload);
    if (existing) return existing;
    try {
      const { data } = await this.json(`${this.reportsUrl}/items`, { method: 'POST', headers: { 'Content-Type': 'application/json;odata=nometadata' }, body });
      return await this.verifySavedReport(id(data.Id), payload);
    } catch (error) {
      if (!(error instanceof ReportServiceError) || !['network', 'conflict', 'server'].includes(error.kind)) throw error;
      // A create may have committed despite a lost response. The indexed ClientId
      // makes this lookup/retry safe across component and browser lifetimes.
      const recovered = await this.findClient(report.clientId, payload);
      if (recovered) return recovered;
      throw error;
    }
  }
  private async existingReceipt(reportId: number, file: File): Promise<ReceiptInfo | undefined> {
    const found = (await this.receipts(reportId)).find(receipt => receipt.name.toLowerCase() === file.name.toLowerCase());
    if (!found) return undefined;
    const response = await this.request(`${this.itemUrl(reportId)}/AttachmentFiles('${literal(found.name)}')/$value`);
    if (!(await equalReceiptBytes(file, await response.arrayBuffer()))) throw new ReportServiceError(`A different receipt named "${found.name}" is already attached. Rename your file before retrying.`, 'conflict');
    return found;
  }
  public async uploadReceipt(reportId: number, file: File): Promise<ReceiptInfo> {
    id(reportId); validateReceiptName(file.name);
    const existing = await this.existingReceipt(reportId, file);
    if (existing) return existing;
    try {
      const { data } = await this.json(`${this.itemUrl(reportId)}/AttachmentFiles/add(FileName='${literal(file.name)}')`, { method: 'POST', headers: { 'Content-Type': 'application/octet-stream' }, body: file });
      return this.receipt(data);
    } catch (error) {
      if (!(error instanceof ReportServiceError) || !['network', 'conflict', 'server'].includes(error.kind)) throw error;
      const recovered = await this.existingReceipt(reportId, file);
      if (recovered) return recovered;
      throw error;
    }
  }
  private async merge(reportId: number, fields: JsonObject, etag = '*'): Promise<void> {
    await this.request(this.itemUrl(reportId), { method: 'POST', headers: { 'Content-Type': 'application/json;odata=nometadata', 'X-HTTP-Method': 'MERGE', 'If-Match': etag }, body: JSON.stringify(fields) });
  }
  public async setStatus(reportId: number, newStatus: ReportStatus, etag?: string): Promise<void> {
    await this.merge(reportId, { Status: newStatus || null }, etag);
  }
  public async deleteAttachment(reportId: number, fileName: string): Promise<void> {
    try {
      await this.request(`${this.itemUrl(reportId)}/AttachmentFiles/getByFileName('${literal(fileName)}')`, { method: 'POST', headers: { 'Content-Type': 'application/json;odata=nometadata', 'X-HTTP-Method': 'DELETE', 'If-Match': '*' } });
    } catch (error) {
      if (!(error instanceof ReportServiceError) || error.status !== 404) throw error;
    }
  }
  public async attachFile(reportId: number, fileName: string, content: ArrayBuffer): Promise<void> {
    await this.deleteAttachment(reportId, fileName);
    await this.request(`${this.itemUrl(reportId)}/AttachmentFiles/add(FileName='${literal(fileName)}')`, { method: 'POST', headers: { 'Content-Type': 'application/octet-stream' }, body: content });
  }
  public async queueEmail(reportId: number, email: { to: string; subject: string; body: string }): Promise<void> {
    await this.merge(reportId, { EmailTo: email.to, EmailSubject: email.subject, EmailBody: email.body, StatusDetail: '', Status: 'Queued' });
  }
}
