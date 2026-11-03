/* eslint-disable @typescript-eslint/no-explicit-any */
import type { ExpenseReport } from './domain/models';
import { encodeReport, decodeReport } from './services/reportCodec';
import { MemoryReportService } from './services/MemoryReportService';
import { SharePointReportService, type HttpTransport } from './services/SharePointReportService';

const roster = [{ id: 'instructor-01', name: 'Test Instructor', active: true }];
function report(): ExpenseReport {
  return { schemaVersion: 1, clientId: 'client-one', instructorId: roster[0].id,
    instructorName: roster[0].name, fromDate: '01/11/26', toDate: '15/11/26',
    lines: Array.from({ length: 13 }, (_, i) => ({ id: `line-${i}`, date: '', startCity: '', endCity: '',
      kilometres: '', rate: '0.48', description: '', gst: '', cost: '' })) };
}
function json(value: unknown, status = 200, etag?: string): Response {
  return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json', ...(etag ? { ETag: etag } : {}) } });
}
type Item = Record<string, unknown> & { Id: number; Payload: string; ClientId: string; version: number };
/** Stateful REST boundary: commits happen before the deliberately interrupted response. */
class SharePointFixture implements HttpTransport {
  public items: Item[] = [];
  public files = new Map<string, Uint8Array>();
  public requests: { url: string; init: RequestInit }[] = [];
  public interruptCreate = false;
  public interruptAttachment = false;
  public denied = false;
  public page = false;
  public omitCollectionEtags = false;
  public nextLinkOverride?: string;
  public failAttachment = false;
  public failRecovery = false;
  public afterReportWrite?: (item: Item) => void;
  public request = async (url: string, init: RequestInit): Promise<Response> => {
    this.requests.push({ url, init });
    if (this.denied) return json({ error: { message: { value: 'Access denied' } } }, 403);
    const decoded = decodeURIComponent(url);
    const method = init.method || 'GET';
    const itemId = /items\((\d+)\)/i.exec(decoded)?.[1];
    const item = itemId ? this.items.find(i => i.Id === Number(itemId)) : undefined;
    if (decoded.includes('AttachmentFiles')) {
      if (!item) return json({ error: { message: 'Missing item' } }, 404);
      const fileName = /(?:FileName=|AttachmentFiles\()'((?:[^']|'')*)'/i.exec(decoded)?.[1]?.replace(/''/g, "'");
      if (method === 'POST') {
        if (this.failAttachment) return json({ error: { message: 'Upload unavailable' } }, 503);
        if (!fileName) throw new Error('Fixture expected a file name');
        if (this.files.has(fileName)) return json({ error: { message: 'Already exists' } }, 409);
        this.files.set(fileName, new Uint8Array(await (init.body as Blob).arrayBuffer()));
        if (this.interruptAttachment) { this.interruptAttachment = false; throw new TypeError('Connection lost after upload'); }
        return json({ FileName: fileName, ServerRelativeUrl: `/sites/expense/Lists/ExpenseReports/Attachments/${item.Id}/${fileName}` }, 200);
      }
      if (decoded.endsWith('/$value')) return this.files.has(fileName!) ? new Response(this.files.get(fileName!)!) : json({}, 404);
      return json({ value: Array.from(this.files.keys(), name => ({ FileName: name, ServerRelativeUrl: `/sites/expense/Lists/ExpenseReports/Attachments/${item.Id}/${name}` })) });
    }
    if (decoded.includes("getbytitle('Instructors')")) {
      if (this.page && !decoded.includes('page=2')) return json({ value: [{ InstructorKey: 'a', Title: 'First', Active: true }], 'odata.nextLink': `${url}&page=2` });
      return json({ value: [{ InstructorKey: 'b', Title: 'Second', Active: true }] });
    }
    if (itemId) {
      if (!item) return json({}, 404);
      if (method === 'POST') {
        const headers = new Headers(init.headers);
        if (headers.get('If-Match') !== `"${item.version}"`) return json({ error: { message: 'Version conflict' } }, 412);
        Object.assign(item, JSON.parse(init.body as string)); item.version++;
        this.afterReportWrite?.(item);
        return new Response(null, { status: 204 });
      }
      return json({ ...item, 'odata.etag': `"${item.version}"` }, 200, `"${item.version}"`);
    }
    if (method === 'POST') {
      const body = JSON.parse(init.body as string);
      if (this.items.some(i => i.ClientId === body.ClientId)) return json({ error: { message: 'Duplicate ClientId' } }, 409);
      const created = { ...body, Id: this.items.length + 1, version: 1, Modified: '2026-10-31T15:00:00Z' } as Item;
      this.items.push(created);
      this.afterReportWrite?.(created);
      if (this.interruptCreate) { this.interruptCreate = false; throw new TypeError('Connection lost after commit'); }
      return json({ Id: created.Id }, 201, '"1"');
    }
    const clientId = /ClientId eq '([^']+)'/.exec(decoded)?.[1];
    if (clientId && this.failRecovery && this.items.length) return json({ error: { message: 'Read temporarily unavailable' } }, 503);
    let matches = clientId ? this.items.filter(i => i.ClientId === clientId) : this.items;
    if (this.page && !clientId && !decoded.includes('page=2')) return json({ value: matches.slice(0, 1), '@odata.nextLink': `${url}&page=2` });
    if (this.page && !clientId) matches = matches.slice(1);
    return json({ value: matches.map(i => ({ ...i, ...(this.omitCollectionEtags ? {} : { 'odata.etag': `"${i.version}"` }) })), ...(this.nextLinkOverride ? { 'odata.nextLink': this.nextLinkOverride } : {}) });
  };
}
const config = { siteUrl: 'https://example.sharepoint.com/sites/expense', reportsListTitle: 'ExpenseReports', instructorsListTitle: 'Instructors' };

describe('strict report codec', () => {
  it('preserves 13 slots, IDs, blanks, zero, arbitrary decimal precision and rate snapshots with ISO dates', () => {
    const original = report();
    original.lines[0] = { ...original.lines[0], date: '02/11/2026', kilometres: '1.010000000000000000001', cost: '0', description: '123', rate: '0.481' };
    const encoded = encodeReport(original);
    expect(JSON.parse(encoded).fromDate).toBe('2026-11-01');
    expect(JSON.parse(encoded).lines[0].date).toBe('2026-11-02');
    const loaded = decodeReport(encoded);
    expect(loaded.lines).toEqual(original.lines.map((l, i) => i === 0 ? { ...l, date: '02/11/26' } : l));
    expect(loaded.fromDate).toBe('01/11/26');
    expect(loaded.lines).toHaveLength(13);
  });
  it.each([
    ['unsupported version', (p: any) => { p.schemaVersion = 2; }],
    ['missing slot', (p: any) => { p.lines.pop(); }],
    ['duplicate ID', (p: any) => { p.lines[1].id = p.lines[0].id; }],
    ['numeric corruption', (p: any) => { p.lines[0].cost = 'NaN'; }],
    ['wrong type', (p: any) => { p.lines[0].gst = 0; }],
    ['invalid date', (p: any) => { p.fromDate = '2026-02-30'; }],
    ['reversed period', (p: any) => { p.toDate = '2026-10-01'; }],
    ['missing field', (p: any) => { delete p.lines[0].rate; }],
    ['unknown field', (p: any) => { p.lines[0].unknown = 'unrecognized'; }]
  ])('rejects %s without silently replacing saved data', (_label, mutate) => {
    const payload = JSON.parse(encodeReport(report())); mutate(payload);
    expect(() => decodeReport(JSON.stringify(payload))).toThrow();
  });
  it('rejects malformed JSON, invalid edited inputs and legacy untyped data', () => {
    expect(() => decodeReport('{')).toThrow();
    expect(() => decodeReport('null')).toThrow();
    const edited = report(); edited.lines[12].date = 'not a date';
    expect(() => encodeReport(edited)).toThrow();
  });
  it('normalizes optional whitespace dates and rejects payloads too large for plain multiline storage', () => {
    const edited = report(); edited.lines[0].date = '   ';
    expect(JSON.parse(encodeReport(edited)).lines[0].date).toBe('');
    edited.lines[0].description = 'a'.repeat(64000);
    expect(() => encodeReport(edited)).toThrow(/capacity/);
  });
});

describe('SharePoint storage', () => {
  it('selects OData v3 explicitly at the SPHttpClient transport boundary', async () => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    await service.saveReport(report());
    for (const request of fixture.requests) {
      const headers = new Headers(request.init.headers);
      expect(headers.get('odata-version')).toBe('');
      expect(headers.get('Accept')).toBe('application/json;odata=nometadata');
    }
  });
  it('uses conditional updates and exposes actual concurrent edit errors', async () => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    const saved = await service.saveReport(report());
    expect(saved.id).toBe(1); expect(saved.etag).toBe('"1"');
    const reopened = await service.loadReport(1);
    expect(reopened.report.lines).toHaveLength(13); expect(reopened.receipts).toEqual([]);
    const updated = await service.saveReport({ ...saved, instructorName: 'Changed snapshot' });
    expect(updated.etag).toBe('"2"');
    await expect(service.saveReport(saved)).rejects.toMatchObject({ kind: 'conflict', status: 412 });
    expect(fixture.items[0].InstructorName).toBe('Changed snapshot');
  });
  it('recovers committed creates after a lost response and retry does not duplicate reports', async () => {
    const fixture = new SharePointFixture(); fixture.interruptCreate = true;
    const service = new SharePointReportService(fixture, config); const input = report();
    const saved = await service.saveReport(input);
    expect(saved.id).toBe(1);
    const retried = await service.saveReport(input);
    expect(retried.id).toBe(1); expect(fixture.items).toHaveLength(1);
    const changed = { ...input, instructorName: 'Different' };
    await expect(service.saveReport(changed)).rejects.toMatchObject({ kind: 'conflict' });
  });
  it('recovers identity across adapter recreation even if the first reconciliation request failed', async () => {
    const fixture = new SharePointFixture(); fixture.interruptCreate = true; fixture.failRecovery = true;
    const service = new SharePointReportService(fixture, config); const input = report();
    await expect(service.saveReport(input)).rejects.toMatchObject({ status: 503 });
    expect(fixture.items).toHaveLength(1);
    fixture.failRecovery = false; fixture.omitCollectionEtags = true;
    const secondService = new SharePointReportService(fixture, config);
    const saved = await secondService.saveReport(input);
    expect(saved.id).toBe(1); expect(saved.etag).toBe('"1"'); expect(fixture.items).toHaveLength(1);
  });
  it('blocks client identity changes when updating a saved report', async () => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    const saved = await service.saveReport(report());
    await expect(service.saveReport({ ...saved, clientId: 'another-client' })).rejects.toMatchObject({ kind: 'invalidData' });
    expect(fixture.items[0].ClientId).toBe('client-one');
  });
  it.each(['create', 'update'])('retains the submitted snapshot when another editor changes content after %s and before verification', async mode => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    const input = mode === 'update' ? await service.saveReport(report()) : report();
    input.lines[0].description = 'My local entry';
    fixture.afterReportWrite = item => {
      const competing = JSON.parse(item.Payload);
      competing.lines[0].description = 'Another editor entry';
      item.Payload = JSON.stringify(competing); item.version++;
    };
    await expect(service.saveReport(input)).rejects.toMatchObject({ kind: 'conflict' });
    expect(input.lines[0].description).toBe('My local entry');
    expect(JSON.parse(fixture.items[0].Payload).lines[0].description).toBe('Another editor entry');
    expect(fixture.items).toHaveLength(1);
  });
  it('paginates instructors and reports using server next links', async () => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    await service.saveReport(report()); await service.saveReport({ ...report(), clientId: 'second' }); fixture.page = true;
    expect((await service.getInstructors()).map(i => i.name)).toEqual(['First', 'Second']);
    expect((await service.listReports()).map(i => i.id)).toEqual([1, 2]);
  });
  it('does not substitute preview data for permission or configuration failures', async () => {
    const fixture = new SharePointFixture(); fixture.denied = true;
    const service = new SharePointReportService(fixture, config);
    await expect(service.getInstructors()).rejects.toMatchObject({ kind: 'permission', status: 403 });
    await expect(service.saveReport(report())).rejects.toMatchObject({ kind: 'permission', status: 403 });
    expect(fixture.items).toHaveLength(0);
    expect(() => new SharePointReportService(fixture, { ...config, reportsListTitle: '' })).toThrow();
  });
  it('does not follow unsafe or repeated server pagination links', async () => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    fixture.nextLinkOverride = 'https://other.example/_api/items';
    await expect(service.listReports()).rejects.toMatchObject({ kind: 'invalidData' });
    fixture.nextLinkOverride = `${config.siteUrl}/_api/repeated`;
    await expect(service.listReports()).rejects.toMatchObject({ kind: 'invalidData' });
  });
  it('keeps saved identity through failed receipts and resolves interrupted uploads by comparing bytes', async () => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    const saved = await service.saveReport(report()); const file = new File(['receipt bytes'], 'receipt.pdf');
    fixture.denied = true;
    await expect(service.uploadReceipt(saved.id!, file)).rejects.toMatchObject({ kind: 'permission', status: 403 });
    fixture.denied = false; fixture.interruptAttachment = true;
    const uploaded = await service.uploadReceipt(saved.id!, file);
    expect(uploaded.name).toBe(file.name);
    expect(await service.uploadReceipt(saved.id!, file)).toEqual(uploaded);
    expect(fixture.files.size).toBe(1); expect(fixture.items).toHaveLength(1);
    expect((await service.loadReport(saved.id!)).receipts).toEqual([uploaded]);
    await expect(service.uploadReceipt(saved.id!, new File(['different bytes'], file.name))).rejects.toMatchObject({ kind: 'conflict' });
  });
  it('rejects corrupted stored report data and missing ETags', async () => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    const saved = await service.saveReport(report());
    await expect(service.saveReport({ ...saved, etag: undefined })).rejects.toMatchObject({ kind: 'conflict' });
    fixture.items[0].Payload = '{';
    await expect(service.loadReport(1)).rejects.toThrow();
  });
  it('reports an uncommitted receipt failure without losing the report and retries only the receipt', async () => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    const saved = await service.saveReport(report()); const file = new File(['receipt'], "receipt's #1.pdf");
    fixture.failAttachment = true;
    await expect(service.uploadReceipt(saved.id!, file)).rejects.toMatchObject({ kind: 'server', status: 503 });
    expect((await service.loadReport(saved.id!)).receipts).toEqual([]);
    fixture.failAttachment = false;
    expect((await service.uploadReceipt(saved.id!, file)).name).toBe(file.name);
    expect(fixture.items).toHaveLength(1); expect(fixture.files.size).toBe(1);
  });
  it.each([
    ['receipt#1.txt', 'receipt%231.txt'],
    ['literal%23.txt', 'literal%2523.txt'],
    ['space and #.txt', 'space%20and%20%23.txt']
  ])('escapes the decoded SharePoint attachment path for %s without changing the file name', async (name, encodedName) => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    const saved = await service.saveReport(report());
    const receipt = await service.uploadReceipt(saved.id!, new File(['receipt'], name));
    expect(receipt).toEqual({ name, url: `${config.siteUrl}/Lists/ExpenseReports/Attachments/1/${encodedName}` });
    expect(new URL(receipt.url).hash).toBe('');
    expect((await service.loadReport(saved.id!)).receipts[0]).toEqual(receipt);
  });
  it('rejects receipt names containing control characters before making an upload request', async () => {
    const fixture = new SharePointFixture(); const service = new SharePointReportService(fixture, config);
    const saved = await service.saveReport(report());
    await expect(service.uploadReceipt(saved.id!, new File(['receipt'], 'bad\u001f.pdf'))).rejects.toMatchObject({ kind: 'invalidData' });
    expect(fixture.files.size).toBe(0);
  });
});

describe('independent memory preview', () => {
  it('round trips cloned values, enforces ETags and isolates service instances', async () => {
    const memory = new MemoryReportService(roster); const other = new MemoryReportService(roster);
    const original = report(); const saved = await memory.saveReport(original);
    saved.lines[0].description = 'Edit';
    expect((await memory.loadReport(saved.id!)).report.lines[0].description).toBe('');
    expect(await other.listReports()).toEqual([]);
    const updated = await memory.saveReport(saved);
    await expect(memory.saveReport(saved)).rejects.toMatchObject({ kind: 'conflict' });
    expect(updated.etag).not.toEqual(saved.etag);
    const receipt = await memory.uploadReceipt(saved.id!, new File(['a'], 'a.txt'));
    expect((await memory.loadReport(saved.id!)).receipts).toEqual([receipt]);
  });
});
