import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { TEMPLATES } from '../timesheet/definitions';
import { TimesheetType } from '../timesheet/types';
import { SubmissionRecord, SubmissionStatus, SubmissionStore, TemplateSource } from './submit';

const jsonHeaders = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata'
};

const odataString = (s: string): string => encodeURIComponent(s.replace(/'/g, "''"));

/** Throws with the step name and SharePoint's own error text on non-2xx. */
async function check(step: string, request: Promise<SPHttpClientResponse>): Promise<SPHttpClientResponse> {
  let res: SPHttpClientResponse;
  try {
    res = await request;
  } catch (err) {
    // Network-level failure (e.g. "Failed to fetch"): no response at all.
    throw new Error(`${step} failed: ${err instanceof Error ? err.message : String(err)}`);
  }
  if (res.ok) return res;
  let message = `${res.status} ${res.statusText}`.trim();
  try {
    const body = await res.json();
    message = body?.['odata.error']?.message?.value || body?.error?.message?.value || body?.error?.message || message;
  } catch {
    // not JSON; keep the status text
  }
  throw new Error(`${step} failed (${res.status}): ${message}`);
}

// Noon UTC keeps a date-only field on the same calendar day in any North American time zone.
const dateOnly = (iso: string): string => `${iso}T12:00:00Z`;

export class SpSubmissionStore implements SubmissionStore {
  private readonly itemsUrl: string;

  public constructor(private readonly client: SPHttpClient, webUrl: string, listTitle: string) {
    this.itemsUrl = `${webUrl}/_api/web/lists/getbytitle('${odataString(listTitle)}')/items`;
  }

  public async create(r: SubmissionRecord): Promise<number> {
    const res = await check(
      'Saving the list item',
      this.client.post(this.itemsUrl, SPHttpClient.configurations.v1, {
        headers: jsonHeaders,
        body: JSON.stringify({
          Title: r.title,
          TimesheetType: r.type,
          EmployeeId: r.employeeId,
          PayPeriodStart: dateOnly(r.from),
          PayPeriodEnd: dateOnly(r.to),
          TotalAmount: r.totalAmount,
          LineItems: r.lineItemsJson,
          Status: 'Pending'
        })
      })
    );
    return (await res.json()).Id;
  }

  public async attach(id: number, fileName: string, content: ArrayBuffer): Promise<void> {
    await check(
      'Attaching the file',
      this.client.post(`${this.itemsUrl}(${id})/AttachmentFiles/add(FileName='${odataString(fileName)}')`, SPHttpClient.configurations.v1, {
        body: content
      })
    );
  }

  public async setStatus(id: number, status: SubmissionStatus, detail?: string): Promise<void> {
    await check(
      'Updating the status',
      this.client.post(`${this.itemsUrl}(${id})`, SPHttpClient.configurations.v1, {
        headers: { ...jsonHeaders, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
        body: JSON.stringify({ Status: status, StatusDetail: detail || '' })
      })
    );
  }
}

export class SpTemplateSource implements TemplateSource {
  public constructor(
    private readonly client: SPHttpClient,
    private readonly webUrl: string,
    private readonly libraryServerRelativeUrl: string
  ) {}

  public async load(type: TimesheetType): Promise<ArrayBuffer> {
    const filePath = `${this.libraryServerRelativeUrl.replace(/\/$/, '')}/${TEMPLATES[type].fileName}`;
    const res = await check(
      `Loading template ${filePath}`,
      this.client.get(
        `${this.webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${odataString(filePath)}')/$value`,
        SPHttpClient.configurations.v1
      )
    );
    return res.arrayBuffer();
  }
}
