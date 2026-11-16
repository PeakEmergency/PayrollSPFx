import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { TEMPLATES } from '../timesheet/definitions';
import { PayPeriod, toPayPeriod } from '../timesheet/payPeriods';
import { PayRates, rateKey, toPayRates } from '../timesheet/rates';
import { formatEmployeeName } from '../timesheet/naming';
import { TimesheetType } from '../timesheet/types';
import { SubmissionRecord, SubmissionStatus, SubmissionStore, TemplateSource } from './submit';

export const PAY_PERIOD_LIST = 'Pay Period Schedule';
export const PAY_RATE_LIST = 'Pay Rates';
export const EMPLOYEE_LIST = 'Employees';

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
          EmailTo: r.emailTo,
          EmailSubject: r.emailSubject,
          EmailBody: r.emailBody,
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

  public load(type: TimesheetType): Promise<ArrayBuffer> {
    return this.loadFile(TEMPLATES[type].fileName);
  }

  public async loadFile(fileName: string): Promise<ArrayBuffer> {
    const filePath = `${this.libraryServerRelativeUrl.replace(/\/$/, '')}/${fileName}`;
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

export class SpPayPeriodSource {
  private readonly itemsUrl: string;

  public constructor(private readonly client: SPHttpClient, webUrl: string, listTitle: string) {
    this.itemsUrl = `${webUrl}/_api/web/lists/getbytitle('${odataString(listTitle)}')/items`;
  }

  public async getPayPeriods(): Promise<PayPeriod[]> {
    const res = await check(
      'Loading pay periods',
      this.client.get(`${this.itemsUrl}?$select=Year,PeriodNumber,FromDate,ToDate&$top=5000`, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      })
    );
    return (await res.json()).value.map(toPayPeriod);
  }
}

export class SpPayRateSource {
  private readonly itemsUrl: string;

  public constructor(private readonly client: SPHttpClient, webUrl: string, listTitle: string) {
    this.itemsUrl = `${webUrl}/_api/web/lists/getbytitle('${odataString(listTitle)}')/items`;
  }

  /** Every row, keyed by rateKey of the name ('Rosie Trinh' → 'TRINH, ROSIE'). */
  public async getAllRates(): Promise<Record<string, PayRates>> {
    const select = 'Title,AdminRate,MedicPSL,RateAP,RateASU,RateNUECAFull,RateNUECABridging,RateNUECI,RateNUECB';
    const res = await check(
      'Loading pay rates',
      this.client.get(`${this.itemsUrl}?$select=${select}&$top=5000`, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      })
    );
    const all: Record<string, PayRates> = {};
    (await res.json()).value.forEach((row: Record<string, unknown>) => {
      all[rateKey(formatEmployeeName(null, null, String(row.Title || '')))] = toPayRates(row);
    });
    return all;
  }
}

export interface Employee {
  key: string;
  name: string;
}

export class SpEmployeeSource {
  private readonly itemsUrl: string;

  public constructor(private readonly client: SPHttpClient, webUrl: string, listTitle: string) {
    this.itemsUrl = `${webUrl}/_api/web/lists/getbytitle('${odataString(listTitle)}')/items`;
  }

  /** Active people in the Employees list, A to Z. Title is 'LAST, FIRST'. */
  public async getEmployees(): Promise<Employee[]> {
    const res = await check(
      'Loading employees',
      this.client.get(`${this.itemsUrl}?$select=Title,InstructorKey,Active&$orderby=Title&$top=5000`, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      })
    );
    return (await res.json()).value
      .filter((r: { Active?: boolean }) => r.Active !== false)
      .map((r: { Title: string; InstructorKey: string }) => ({ key: r.InstructorKey, name: r.Title }));
  }
}
