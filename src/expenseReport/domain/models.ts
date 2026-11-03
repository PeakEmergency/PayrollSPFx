/* eslint-disable @rushstack/no-new-null -- null explicitly represents unavailable calculated totals in the shared domain contract. */
/** Editable form values remain raw strings; placeholders never enter this model. */
export interface ExpenseLine {
  id: string;
  date: string;
  startCity: string;
  endCity: string;
  kilometres: string;
  rate: string;
  description: string;
  gst: string;
  cost: string;
}

export interface ExpenseReport {
  schemaVersion: 1;
  clientId: string;
  id?: number;
  etag?: string;
  instructorId: string;
  instructorName: string;
  fromDate: string;
  toDate: string;
  lines: ExpenseLine[];
}

export interface ReceiptInfo { name: string; url: string; }
export interface ReportSummary {
  id: number;
  title: string;
  instructorName: string;
  fromDate: string;
  toDate: string;
  modified: string;
}
export interface ValidationIssue { field: string; message: string; }
export interface ReportTotals {
  lines: { mileage: string | null; amount: string | null }[];
  kilometres: string | null;
  mileage: string | null;
  gst: string | null;
  cost: string | null;
  amount: string | null;
}
