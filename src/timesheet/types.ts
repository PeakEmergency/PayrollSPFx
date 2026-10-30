/* eslint-disable @rushstack/no-new-null -- null marks a blank cell; lines are stored as JSON */
export type TimesheetType = 'Admin' | 'Instructor' | 'Medic';

export type ColumnKind = 'date' | 'time' | 'upper' | 'number' | 'dropdown' | 'computed';

export interface ColumnDef {
  key: string;
  col: string;
  header: string;
  kind: ColumnKind;
  options?: number[];
  width: number;
}

export interface TemplateDef {
  type: TimesheetType;
  title: string;
  sheetName: string;
  fileName: string;
  nameCell: string;
  fromCell: string;
  toCell: string;
  firstRow: number;
  rowCount: number;
  headerRow: number;
  columns: ColumnDef[];
  /** On-screen totals, same cells as the template totals row. */
  totals: { key: string; cell: string }[];
  grandTotalKey: string;
}

/** Dates are 'yyyy-mm-dd', times 'HH:MM'. */
export type Line = Record<string, string | number | null>;

export interface TimesheetData {
  type: TimesheetType;
  name: string;
  from: string | null;
  to: string | null;
  lines: Line[];
}
