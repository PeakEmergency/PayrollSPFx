import { ColumnDef, ColumnKind, TemplateDef, TimesheetType } from './types';

// Cell map of the real templates (spec section 6). Widths are the template
// column widths in Excel units × 7 px.
const col = (key: string, c: string, header: string, kind: ColumnKind, excelWidth: number, options?: number[]): ColumnDef =>
  ({ key, col: c, header, kind, width: Math.round(excelWidth * 7), options });

const quarterHours: number[] = Array.from({ length: 32 }, (_, i) => (i + 1) * 0.25);

export const TEMPLATES: Record<TimesheetType, TemplateDef> = {
  Admin: {
    type: 'Admin',
    title: 'ADMIN TIMESHEET',
    sheetName: 'Admin Timesheet Template',
    fileName: 'Admin Timesheet Template.xlsx',
    nameCell: 'A4',
    fromCell: 'G4',
    toCell: 'I4',
    firstRow: 7,
    rowCount: 16,
    headerRow: 6,
    columns: [
      col('date', 'A', 'DATE', 'date', 10.71),
      col('start', 'C', 'START TIME', 'time', 10.71),
      col('end', 'D', 'END TIME', 'time', 10.71),
      col('deductions', 'E', 'DEDUCTIONS', 'dropdown', 15.71, [0, 0.5, 1, 1.25, 1.5, 1.75, 2]),
      col('hours', 'F', 'TOTAL HOURS(s)', 'computed', 10.14),
      col('rate', 'G', 'PAY RATE', 'number', 14.14),
      col('amount', 'I', 'AMOUNT', 'computed', 15)
    ],
    totals: [
      { key: 'deductions', cell: 'E25' },
      { key: 'hours', cell: 'F25' },
      { key: 'amount', cell: 'I25' }
    ],
    grandTotalKey: 'amount'
  },
  Instructor: {
    type: 'Instructor',
    title: 'INSTRUCTOR TIMESHEET',
    sheetName: 'Instructor Timesheet Template',
    fileName: 'Instructor Timesheet Template.xlsx',
    nameCell: 'A4',
    fromCell: 'J4',
    toCell: 'M4',
    firstRow: 8,
    rowCount: 13,
    headerRow: 7,
    columns: [
      col('date', 'A', 'DATE', 'date', 10.71),
      col('course', 'C', 'COURSE #', 'upper', 14.29),
      col('days', 'D', 'DAY(s)', 'dropdown', 14.29, [0.5, 1]),
      col('dayRate', 'E', 'RATE', 'dropdown', 14.29, [0, 150, 225, 280, 320, 340, 350]),
      col('instrTotal', 'F', 'TOTAL', 'computed', 1.14),
      col('adminHours', 'G', 'HOUR(s)', 'dropdown', 14.29, quarterHours),
      col('adminRate', 'H', 'RATE', 'dropdown', 14.29, [25]),
      col('adminTotal', 'I', 'TOTAL', 'computed', 1.14),
      col('travel', 'J', 'TRAVEL TIME', 'dropdown', 14.29, [100]),
      col('perDiem', 'K', 'PER DIEM', 'dropdown', 14.29, [47]),
      col('otherTotal', 'L', 'TOTAL', 'computed', 1.14),
      col('amount', 'M', 'AMOUNT', 'computed', 13.29)
    ],
    totals: [
      { key: 'days', cell: 'D22' },
      { key: 'adminHours', cell: 'G22' },
      { key: 'travel', cell: 'J22' },
      { key: 'perDiem', cell: 'K22' },
      { key: 'amount', cell: 'M22' }
    ],
    grandTotalKey: 'amount'
  },
  Medic: {
    type: 'Medic',
    title: 'MEDIC TIMESHEET',
    sheetName: 'Medic Timesheet Template',
    fileName: 'Medic Timesheet Template.xlsx',
    nameCell: 'A4',
    fromCell: 'H4',
    toCell: 'K4',
    firstRow: 8,
    rowCount: 13,
    headerRow: 7,
    columns: [
      col('date', 'A', 'DATE', 'date', 10.71),
      col('client', 'C', 'CLIENT', 'upper', 30.71),
      col('start', 'D', 'TIME START', 'time', 10.71),
      col('end', 'E', 'TIME END', 'time', 10.71),
      col('hours', 'F', 'HOURS(s)', 'computed', 10.71),
      {
        ...col('dayRate', 'G', 'RATE', 'dropdown', 10, [250, 350, 450, 550]),
        // Template lookup table M6:N9.
        optionLabels: [
          'Medic – Level NUEC-A/AP (PSL 8)',
          'Medic – Level PCP(PSL 9)',
          'Medic – Level ACP (PSL 10)',
          'Medic – Level CCP/MD (PSL 11)'
        ]
      },
      col('medicTotal', 'H', 'TOTAL', 'computed', 0.86),
      col('travel', 'I', 'TRAVEL TIME', 'dropdown', 14.29, [100]),
      col('perDiem', 'J', 'PER DIEM', 'dropdown', 14.29, [47]),
      col('otherTotal', 'K', 'TOTAL', 'computed', 0.86),
      col('amount', 'L', 'AMOUNT', 'computed', 13.86)
    ],
    totals: [
      { key: 'hours', cell: 'F22' },
      { key: 'travel', cell: 'I22' },
      { key: 'perDiem', cell: 'J22' },
      { key: 'amount', cell: 'L22' }
    ],
    grandTotalKey: 'amount'
  }
};
