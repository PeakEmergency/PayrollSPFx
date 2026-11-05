import type * as ExcelJSType from 'exceljs';
import { parseDateInput } from '../expenseReport/domain/dates';
import { calculateReport } from '../expenseReport/domain/calculations';
import type { ExpenseReport } from '../expenseReport/domain/models';
import { toExcelDate } from '../timesheet/values';
import { loadExcelJs } from './fillTemplate';

export const EXPENSE_TEMPLATE_FILE = 'Expense Report Template.xlsx';

const FIRST_ROW = 8;
const TOTALS_ROW = 22;

/** Fills a copy of the stored Expense Report template. Only cell values change; styles and formats stay. */
export async function fillExpenseReport(template: ArrayBuffer, report: ExpenseReport): Promise<ArrayBuffer> {
  const ExcelJS = await loadExcelJs();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(template);
  const ws = wb.worksheets[0];

  const put = (address: string, raw: string, kind: 'text' | 'number' | 'date'): void => {
    if (raw.trim() === '') return;
    const cell = ws.getCell(address);
    if (kind === 'text') {
      cell.value = raw;
    } else if (kind === 'number') {
      const n = Number(raw);
      if (!isNaN(n)) cell.value = n;
    } else {
      const iso = parseDateInput(raw);
      if (iso) cell.value = toExcelDate(iso);
    }
  };

  put('A4', report.instructorName, 'text');
  put('K4', report.fromDate, 'date');
  put('M4', report.toDate, 'date');
  ws.getCell('M4').numFmt = 'dd/mm/yy;@'; // the template leaves M4 unformatted
  report.lines.forEach((line, i) => {
    const r = FIRST_ROW + i;
    put(`A${r}`, line.date, 'date');
    put(`C${r}`, line.startCity, 'text');
    put(`D${r}`, line.endCity, 'text');
    put(`E${r}`, line.kilometres, 'number');
    put(`F${r}`, line.rate, 'number');
    put(`I${r}`, line.description, 'text');
    put(`K${r}`, line.gst, 'number');
    put(`L${r}`, line.cost, 'number');
  });

  // Cached results, so previewers that don't recalculate (Outlook, phones) show real totals.
  const setResult = (address: string, raw: string | null): void => {
    const cell = ws.getCell(address);
    if (raw !== null && cell.formula) cell.value = { ...(cell.value as ExcelJSType.CellFormulaValue), result: Number(raw) };
  };
  const totals = calculateReport(report);
  totals.lines.forEach((l, i) => {
    setResult(`G${FIRST_ROW + i}`, l.mileage);
    setResult(`N${FIRST_ROW + i}`, l.amount);
  });
  setResult(`E${TOTALS_ROW}`, totals.kilometres);
  setResult(`G${TOTALS_ROW}`, totals.mileage);
  setResult(`K${TOTALS_ROW}`, totals.gst);
  setResult(`L${TOTALS_ROW}`, totals.cost);
  setResult(`N${TOTALS_ROW}`, totals.amount);

  wb.calcProperties.fullCalcOnLoad = true;
  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}
