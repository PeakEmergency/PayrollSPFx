import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { EXPENSE_TEMPLATE_FILE, fillExpenseReport } from './fillExpenseReport';
import type { ExpenseLine, ExpenseReport } from '../expenseReport/domain/models';

// Compiled to lib-commonjs/excel, so the repo root is two levels up.
const fixture = fs.readFileSync(path.resolve(__dirname, '../../Templates/xlsx', EXPENSE_TEMPLATE_FILE));
// Copy into an ArrayBuffer of Jest's sandbox realm; Buffer's own fails jszip's instanceof check.
const toArrayBuffer = (b: Buffer): ArrayBuffer => new Uint8Array(b).buffer;

const blank = (i: number): ExpenseLine => ({
  id: String(i), date: '', startCity: '', endCity: '', kilometres: '', rate: '0.48', description: '', gst: '', cost: ''
});
const lines = (patch: Partial<ExpenseLine>[]): ExpenseLine[] =>
  Array.from({ length: 13 }, (_, i) => ({ ...blank(i), ...(patch[i] || {}) }));
const base = { schemaVersion: 1 as const, clientId: 'c', instructorId: '1', instructorName: 'TRINH, ROSIE' };
const report: ExpenseReport = {
  ...base, fromDate: '05/11/26', toDate: '18/11/26',
  lines: lines([
    { date: '05/11/26', startCity: 'Burnaby V5H 2A9', endCity: 'Golden V0A 1H1', kilometres: '718' },
    { description: 'Parking', gst: '1.5', cost: '10' }
  ])
};
const blankReport: ExpenseReport = { ...base, fromDate: '', toDate: '', lines: lines([]) };

async function fillAndLoad(r: ExpenseReport): Promise<ExcelJS.Worksheet> {
  const out = await fillExpenseReport(toArrayBuffer(fixture), r);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(out) as unknown as ArrayBuffer);
  return wb.worksheets[0];
}

it('fills header, rows and cached totals', async () => {
  const ws = await fillAndLoad(report);
  expect(ws.name).toBe('Expense Report Template');
  expect(ws.getCell('A4').value).toBe('TRINH, ROSIE');
  expect((ws.getCell('K4').value as Date).toISOString().slice(0, 10)).toBe('2026-11-05');
  expect((ws.getCell('M4').value as Date).toISOString().slice(0, 10)).toBe('2026-11-18');
  expect(ws.getCell('M4').numFmt).toBe('dd/mm/yy;@');
  expect((ws.getCell('A8').value as Date).toISOString().slice(0, 10)).toBe('2026-11-05');
  expect(ws.getCell('C8').value).toBe('Burnaby V5H 2A9');
  expect(ws.getCell('D8').value).toBe('Golden V0A 1H1');
  expect(ws.getCell('E8').value).toBe(718);
  expect(ws.getCell('F8').value).toBe(0.48);
  expect(ws.getCell('G8').formula).toBe('E8*F8');
  expect((ws.getCell('G8').value as ExcelJS.CellFormulaValue).result).toBeCloseTo(344.64);
  expect(ws.getCell('I9').value).toBe('Parking');
  expect(ws.getCell('K9').value).toBe(1.5);
  expect(ws.getCell('L9').value).toBe(10);
  expect((ws.getCell('N9').value as ExcelJS.CellFormulaValue).result).toBeCloseTo(11.5);
  expect((ws.getCell('N22').value as ExcelJS.CellFormulaValue).result).toBeCloseTo(356.14);
  expect((ws.getCell('K22').value as ExcelJS.CellFormulaValue).result).toBeCloseTo(1.5);
  expect(ws.getCell('A10').value).toBeNull();
});

it('leaves blank inputs blank', async () => {
  const ws = await fillAndLoad(blankReport);
  expect(ws.getCell('C8').value).toBeNull();
  expect(ws.getCell('E8').value).toBeNull();
});
