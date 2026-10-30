import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { fillTemplate } from './fillTemplate';
import { TEMPLATES } from '../timesheet/definitions';
import { Line, TimesheetData, TimesheetType } from '../timesheet/types';

// Compiled to lib-commonjs/excel, so the repo root is two levels up.
const fixture = (type: TimesheetType): Buffer =>
  fs.readFileSync(path.resolve(__dirname, '../../Templates/xlsx', TEMPLATES[type].fileName));

// Copy into an ArrayBuffer of Jest's sandbox realm; Buffer's own fails jszip's instanceof check.
const toArrayBuffer = (b: Buffer): ArrayBuffer => new Uint8Array(b).buffer;

async function load(buf: ArrayBuffer | Buffer, type: TimesheetType): Promise<ExcelJS.Worksheet> {
  const wb = new ExcelJS.Workbook();
  // Jest's sandbox fails jszip's `instanceof ArrayBuffer` check, so hand it a Buffer.
  await wb.xlsx.load((Buffer.isBuffer(buf) ? buf : Buffer.from(buf)) as unknown as ArrayBuffer);
  return wb.getWorksheet(TEMPLATES[type].sheetName)!;
}

async function fill(type: TimesheetType, lines: Line[]): Promise<ExcelJS.Worksheet> {
  const data: TimesheetData = { type, name: 'NIKODYM, CAROLYN', from: '2026-10-01', to: '2026-10-15', lines };
  return load(await fillTemplate(toArrayBuffer(fixture(type)), data), type);
}

const line: Line = { date: '2026-10-15', client: 'acme', start: '06:00', end: '18:00', dayRate: 450, travel: 100, perDiem: 47 };

it('fillTemplate writes header', async () => {
  const ws = await fill('Medic', [line]);
  expect(ws.getCell('A4').value).toBe('NIKODYM, CAROLYN');
  expect(ws.getCell('H4').value).toBeInstanceOf(Date);
  expect(ws.getCell('E4').isMerged && ws.getCell('E4').master.address).toBe('A4');
});

it('fillTemplate writes dates as the same calendar day', async () => {
  const ws = await fill('Medic', [line]);
  const original = await load(fixture('Medic'), 'Medic');
  expect((ws.getCell('A8').value as Date).toISOString().slice(0, 10)).toBe('2026-10-15');
  expect(ws.getCell('A8').numFmt).toBe(original.getCell('A8').numFmt);
});

it('fillTemplate keeps row positions', async () => {
  const ws = await fill('Medic', [line, {}, {}, {}, { ...line, client: 'beta' }]);
  expect(ws.getCell('C8').value).toBe('ACME');
  expect(ws.getCell('C12').value).toBe('BETA');
  ['C9', 'C10', 'C11'].forEach((a) => expect(ws.getCell(a).value).toBeNull());
});

it('medic formula fixes', async () => {
  const ws = await fill('Medic', [line]);
  expect(ws.getCell('H8').formula).toBe('G8');
  expect(ws.getCell('H19').formula).toBe('G19');
  expect(ws.getCell('F8').formula).toBe('MOD(E8-D8,1)*24');
  expect(ws.getCell('F20').formula).toBe('MOD(E20-D20,1)*24');
  expect(ws.getCell('K20').formula).toBe('SUM(I20:J20)');
  expect(ws.getCell('J22').formula).toBe('SUM(J8:J20)');
});

it('other templates untouched', async () => {
  const instructor = await fill('Instructor', []);
  expect(instructor.getCell('M8').formula).toBe('SUM(F8,I8,L8)');
  const admin = await fill('Admin', []);
  expect(admin.getCell('F7').formula).toBe('((D7-C7)*24-E7)');

  for (const type of ['Admin', 'Instructor', 'Medic'] as TimesheetType[]) {
    const filled = await fill(type, []);
    const original = await load(fixture(type), type);
    // Hidden columns compare as 'hidden': ExcelJS round-trips some hidden width-0 ranges as width 9.
    const widths = (ws: ExcelJS.Worksheet): (number | string | undefined)[] => ws.columns.map((c) => (c.hidden ? 'hidden' : c.width));
    expect(widths(filled)).toEqual(widths(original));
  }
});

it('times written as day fractions', async () => {
  const ws = await fill('Medic', [line]);
  // ExcelJS reads time-formatted numbers back as Dates on the 1899-12-30 epoch.
  const serial = ((ws.getCell('D8').value as Date).getTime() - Date.UTC(1899, 11, 30)) / 86400000;
  expect(serial).toBe(0.25);
});

it('uppercase', async () => {
  const ws = await fill('Medic', [line]);
  expect(ws.getCell('C8').value).toBe('ACME');
});
