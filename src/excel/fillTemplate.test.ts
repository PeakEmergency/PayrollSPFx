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

const line: Line = { date: '2026-10-15', client: 'acme', days: 1, dayRate: 450, travel: 100, perDiem: 47 };

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

it('medic writes DAYS in F and hides D:E', async () => {
  const ws = await fill('Medic', [line]);
  expect(ws.getCell('F8').value).toBe(1);
  expect(ws.getCell('F7').value).toBe('DAYS');
  expect(ws.getColumn('D').hidden && ws.getColumn('E').hidden).toBe(true);
  expect(ws.getCell('F22').formula).toBe('SUM(F8:F20)');
  expect(ws.getCell('F9').formula).toBeFalsy();
  expect(ws.getCell('H8').formula).toBe('IF(G8="",0,F8*VLOOKUP(G8,$M$6:$N$9,2,FALSE))');
  expect(ws.getCell('H19').formula).toBe('IF(G19="",0,F19*VLOOKUP(G19,$M$6:$N$9,2,FALSE))');
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
    const expected = widths(original);
    if (type === 'Medic') expected[3] = expected[4] = 'hidden'; // layout hides D:E (start/end)
    expect(widths(filled)).toEqual(expected);
  }
}, 60000); // loads and writes 8 workbooks

it('uppercase', async () => {
  const ws = await fill('Medic', [line]);
  expect(ws.getCell('C8').value).toBe('ACME');
});

it('formula cells carry computed results for previewers that do not recalculate', async () => {
  const ws = await fill('Medic', [line]);
  const result = (a: string): unknown => (ws.getCell(a).value as { result?: unknown }).result;
  expect(result('H8')).toBe(450);
  expect(result('K8')).toBe(147);
  expect(result('L8')).toBe(597);
  expect(result('F22')).toBe(1);
  expect(result('J22')).toBe(47);
  expect(result('L22')).toBe(597);
});

it('applyLayout writes the new headings', async () => {
  const admin = await fill('Admin', [{}]);
  expect(admin.getCell('F6').value).toBe('TOTAL\nHOURS');
  expect(admin.getCell('G6').value).toBe('PAY\nRATE');
  expect(admin.getCell('F6').alignment?.wrapText).toBe(true);
  const instr = await fill('Instructor', [{}]);
  expect(instr.getCell('D7').value).toBe('DAYS');
  expect(instr.getCell('E7').value).toBe('PAY SCALE LEVEL (PSL)');
});

it('instructor PSL cell shows the level and the total looks it up', async () => {
  const ws = await fill('Instructor', [{ date: '2026-10-15', course: 'A1', days: 1, dayRate: 225 }]);
  expect(ws.getCell('E8').value).toBe('Instructor – Level 1 (PSL 3)');
  expect(ws.getCell('F8').formula).toBe('IF(E8="",0,D8*VLOOKUP(E8,$N$8:$O$14,2,FALSE))');
  expect(ws.getCell('N10').value).toBe('Instructor – Level 1 (PSL 3)');
  expect(ws.getCell('O10').value).toBe(225);
  expect((ws.getCell('F8').value as { result?: number }).result).toBe(225);
});

it('medic PSL cell shows the level', async () => {
  const ws = await fill('Medic', [{ date: '2026-10-15', client: 'X', days: 0.5, dayRate: 350 }]);
  expect(ws.getCell('G8').value).toBe('Medic – Level PCP(PSL 9)');
  expect(ws.getCell('H8').formula).toBe('IF(G8="",0,F8*VLOOKUP(G8,$M$6:$N$9,2,FALSE))');
  expect((ws.getCell('H8').value as { result?: number }).result).toBe(175);
});

it('blank grid rows use the guarded lookup and cache 0, not an error', async () => {
  const i = await fill('Instructor', [{ date: '2026-10-15', course: 'A1', days: 1, dayRate: 225 }]);
  expect(i.getCell('F9').formula).toBe('IF(E9="",0,D9*VLOOKUP(E9,$N$8:$O$14,2,FALSE))');
  expect((i.getCell('F9').value as { result?: number }).result || 0).toBe(0);
  expect((i as unknown as { dataValidations: { model: Record<string, unknown> } }).dataValidations.model.E8).toBeUndefined();
  const m = await fill('Medic', [{ date: '2026-10-15', client: 'X', days: 0.5, dayRate: 350 }]);
  expect(m.getCell('H9').formula).toBe('IF(G9="",0,F9*VLOOKUP(G9,$M$6:$N$9,2,FALSE))');
  expect((m.getCell('H9').value as { result?: number }).result || 0).toBe(0);
  expect((m as unknown as { dataValidations: { model: Record<string, unknown> } }).dataValidations.model.G8).toBeUndefined();
});
