import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { TEMPLATES } from './definitions';
import { applyLayout } from '../excel/fillTemplate';
import { TimesheetType } from './types';

// Compiled to lib-commonjs/timesheet, so the repo root is two levels up.
const templateDir: string = path.resolve(__dirname, '../../Templates/xlsx');

const collapse = (s: string): string => s.replace(/\s+/g, ' ').trim();

describe('definitions match the real templates', () => {
  const types: TimesheetType[] = ['Admin', 'Instructor', 'Medic'];

  types.forEach((type) => {
    it(type, async () => {
      const def = TEMPLATES[type];
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(path.join(templateDir, def.fileName));
      const ws = wb.getWorksheet(def.sheetName);
      expect(ws).toBeDefined();
      applyLayout(ws!, def);

      // Group headings sit on row 6: merged title over the group's columns, with its fill.
      (def.groups || []).forEach((g) => {
        const top = ws!.getCell(`${g.from}${def.headerRow - 1}`);
        expect(collapse(String(top.master.value))).toBe(g.title);
        expect(ws!.getCell(`${g.to}${def.headerRow - 1}`).master.address).toBe(top.master.address);
        expect((top.fill as { fgColor?: { argb?: string } }).fgColor?.argb).toBe(`FF${g.fill.slice(1).toUpperCase()}`);
      });

      def.columns.forEach((c) => {
        expect(collapse(String(ws!.getCell(`${c.col}${def.headerRow}`).value))).toBe(collapse(c.header));
        if (c.kind === 'computed') {
          expect(ws!.getCell(`${c.col}${def.firstRow}`).formula).toBeTruthy();
        }
      });
    });
  });

  it('medic rate classifications match the template lookup table', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(templateDir, TEMPLATES.Medic.fileName));
    const ws = wb.getWorksheet(TEMPLATES.Medic.sheetName)!;
    const rate = TEMPLATES.Medic.columns.find((c) => c.key === 'dayRate')!;
    // Template dropdown G8 = O6:O9 (=N6:N9 rates), labels in M6:M9.
    [6, 7, 8, 9].forEach((r, i) => {
      expect(rate.options![i]).toBe(ws.getCell(`N${r}`).value);
      expect(rate.optionLabels![i]).toBe(collapse(String(ws.getCell(`M${r}`).value)));
    });
  });

  it('every labelled dropdown has one label per option', () => {
    Object.values(TEMPLATES).forEach((t) =>
      t.columns.filter((c) => c.optionLabels).forEach((c) => expect(c.optionLabels!.length).toBe(c.options!.length))
    );
    expect(TEMPLATES.Instructor.columns.find((c) => c.key === 'dayRate')!.optionLabels).toBeDefined();
  });

  it('instructor and medic have group headings', () => {
    expect(TEMPLATES.Instructor.groups!.map((g) => g.title)).toEqual(['INSTRUCTIONAL FEES', 'ADMINISTRATIVE FEES', 'OTHER FEES']);
    expect(TEMPLATES.Medic.groups!.map((g) => g.title)).toEqual(['MEDIC FEES', 'OTHER FEES']);
  });

  it('medic day rates', () => {
    expect(TEMPLATES.Medic.columns.find((c) => c.key === 'dayRate')!.options).toEqual([250, 350, 450, 550]);
  });

  it('new headings', () => {
    const h = (t: TimesheetType, k: string): string => TEMPLATES[t].columns.find((c) => c.key === k)!.header;
    expect(h('Admin', 'hours')).toBe('TOTAL\nHOURS');
    expect(h('Admin', 'rate')).toBe('PAY\nRATE');
    expect(h('Instructor', 'days')).toBe('DAYS');
    expect(h('Instructor', 'dayRate')).toBe('PAY SCALE LEVEL (PSL)');
  });

  it('per diem is one day or half day', () => ['Instructor', 'Medic'].forEach((t) => {
    const c = TEMPLATES[t as TimesheetType].columns.find((x) => x.key === 'perDiem')!;
    expect(c.options).toEqual([47, 23.5]);
    expect(c.optionLabels).toEqual(['One day ($47.00)', 'Half day ($23.50)']);
  }));
});
