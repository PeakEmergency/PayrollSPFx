import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { TEMPLATES } from './definitions';
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

      def.columns.forEach((c) => {
        expect(collapse(String(ws!.getCell(`${c.col}${def.headerRow}`).value))).toBe(c.header);
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

  it('medic day rates', () => {
    expect(TEMPLATES.Medic.columns.find((c) => c.key === 'dayRate')!.options).toEqual([250, 350, 450, 550]);
  });
});
