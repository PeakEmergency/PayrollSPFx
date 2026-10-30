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

  it('medic day rates', () => {
    expect(TEMPLATES.Medic.columns.find((c) => c.key === 'dayRate')!.options).toEqual([250, 350, 450, 550]);
  });
});
