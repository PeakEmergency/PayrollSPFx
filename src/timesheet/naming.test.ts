import { buildFileName, buildSubject, formatEmployeeName } from './naming';
import { TimesheetData } from './types';

it('formatEmployeeName', () => {
  expect(formatEmployeeName('Carolyn', 'Nikodym', 'Carolyn Nikodym')).toBe('NIKODYM, CAROLYN');
  expect(formatEmployeeName(null, null, "Katie O'Connor")).toBe("O'CONNOR, KATIE");
  expect(formatEmployeeName(null, null, 'Madonna')).toBe('MADONNA');
});

const data: TimesheetData = { type: 'Medic', name: 'NIKODYM, CAROLYN', from: '2026-10-01', to: '2026-10-15', lines: [] };

it('buildFileName', () => {
  expect(buildFileName(data)).toBe('Medic Timesheet - NIKODYM, CAROLYN - 2026-10-01_2026-10-15.xlsx');
});

it('buildSubject', () => {
  expect(buildSubject(data)).toBe('Medic Timesheet from NIKODYM, CAROLYN – 01/10/26 to 15/10/26');
});
