import { validate } from './validate';
import { Line, TimesheetData } from './types';

const base: TimesheetData = { type: 'Medic', name: 'NIKODYM, CAROLYN', from: '2026-10-01', to: '2026-10-15', lines: [] };

it('empty sheet reports header fields and no rows', () => {
  expect(validate({ ...base, name: '', from: null, to: null })).toEqual(
    expect.arrayContaining([
      'Your name is missing.',
      'Pick the pay period From date.',
      'Pick the pay period To date.',
      'Enter at least one row.'
    ])
  );
});

it('To before From', () => {
  expect(validate({ ...base, from: '2026-10-15', to: '2026-10-01' })).toContain("The To date can't be before the From date.");
});

it('invalid time is reported with its row', () => {
  const lines: Line[] = [{}, {}, { date: '2026-10-15', start: 'abc' }];
  expect(validate({ ...base, type: 'Admin', lines })).toContain('Row 3: "abc" is not a valid time.');
});

it('instructor course # must match a course type', () => {
  const lines = [{ date: '2026-10-15', course: 'ABC-1', days: 1, dayRate: 225 }, { date: '2026-10-16', course: 'NUEC-A-443F', days: 1, dayRate: 150 }];
  const errors = validate({ ...base, type: 'Instructor', lines });
  expect(errors).toContain('Row 1: course # "ABC-1" doesn\'t match a course type. Use the course # from the course title, for example AP-I-0182, ASU-AB-998, NUEC-A-443F, NUEC-A-468B, NUEC-I-195 or NUEC-B-238.');
  expect(errors.filter((e) => e.indexOf('Row 2') === 0)).toEqual([]);
});

it('instructor rate needs a course #', () => {
  const lines = [{ date: '2026-10-15', days: 1, dayRate: 225 }];
  expect(validate({ ...base, type: 'Instructor', lines })).toContain('Row 1: course # missing.');
});

it('admin end before start', () => {
  const lines = [{ date: '2026-10-15', start: '17:00', end: '09:00', rate: 20 }];
  expect(validate({ ...base, type: 'Admin', lines })).toContain('Row 1: end time is before start time.');
});

it('row without date', () => {
  expect(validate({ ...base, lines: [{ travel: 100 }] })).toContain('Row 1: date missing.');
});

it('valid admin data', () => {
  const lines = [{ date: '2026-10-15', start: '08:00', end: '16:30', deductions: 0.5, rate: 20 }];
  expect(validate({ ...base, type: 'Admin', lines })).toEqual([]);
});

it('invalid pay rate is reported with its row', () => {
  const lines = [{ date: '2026-10-15', start: '08:00', end: '16:00', rate: '$25' }];
  expect(validate({ ...base, type: 'Admin', lines })).toContain('Row 1: "$25" is not a valid number.');
});

it('medic row with a PSL but no days', () =>
  expect(validate({ ...base, from: '2026-11-05', to: '2026-11-18', lines: [{ date: '2026-11-05', dayRate: 350 }] })).toContain('Row 1: days missing.'));

it('instructor row with a PSL but no days', () =>
  expect(validate({ ...base, type: 'Instructor', from: '2026-11-05', to: '2026-11-18', lines: [{ date: '2026-11-05', dayRate: 280 }] })).toContain('Row 1: days missing.'));
