import { courseType, fillRate, hasRate, PayRates, toPayRates } from './rates';

const ben: PayRates = { course: { AP: 225, ASU: 225, NUECAFull: 150, NUECABridging: 150, NUECI: 225, NUECB: 225 }, medicPSL: 8 };
const emma: PayRates = { course: {}, admin: 30 };

describe('courseType', () => {
  it.each([
    ['AP-I-0182', 'AP'],
    ['ASU-AB-998', 'ASU'],
    ['NUEC-A-443F', 'NUECAFull'],
    ['NUEC-A-468B', 'NUECABridging'],
    ['NUEC-I-195', 'NUECI'],
    ['NUEC-B-238', 'NUECB'],
    [' nuec-a-443f ', 'NUECAFull']
  ])('%s is %s', (course, type) => expect(courseType(course)).toBe(type));

  it.each(['NUEC-A-443', 'NUEC-443', 'APX-1', '12345', ''])('%s matches no type', (course) => expect(courseType(course)).toBeUndefined());
});

describe('toPayRates', () => {
  it('maps the list row, leaving blanks out', () =>
    expect(toPayRates({ AdminRate: null, MedicPSL: 9, RateAP: 225, RateASU: 225, RateNUECAFull: 150, RateNUECABridging: 150, RateNUECI: 225, RateNUECB: null })).toEqual({
      medicPSL: 9,
      course: { AP: 225, ASU: 225, NUECAFull: 150, NUECABridging: 150, NUECI: 225 }
    }));
});

describe('hasRate', () => {
  it('admin needs an admin rate', () => {
    expect(hasRate('Admin', emma)).toBe(true);
    expect(hasRate('Admin', ben)).toBe(false);
  });
  it('instructor needs a course rate and medic a PSL', () => {
    expect(hasRate('Instructor', ben)).toBe(true);
    expect(hasRate('Medic', ben)).toBe(true);
    expect(hasRate('Instructor', emma)).toBe(false);
    expect(hasRate('Medic', emma)).toBe(false);
    expect(hasRate('Medic', undefined)).toBe(false);
  });
});

describe('fillRate', () => {
  it('fills a blank admin rate once the row has input', () => expect(fillRate('Admin', { date: '2026-11-05' }, emma, 'date').rate).toBe(30));
  it('keeps an admin rate the user changed', () => expect(fillRate('Admin', { date: '2026-11-05', rate: 32 }, emma, 'date').rate).toBe(32));
  it('leaves an empty row alone', () => expect(fillRate('Admin', { rate: null }, emma, 'rate').rate).toBeNull());
  it('fills the medic PSL amount from the level', () => expect(fillRate('Medic', { date: '2026-11-05' }, { course: {}, medicPSL: 9 }, 'date').dayRate).toBe(350));
  it('picks the instructor rate from the course type', () => {
    expect(fillRate('Instructor', { course: 'NUEC-A-443F' }, ben, 'course').dayRate).toBe(150);
    expect(fillRate('Instructor', { course: 'AP-I-0182' }, ben, 'course').dayRate).toBe(225);
  });
  it('changing the course resets the rate, other edits keep it', () => {
    expect(fillRate('Instructor', { course: 'AP-I-0182', dayRate: 150 }, ben, 'course').dayRate).toBe(225);
    expect(fillRate('Instructor', { course: 'AP-I-0182', dayRate: 150 }, ben, 'days').dayRate).toBe(150);
  });
  it('leaves the rate blank for an unknown course', () => expect(fillRate('Instructor', { course: 'XYZ' }, ben, 'course').dayRate).toBeUndefined());
});
