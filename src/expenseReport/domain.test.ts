import { calculateReport, formatCurrency, formatQuantity } from './domain/calculations';
import { createReport } from './domain/report';
import { validateReport } from './domain/report';
import { calendarStartDate, formatDateInput, fromCalendarDate, parseDateInput, toCalendarDate } from './domain/dates';
import type { ExpenseReport, Instructor } from './domain/models';

describe('blank expense report', () => {
  it('creates thirteen independently editable slots and snapshots the configured rate', () => {
    const first = createReport('0.51');
    const second = createReport();
    expect(first.schemaVersion).toBe(1);
    expect(first.lines).toHaveLength(13);
    expect(new Set(first.lines.map(line => line.id)).size).toBe(13);
    expect(first.clientId).not.toBe(second.clientId);
    expect(first.lines.every(line => line.rate === '0.51')).toBe(true);
    expect(second.lines[0].rate).toBe('0.48');
    first.lines[0].description = 'Edited';
    expect(first.lines[1].description).toBe('');
    expect(second.lines[0].description).toBe('');
    expect(first.instructorId).toBe('');
    expect(first.fromDate).toBe('');
    expect(first.id).toBeUndefined();
  });
});

describe('explicit date-only helpers', () => {
  it.each([
    ['29/02/24', '2024-02-29'],
    ['31/12/29', '2029-12-31'],
    ['01/01/30', '1930-01-01'],
    ['01/01/00', '2000-01-01'],
    ['31/12/99', '1999-12-31'],
    ['01/01/2030', '2030-01-01'],
    ['29/02/2000', '2000-02-29'],
    [' 02/03/24 ', '2024-03-02']
  ])('parses %s explicitly as %s using the fixed 2029 cutoff', (input, expected) => {
    expect(parseDateInput(input)).toBe(expected);
  });

  it.each(['', 'DD/MM/YY', '29/02/23', '29/02/1900', '31/04/24', '00/01/24', '01/13/24', '2/3/24', '02/03/2', '02/03/0000', '2024-03-02', '03-02-24', '02/30/24'])('rejects malformed or impossible date %s', raw => {
    expect(parseDateInput(raw)).toBeUndefined();
    expect(toCalendarDate(raw)).toBeUndefined();
  });

  it('formats ISO date-only values and preserves centuries outside the two-digit-year window', () => {
    expect(formatDateInput('2024-03-02')).toBe('02/03/24');
    expect(formatDateInput('1930-01-01')).toBe('01/01/30');
    expect(formatDateInput('2030-01-01')).toBe('01/01/2030');
    expect(formatDateInput('1900-01-01')).toBe('01/01/1900');
    expect(formatDateInput('invalid')).toBe('');
    expect(formatDateInput('2023-02-29')).toBe('');
    expect(fromCalendarDate(new Date(10001, 0, 1))).toBe('');
  });

  it.each(['America/Chicago', 'Pacific/Kiritimati', 'America/Los_Angeles'])('preserves calendar days at timezone and DST boundaries in %s', timezone => {
    const original = process.env.TZ;
    process.env.TZ = timezone;
    try {
      const spring = toCalendarDate('10/03/24');
      const autumn = toCalendarDate('03/11/24');
      expect(spring?.getFullYear()).toBe(2024);
      expect(spring?.getMonth()).toBe(2);
      expect(spring?.getDate()).toBe(10);
      expect(fromCalendarDate(spring!)).toBe('10/03/24');
      expect(fromCalendarDate(autumn!)).toBe('03/11/24');
      expect(parseDateInput(fromCalendarDate(toCalendarDate('01/01/2030')!))).toBe('2030-01-01');
    } finally {
      if (original === undefined) delete process.env.TZ;
      else process.env.TZ = original;
    }
  });

  it('opens on the current field, then the other date, then the current month', () => {
    expect(fromCalendarDate(calendarStartDate('15/04/24', '01/02/24'))).toBe('15/04/24');
    expect(fromCalendarDate(calendarStartDate('', '01/02/24'))).toBe('01/02/24');
    const now = new Date();
    const fallback = calendarStartDate('');
    expect(fallback.getFullYear()).toBe(now.getFullYear());
    expect(fallback.getMonth()).toBe(now.getMonth());
    expect(fromCalendarDate(new Date(NaN))).toBe('');
  });
});

const roster: Instructor[] = [{ id: 'instructor-1', name: 'Ada Instructor' }, { id: 'instructor-2', name: 'Bea Instructor' }];
function validHeader(): ExpenseReport {
  const report = createReport();
  report.instructorId = 'instructor-1';
  report.instructorName = 'Ada Instructor';
  report.fromDate = '01/03/24';
  report.toDate = '31/03/24';
  return report;
}

describe('save validation', () => {
  it('allows valid roster and pay-period headers with every expense row empty and no receipts', () => {
    expect(validateReport(validHeader(), roster)).toEqual([]);
  });

  it('collects missing header fields without requiring any expense detail', () => {
    const issues = validateReport(createReport(), roster);
    expect(issues.map(issue => issue.field)).toEqual(['instructorId', 'fromDate', 'toDate']);
    expect(issues.every(issue => issue.message.length > 0)).toBe(true);
  });

  it.each([
    ['', ''],
    ['instructor-1', 'PICK FROM DROPDOWN'],
    ['unknown', 'Ada Instructor']
  ])('rejects instructor selection outside the actual roster: %s / %s', (id, name) => {
    const report = validHeader(); report.instructorId = id; report.instructorName = name;
    expect(validateReport(report, roster).some(issue => issue.field === 'instructorId')).toBe(true);
  });

  it('preserves historical display-name snapshots when a valid roster identity is renamed or inactive', () => {
    const report = validHeader(); report.instructorName = 'Ada Previous Name';
    expect(validateReport(report, [{ id: 'instructor-1', name: 'Ada Renamed', active: false }])).toEqual([]);
    report.instructorName = '   ';
    expect(validateReport(report, roster).some(issue => issue.field === 'instructorId')).toBe(true);
  });

  it('allows same-day periods and rejects reversed or invalid pay dates', () => {
    const report = validHeader();
    report.toDate = report.fromDate;
    expect(validateReport(report, roster)).toEqual([]);
    report.toDate = '29/02/24';
    expect(validateReport(report, roster).some(issue => issue.field === 'toDate')).toBe(true);
    report.fromDate = '31/02/24';
    expect(validateReport(report, roster).some(issue => issue.field === 'fromDate')).toBe(true);
  });

  it('requires thirteen slots and unique nonempty line identities', () => {
    const report = validHeader(); report.lines.pop();
    expect(validateReport(report, roster).some(issue => issue.field === 'lines')).toBe(true);
    const duplicated = validHeader(); duplicated.lines[1].id = duplicated.lines[0].id;
    expect(validateReport(duplicated, roster).some(issue => issue.field === 'lines.1.id')).toBe(true);
  });

  it('rejects malformed populated detail dates and numeric fields while retaining raw input', () => {
    const report = validHeader();
    report.lines[0].date = '31/04/24';
    report.lines[0].kilometres = 'ten';
    report.lines[1].gst = 'NaN';
    report.lines[2].cost = 'Infinity';
    report.lines[3].rate = 'bad';
    expect(validateReport(report, roster).map(issue => issue.field)).toEqual([
      'lines.0.date', 'lines.0.kilometres', 'lines.1.gst', 'lines.2.cost', 'lines.3.rate'
    ]);
    expect(report.lines[0].kilometres).toBe('ten');
  });

  it('allows negative adjustments, partial rows, fractional distance, and dates outside the pay period', () => {
    const report = validHeader();
    report.lines[0].date = '01/01/24';
    report.lines[0].kilometres = '-1.25';
    report.lines[0].cost = '-2.50';
    report.lines[1].gst = '5';
    report.lines[2].description = '123';
    expect(validateReport(report, roster)).toEqual([]);
  });
});

describe('workbook calculations', () => {
  it('calculates mileage, other fees, and their combined amount without applying extra GST', () => {
    const report = createReport();
    report.lines[0].kilometres = '100';
    report.lines[1].cost = '100';
    report.lines[1].gst = '5';
    report.lines[2].kilometres = '100';
    report.lines[2].cost = '100';
    report.lines[2].gst = '5';
    const totals = calculateReport(report);
    expect(totals.lines[0]).toEqual({ mileage: '48', amount: '48' });
    expect(totals.lines[1]).toEqual({ mileage: '0', amount: '105' });
    expect(totals.lines[2]).toEqual({ mileage: '48', amount: '153' });
    expect(totals.amount).toBe('306');
  });

  it('sums all thirteen rows exactly once, including first, middle, and last slots', () => {
    const report = createReport();
    report.lines.forEach(line => { line.kilometres = '1'; line.cost = '2'; line.gst = '0.1'; });
    report.lines[0].kilometres = '2';
    report.lines[6].kilometres = '3';
    report.lines[12].kilometres = '4';
    const totals = calculateReport(report);
    expect(totals.lines).toHaveLength(13);
    expect(totals.kilometres).toBe('19');
    expect(totals.mileage).toBe('9.12');
    expect(totals.cost).toBe('26');
    expect(totals.gst).toBe('1.3');
    expect(totals.amount).toBe('36.42');
  });

  it('retains unrounded decimal calculations before formatting totals', () => {
    const report = createReport();
    report.lines.slice(0, 3).forEach(line => { line.kilometres = '1.01'; });
    const totals = calculateReport(report);
    expect(totals.lines[0].mileage).toBe('0.4848');
    expect(totals.mileage).toBe('1.4544');
    expect(formatCurrency(totals.mileage)).toBe('$1.45');
    expect(report.lines[0].kilometres).toBe('1.01');
  });

  it('preserves significant input digits instead of silently rounding at library defaults', () => {
    const report = createReport();
    report.lines[0].kilometres = '123456789012345678901234567890.123456789';
    report.lines[0].rate = '0.48';
    expect(calculateReport(report).mileage).toBe('59259258725925925872592592587.25925925872');
  });

  it('keeps blank input distinct from zero while calculating an empty report as zero', () => {
    const report = createReport();
    report.lines[1].kilometres = '0';
    const totals = calculateReport(report);
    expect(totals.amount).toBe('0');
    expect(totals.kilometres).toBe('0');
    expect(formatCurrency(totals.amount)).toBe('$0.00');
    expect(report.lines[0].kilometres).toBe('');
    expect(report.lines[1].kilometres).toBe('0');
  });

  it('treats numeric-looking descriptions and city values as text only', () => {
    const report = createReport();
    report.lines[0].description = '100';
    report.lines[0].startCity = '500';
    report.lines[0].endCity = '200';
    report.lines[0].gst = '5';
    expect(calculateReport(report).lines[0].amount).toBe('5');
  });

  it('allows negative and decimal inputs without truncating distance or changing snapshots', () => {
    const report = createReport('0.5');
    report.lines[0].kilometres = '-1.25';
    report.lines[0].gst = '-.25';
    report.lines[0].cost = '+2.00';
    const totals = calculateReport(report);
    expect(totals.kilometres).toBe('-1.25');
    expect(totals.mileage).toBe('-0.625');
    expect(totals.amount).toBe('1.125');
    expect(formatQuantity(totals.kilometres)).toBe('-1.25');
    expect(formatCurrency(totals.amount)).toBe('$1.13');
    expect(report.lines[0].rate).toBe('0.5');
  });

  it.each(['no', 'NaN', 'Infinity', '1e3', '1,000', '$5', '--1'])('keeps malformed numeric %s unavailable instead of zero', raw => {
    const report = createReport();
    report.lines[0].kilometres = raw;
    report.lines[0].gst = '5';
    const totals = calculateReport(report);
    expect(totals.lines[0]).toEqual({ mileage: null, amount: null });
    expect(totals.kilometres).toBeNull();
    expect(totals.mileage).toBeNull();
    expect(totals.amount).toBeNull();
    expect(totals.gst).toBe('5');
    expect(totals.cost).toBe('0');
    expect(formatCurrency(totals.amount)).toBe('—');
    expect(report.lines[0].kilometres).toBe(raw);
  });

  it('only invalidates totals that depend on the malformed numeric field', () => {
    const report = createReport();
    report.lines[0].kilometres = '100';
    report.lines[0].gst = 'bad';
    const totals = calculateReport(report);
    expect(totals.lines[0]).toEqual({ mileage: '48', amount: null });
    expect(totals.kilometres).toBe('100');
    expect(totals.mileage).toBe('48');
    expect(totals.gst).toBeNull();
    expect(totals.cost).toBe('0');
    expect(totals.amount).toBeNull();
  });

  it('requires a captured numeric rate even when a line is otherwise empty', () => {
    const report = createReport();
    report.lines[0].rate = '';
    expect(calculateReport(report).lines[0].mileage).toBeNull();
    expect(validateReport(report, roster).some(issue => issue.field === 'lines.0.rate')).toBe(true);
  });

  it('formats quantities without hiding fractional kilometres and currency with grouping', () => {
    expect(formatQuantity('1234.56789')).toBe('1,234.56789');
    expect(formatQuantity(null)).toBe('—');
    expect(formatCurrency('1234.5')).toBe('$1,234.50');
    expect(formatCurrency('-1234.5')).toBe('-$1,234.50');
    expect(formatCurrency('+2.00')).toBe('$2.00');
    expect(formatQuantity('+1.25')).toBe('1.25');
    expect(formatCurrency('bad')).toBe('—');
  });
});
