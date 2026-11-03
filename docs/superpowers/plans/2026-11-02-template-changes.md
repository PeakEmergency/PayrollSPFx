# Template Changes (jeff's review) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Pay period dropdown backed by a SharePoint list on every form, the renamed and reshaped Instructor/Medic/Admin columns (on screen and in the emailed Excel), an Expense Report tied to the signed-in user, and timesheets restyled to look like the Expense Report.

**Architecture:**
- `src/timesheet/definitions.ts` stays the single source of truth for column headings.
- `fillTemplate.ts` gains `applyLayout()`, which rewrites the template's headings, formulas and lookup tables when the file is filled. The `.xlsx` files in the library do not need re-uploading, and Excel COM never touches the Medic file.
- Pay periods come from a new list through one pure module (`payPeriods.ts`) and one shared Fluent UI picker used by both forms.

**Tech Stack:** SPFx 1.24, React 18, Fluent UI 8, Jspreadsheet CE 5, ExcelJS 4, Jest via Heft (Node environment, runs on `lib-commonjs`).

**Spec:** `docs/superpowers/specs/2026-11-02-template-changes-design.md`

## Global Constraints

- Build target is ES5. Custom `Error` subclasses need `Object.setPrototypeOf(this, X.prototype)`.
- The only Graph permission is `User.Read`. Read lists with `SPHttpClient`.
- `npm test` runs `heft test --clean`. Tests live next to the source as `*.test.ts`. If `EPERM` appears on `jest-output`, run `rm -rf jest-output` and retry.
- Name always comes from the signed-in user. No name dropdown anywhere.
- Pay period list title: `Pay Period Schedule`, on the timesheet web (`pageContext.web.absoluteUrl`).
- Period label format: `2026 – PP 23 (Nov 5 – Nov 18)`, with an en dash (U+2013).
- Per diem options: `47` labelled `One day ($47.00)`, and `23.5` labelled `Half day ($23.50)`. Instructor and Medic only.
- Medic DAYS options: `0.5`, `1`.
- Headings (exact; `\n` means a line break):
  - Admin F: `TOTAL\nHOURS`
  - Admin G: `PAY\nRATE`
  - Instructor D: `DAYS`
  - Instructor E: `PAY SCALE LEVEL (PSL)`
  - Medic F: `DAYS`
  - Medic G: `PSL`
- The PSL cell shows the level label, on screen and in Excel. The dollar amount is used for totals only.
- No preselected pay period. Only periods with From ≤ today are listed, newest first. From and To can't be edited.

## Review Focus

1. **Pay period list empty or unreachable.** The form shows `Pay periods couldn't be loaded: <reason>` or `No pay periods are set up yet. Contact accounts.`, and Submit/Save stays disabled. Test in Task 2 (`selectablePeriods([])`), with a manual check in Task 4.
2. **Period crossing a year.** 2027 PP 1 (2026-12-31 to 2027-01-13) is labelled `2027 – PP 1 (Dec 31 – Jan 13)` and sorts after 2026 PP 26. Test in Task 2.
3. **SharePoint date-only values come back as UTC timestamps** (`2026-11-05T07:00:00Z`, or `…T00:00:00Z`). Both must read as `2026-11-05`. Test in Task 3.
4. **A Medic or Instructor row with a PSL but no DAYS** must not silently total $0. Validation says `Row N: days missing.` Test in Task 6.
5. **A reopened expense report whose dates match no listed period** (for example, the period was removed from the list). The dates still show and the dropdown shows no selection. Test in Task 8 (`findPeriod`).

---

### Task 1: Commit the baseline

The Expense Report fold-in and the spec are uncommitted. Commit them first, so every later diff is reviewable on its own.

- [x] **Step 1:** `npm test`. Expected: all suites PASS.
- [x] **Step 2:** Commit.

```bash
git add -A
git commit -m "feat: add Expense Report button that shows the expense report in place of the timesheet"
```

### Task 2: Pay period logic

**Files:**
- Create: `src/timesheet/payPeriods.ts`
- Test: `src/timesheet/payPeriods.test.ts`

**Interfaces:**
- Produces:
  - `interface PayPeriod { year: number; number: number; from: string; to: string }`, with dates in `yyyy-mm-dd` format
  - `selectablePeriods(all: PayPeriod[], todayIso: string): PayPeriod[]`, which keeps periods with `from <= todayIso` and sorts by `from`, newest first
  - `periodLabel(p: PayPeriod): string`
  - `findPeriod(all: PayPeriod[], fromIso: string, toIso: string): PayPeriod | undefined`

- [x] **Step 1: Write the failing tests**

```ts
const pp = (year: number, number: number, from: string, to: string): PayPeriod => ({ year, number, from, to });
const p23 = pp(2026, 23, '2026-11-05', '2026-11-18');
const p24 = pp(2026, 24, '2026-11-19', '2026-12-02');
const p27_1 = pp(2027, 1, '2026-12-31', '2027-01-13');

it('labels a period', () => expect(periodLabel(p23)).toBe('2026 – PP 23 (Nov 5 – Nov 18)'));
it('labels a period that crosses the year', () => expect(periodLabel(p27_1)).toBe('2027 – PP 1 (Dec 31 – Jan 13)'));
it('hides future periods and sorts newest first', () =>
  expect(selectablePeriods([p23, p27_1, p24], '2026-11-20')).toEqual([p24, p23]));
it('includes a period starting today', () => expect(selectablePeriods([p24], '2026-11-19')).toEqual([p24]));
it('orders across the year boundary', () =>
  expect(selectablePeriods([p27_1, p24], '2027-01-02').map((p) => p.year)).toEqual([2027, 2026]));
it('empty list stays empty', () => expect(selectablePeriods([], '2026-11-20')).toEqual([]));
it('finds by dates', () => expect(findPeriod([p23, p24], '2026-11-05', '2026-11-18')).toBe(p23));
it('no match is undefined', () => expect(findPeriod([p23], '2026-11-06', '2026-11-18')).toBeUndefined());
```

- [x] **Step 2:** Run `npm test`. Expected: FAIL, because `payPeriods` does not exist.
- [x] **Step 3:** Implement the module. Month names are the short English names (`Jan`…`Dec`). Parse dates with string splitting, not `new Date(iso)`, to avoid time zone shifts.
- [x] **Step 4:** Run `npm test`. Expected: PASS.
- [x] **Step 5:** Commit: `feat: pay period labels and selection`.

### Task 3: Pay Period Schedule list (provisioning and reading)

**Files:**
- Modify: `provisioning/provision-timesheets.js` (new section after the template library)
- Modify: `src/services/sharepoint.ts`
- Test: `src/services/sharepoint.test.ts` (create)
- Modify: `README.md` (setup steps)

**Interfaces:**
- Consumes: `PayPeriod` (Task 2)
- Produces:
  - `toPayPeriod(row: { Year: number; PeriodNumber: number; FromDate: string; ToDate: string }): PayPeriod`
  - `class SpPayPeriodSource { constructor(client: SPHttpClient, webUrl: string, listTitle: string); getPayPeriods(): Promise<PayPeriod[]> }`
  - `const PAY_PERIOD_LIST = 'Pay Period Schedule'`, exported from `sharepoint.ts`

- [x] **Step 1: Write the failing tests**

```ts
it('reads a date-only value returned at local midnight in UTC', () =>
  expect(toPayPeriod({ Year: 2026, PeriodNumber: 23, FromDate: '2026-11-05T07:00:00Z', ToDate: '2026-11-18T07:00:00Z' }))
    .toEqual({ year: 2026, number: 23, from: '2026-11-05', to: '2026-11-18' }));
it('reads a value stored at UTC midnight', () =>
  expect(toPayPeriod({ Year: 2026, PeriodNumber: 1, FromDate: '2026-01-01T00:00:00Z', ToDate: '2026-01-14T00:00:00Z' }).from).toBe('2026-01-01'));
```

- [x] **Step 2:** Run `npm test`. Expected: FAIL.
- [x] **Step 3:** Implement `toPayPeriod`, which takes the first 10 characters of the date. Peak's sites are in Canadian time zones, west of UTC.
- [x] **Step 4:** Implement `getPayPeriods`, which runs `GET items?$select=Year,PeriodNumber,FromDate,ToDate&$top=5000` through the existing `check()` with step name `Loading pay periods`.
- [x] **Step 5:** Extend the provisioning script. It creates list `Pay Period Schedule` (BaseTemplate 100) if missing, with fields:
  - `Year` (Number, required)
  - `PeriodNumber` (Number, required)
  - `FromDate`, `ToDate`, `PayrollRun` (DateTime, `Format="DateOnly"`)

  Seed 52 rows: period k (0–51) has From = 2026-01-01 + 14k days, To = From + 13, PayrollRun = To + 1, Year = 2026 for k < 26 and otherwise 2027, PeriodNumber = (k % 26) + 1, Title = `YYYY-PP`. Skip a row if one with the same Year and PeriodNumber exists. Write dates as `YYYY-MM-DDT12:00:00Z`.
- [x] **Step 6:** Check the seed against both PDFs in `docs/`. The 2026 PP 23 row must be Nov 5–Nov 18 (run Nov 19), 2027 PP 1 must be 2026-12-31 to 2027-01-13, and 2027 PP 26 must be Dec 16–Dec 29. Do this by running the date loop in Node and comparing the three rows.
- [x] **Step 7:** Add to README setup: "Rerun `provision-timesheets.js`; it creates and seeds Pay Period Schedule. Add future years as rows in that list."
- [x] **Step 8:** Run `npm test`, expect PASS. Commit: `feat: Pay Period Schedule list`.

### Task 4: Shared pay period picker on the timesheets

**Files:**
- Create: `src/webparts/timesheet/components/PayPeriodPicker.tsx`
- Modify: `src/webparts/timesheet/components/TimesheetForm.tsx`
- Modify: `src/webparts/timesheet/TimesheetWebPart.ts`

**Interfaces:**
- Consumes: `selectablePeriods`, `periodLabel`, `findPeriod` (Task 2); `SpPayPeriodSource`, `PAY_PERIOD_LIST` (Task 3)
- Produces:
  - `PayPeriodPicker(props: { id: string; periods: PayPeriod[] | undefined; error?: string; from: string; to: string; disabled?: boolean; onChange(p: PayPeriod): void })`, where `from` and `to` are ISO dates or `''`, and `periods === undefined` means still loading
  - `ITimesheetFormProps.getPayPeriods(): Promise<PayPeriod[]>`

- [x] **Step 1:** Build `PayPeriodPicker`:
  - A Fluent `Dropdown` labelled `Pay period`, with placeholder `Pick a pay period`.
  - Options come from `selectablePeriods(periods, today)`, keyed `${year}-${number}`.
  - The selected key is `findPeriod(...)` of `from` and `to`.
  - Below it, read-only `From:` and `To:` values in `DD/MM/YY` format. Render them as text, not inputs.
  - When loading fails, show the `error` prop text. When the list is empty, show `No pay periods are set up yet. Contact accounts.`
- [x] **Step 2:** In `TimesheetForm`, replace the two `<input type="date">` elements with `PayPeriodPicker`. Load periods once in an effect, and on failure set the error `Pay periods couldn't be loaded: <message>`. `onChange` sets `from` and `to`. Submit stays disabled while `!from || !to`, as today.
- [x] **Step 3:** In the web part, pass `getPayPeriods: () => new SpPayPeriodSource(this.context.spHttpClient, web.absoluteUrl, PAY_PERIOD_LIST).getPayPeriods()`.
- [x] **Step 4:** Run `npm run build`. Expected: package produced, no errors.
- [ ] **Step 5: Manual check (hosted workbench, `npm start`):**
  - Dropdown lists PP 23 and earlier, newest first.
  - Picking PP 23 shows From 05/11/26 and To 18/11/26, and neither can be edited.
  - Temporarily renaming the list shows the load error, and Submit stays disabled.
- [x] **Step 6:** Commit: `feat: pay period dropdown on timesheets`.

### Task 5: Headings, per diem and Excel layout patching

**Files:**
- Modify: `src/timesheet/definitions.ts`
- Modify: `src/timesheet/definitions.test.ts`
- Modify: `src/excel/fillTemplate.ts`
- Test: `src/excel/fillTemplate.test.ts`

**Interfaces:**
- Produces: `applyLayout(ws: ExcelJS.Worksheet, def: TemplateDef): void`, exported from `fillTemplate.ts` and called by `fillTemplate` right after loading. For every column in `def.columns`, it writes `def.columns[i].header` into `${col}${headerRow}`. It keeps each cell's style and turns on wrap text when the header contains `\n`.

- [x] **Step 1: Update the definitions tests to the new headings**
  - Change the per-type test so that it loads the template, calls `applyLayout(ws, def)`, then runs the existing header, group and formula checks.
  - Add:

```ts
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
```

- [x] **Step 2:** Add a fillTemplate test: fill Admin and read back `F6` = `TOTAL\nHOURS` and `G6` = `PAY\nRATE`; fill Instructor and read back `D7` = `DAYS` and `E7` = `PAY SCALE LEVEL (PSL)`.
- [x] **Step 3:** Run `npm test`. Expected: FAIL.
- [x] **Step 4:** Change the headings and per diem options and labels in `definitions.ts`. Add `applyLayout` and call it in `fillTemplate`.
- [x] **Step 5:** Run `npm test`. Expected: PASS.
- [x] **Step 6:** Commit: `feat: renamed headings and half-day per diem`.

### Task 6: Medic DAYS replaces start and end times

**Files:**
- Modify: `src/timesheet/definitions.ts` (Medic)
- Modify: `src/timesheet/calc.ts`, `src/timesheet/validate.ts`
- Modify: `src/excel/fillTemplate.ts` (replace `fixMedicFormulas`)
- Tests: `calc.test.ts`, `validate.test.ts`, `fillTemplate.test.ts`, `definitions.test.ts`

**Interfaces:**
- Medic columns become:
  - `date` A
  - `client` C
  - `days` F (dropdown `[0.5, 1]`, header `DAYS`)
  - `dayRate` G (header `PSL`; options and labels unchanged)
  - `medicTotal` H (computed)
  - `travel` I, `perDiem` J
  - `otherTotal` K, `amount` L
- Medic totals: `{ key: 'days', cell: 'F22' }`, travel I22, perDiem J22, amount L22.
- In the Excel file, columns D and E are hidden, because the `From:` label is merged over F4:G4 and must stay visible. `H${r}` gets formula `F${r}*G${r}`, which Task 7 replaces with the lookup.

- [x] **Step 1: Write the failing tests**

```ts
// calc.test.ts
it('medic total is days × rate', () =>
  expect(computeLine('Medic', { date: '2026-11-05', days: 0.5, dayRate: 350 }).medicTotal).toBe(175));
// validate.test.ts
it('medic row with a PSL but no days', () =>
  expect(validate({ type: 'Medic', name: 'A', from: '2026-11-05', to: '2026-11-18',
    lines: [{ date: '2026-11-05', dayRate: 350 }] })).toContain('Row 1: days missing.'));
it('instructor row with a PSL but no days', () =>
  expect(validate({ type: 'Instructor', name: 'A', from: '2026-11-05', to: '2026-11-18',
    lines: [{ date: '2026-11-05', dayRate: 280 }] })).toContain('Row 1: days missing.'));
// fillTemplate.test.ts
it('medic writes DAYS in F and hides D:E', async () => {
  /* fill Medic with one line { date, client: 'X', days: 1, dayRate: 350 }; then: */
  expect(ws.getCell('F8').value).toBe(1);
  expect(ws.getCell('F7').value).toBe('DAYS');
  expect(ws.getColumn('D').hidden && ws.getColumn('E').hidden).toBe(true);
  expect(ws.getCell('F22').formula).toBe('SUM(F8:F20)');
});
```

  Update the existing Medic tests that use `start` and `end`.
- [x] **Step 2:** Run `npm test`. Expected: FAIL.
- [x] **Step 3:** Implement:
  - Medic definitions.
  - `calc.ts` Medic branch: `medicTotal = round2(num(days) * num(dayRate))`, and drop `hoursBetween` for Medic.
  - In `validate.ts`, for Instructor and Medic, push `Row N: days missing.` when `dayRate` is set and `days` is not.
  - In `applyLayout`, for Medic: clear `D7` and `E7` and hide D and E. Clear the `F8:F20` formulas, because F is now an input. Set `H8:H20` to `F${r}*G${r}`. Delete `fixMedicFormulas`, but keep its `K20` and `J22` corrections inside `applyLayout`.
- [x] **Step 4:** Run `npm test`. Expected: PASS.
- [x] **Step 5:** Commit: `feat: medic DAYS column replaces start and end times`.

### Task 7: PSL shows the level, in the grid and in Excel

**Files:**
- Modify: `src/timesheet/types.ts` (`ColumnDef`)
- Modify: `src/timesheet/definitions.ts`
- Modify: `src/excel/fillTemplate.ts`
- Test: `fillTemplate.test.ts`

**Interfaces:**
- `ColumnDef.levelTable?: { labelCol: string; amountCol: string; firstRow: number }`
  - Instructor `dayRate`: `{ labelCol: 'N', amountCol: 'O', firstRow: 8 }`. These are hidden columns, so the 7 rows run N8:O14.
  - Medic `dayRate`: `{ labelCol: 'M', amountCol: 'N', firstRow: 6 }`, the existing hidden table M6:N9.
- The line still stores the dollar amount, so `calc.ts` doesn't change. The grid already shows `optionLabels`.

- [x] **Step 1: Write the failing tests**

```ts
it('instructor PSL cell shows the level and the total looks it up', async () => {
  /* fill Instructor with { date, course: 'A1', days: 1, dayRate: 280 } */
  expect(ws.getCell('E8').value).toBe('Instructor – Level 1 (PSL 3)');
  expect(ws.getCell('F8').formula).toBe('D8*VLOOKUP(E8,$N$8:$O$14,2,FALSE)');
  expect(ws.getCell('N10').value).toBe('Instructor – Level 1 (PSL 3)');
  expect(ws.getCell('O10').value).toBe(280);
  expect((ws.getCell('F8').value as { result?: number }).result).toBe(280);
});
it('medic PSL cell shows the level', async () => {
  /* fill Medic with { date, client: 'X', days: 0.5, dayRate: 350 } */
  expect(ws.getCell('G8').value).toBe('Medic – Level PCP(PSL 9)');
  expect(ws.getCell('H8').formula).toBe('F8*VLOOKUP(G8,$M$6:$N$9,2,FALSE)');
  expect((ws.getCell('H8').value as { result?: number }).result).toBe(175);
});
```

- [x] **Step 2:** Run `npm test`. Expected: FAIL.
- [x] **Step 3:** Implement:
  - `applyLayout` writes each `levelTable` (labels and amounts from `optionLabels`/`options`) and sets the total formulas: Instructor `F${r}`, Medic `H${r}`, for every grid row.
  - The fill loop writes the label instead of the number for a column with `levelTable`.
  - Remove the template's Medic data-validation lists on G8:G20 and the Instructor ones on E8:E20, if present, because they would reject the label text.
  - Remove the Medic table check in `definitions.test.ts`, since it would now compare the table against itself.
- [x] **Step 4:** Run `npm test`. Expected: PASS. Open one filled Medic file in Excel to confirm the totals recalculate.
- [x] **Step 5:** Commit: `feat: PSL shows the level in the emailed Excel`.

### Task 8: Expense Report uses the signed-in user and shows only their reports

**Files:**
- Modify: `src/expenseReport/domain/models.ts`, `domain/report.ts`
- Modify: `src/expenseReport/services/IReportService.ts`, `SharePointReportService.ts`, `MemoryReportService.ts`
- Modify: `src/expenseReport/components/ExpenseReport.tsx`
- Modify: `src/webparts/timesheet/TimesheetWebPart.ts` and its manifest (drop `instructorsListTitle`)
- Create: `provisioning/provision-expense-reports.js`, ported from `../ExpenseReportSPFx/scripts/provision-sharepoint-console.js`, without the Instructors list
- Tests: `src/expenseReport/domain.test.ts`, `storage.test.ts`

**Interfaces:**
- Remove `Instructor`, `getInstructors()`, `instructorsListTitle`, and the Instructor dropdown.
- `validateReport(report: ExpenseReport): ValidationIssue[]` loses its roster parameter. With a blank `instructorName`, it returns `{ field: 'instructorId', message: "Your name couldn't be loaded. Refresh the page." }`.
- `IReportService.listReports(ownerEmail: string): Promise<ReportSummary[]>`. SharePoint adds `$expand=Author&$filter=Author/EMail eq '<email>'`, with the email escaped like `odataString`.
- `IExpenseReportProps` adds `user: { key: string; name: string }` (`key` = `pageContext.user.loginName`, `name` = the same `getCurrentEmployee` name the timesheet shows) and `userEmail: string`.
- `instructorId` and `instructorName` stay in the model and are filled from `user` on create and on "New report". No data exists to migrate.

- [x] **Step 1: Write the failing tests**
  - Domain: `validateReport(createReport())` with blank name contains the message above. A report with `instructorName: 'SMITH, A'` has no `instructorId` issue.
  - Storage: `listReports('a@peak.com')` requests a URL containing `$filter=Author/EMail eq 'a%40peak.com'` (or the encoded form the fixture sees) and `$expand=Author`.
  - Remove the roster and instructor tests.
- [x] **Step 2:** Run `npm test`. Expected: FAIL.
- [x] **Step 3:** Implement the changes above. The web part renders the Expense Report only after the name resolves. Until then it shows `Spinner` with the label `Loading your details…`.
- [x] **Step 4:** In the provisioning script, port the ExpenseReports list creation and set `ReadSecurity: 2, WriteSecurity: 2`. Members get Contribute and the `Payroll Owners` group gets Edit, copying the role code in `provision-timesheets.js`. Add a README section.
- [x] **Step 5:** Run `npm test`. Expected: PASS. Commit: `feat: expense report uses the signed-in user and lists only their reports`.

### Task 9: Expense Report pay period, date picker, logo and table styling

**Files:**
- Rename: `src/expenseReport/components/PayPeriodDateField.tsx` to `DateField.tsx`. `label: string`, `otherValue?: string` becomes optional, and the range check only runs when it is given.
- Modify: `ExpenseReport.tsx`, `ExpenseTable.tsx`, `ExpenseReport.module.scss`
- Delete: `src/expenseReport/components/assets/peak-logo.jpg`. Import `../../webparts/timesheet/assets/peak-logo.jpg` and set alt to `Peak Emergency`.

**Interfaces:**
- Consumes: `PayPeriodPicker` (Task 4), `formatDateInput` and `parseDateInput` (`domain/dates.ts`)
- `IExpenseReportProps.getPayPeriods(): Promise<PayPeriod[]>`

- [x] **Step 1:** Replace the two header date fields with `PayPeriodPicker`. Pass `from={parseDateInput(report.fromDate) || ''}` and set `fromDate`/`toDate` with `formatDateInput(p.from/p.to)` on change. Save stays disabled until a period is set; it already fails validation when the dates are empty.
- [x] **Step 2:** In `ExpenseTable`, replace the line date `TextField` with `DateField`, without a label (use `ariaLabel` `Row N date`). The value is still `DD/MM/YY`.
- [x] **Step 3:** In the SCSS:
  - Add `text-align: center` to `.expenseTable thead th`.
  - Add `font-weight: 400` for the tfoot km and cost cells, via new classes `.plainTotal` on those two `<td>`s.
- [x] **Step 4:** Run `npm run build`. Expected: success.
- [ ] **Step 5: Manual check:**
  - The header shows the PEAK-only logo.
  - Picking PP 23 sets the read-only dates.
  - Line dates open a calendar.
  - Headings are centered.
  - The km and cost totals are not bold, and the other totals are still bold.
  - Reopening a saved report selects its period.
- [x] **Step 6:** Commit: `feat: expense report pay period dropdown, date picker and styling`.

### Task 10: Restyle the timesheets like the Expense Report

**Files:**
- Modify: `src/webparts/timesheet/components/TimesheetForm.tsx`, `TimesheetForm.module.scss`
- Modify: `src/webparts/timesheet/components/TimesheetGrid.tsx` (header fills)

The source of truth is `src/expenseReport/components/ExpenseReport.module.scss`. Copy its values, not new ones.

- [x] **Step 1:** Header:
  - The title row uses `.brandRow` values: red `#f00` h1, `clamp(25px, 3.2vw, 38px)`, bold, logo width 158px.
  - The name shows like the `.instructor` dropdown title: bold, `clamp(18px, 2.2vw, 27px)`.
  - `PAY PERIOD` as an h2 in red, 22px, centered, above `PayPeriodPicker`.
  - Surface padding `18px 8px 12px`.
- [x] **Step 2:** Buttons:
  - Remove the `.form button` rule.
  - The type picker uses Fluent `DefaultButton`, with `checked` and `toggle` for the active form.
  - Submit is `PrimaryButton`, with text `Submit` or `Sending…`.
  - The Switch and Cancel confirmation uses `PrimaryButton` and `DefaultButton`.
- [x] **Step 3:** Grid:
  - Header cells: background `#f1f1f1`, 14px bold, centered.
  - Group fills stay as defined.
  - Cell borders `#b2b2b2`, 13px text, row height 36px.
  - Read-only cells: background `#f1f1f1`.
- [x] **Step 4:** Totals row: background `#f1f1f1`, colour `#e00000`, bold.
- [x] **Step 5:** Messages:
  - Errors use the `.errorSummary` look: `2px solid #a4262c`, background `#fff4f4`, colour `#73191e`, padding 12px, with a `role="alert"` heading `Can't submit yet`.
  - Success uses Fluent `MessageBar` with `MessageBarType.success`.
- [x] **Step 6:** Run `npm run build`. Expected: success.
- [ ] **Step 7: Manual check:** put the Admin, Instructor, Medic and Expense Report forms side by side in the workbench. Fonts, header, buttons, borders and totals look the same.
- [x] **Step 8:** Commit: `style: timesheets match the expense report look`.

### Task 11: Pay rates from the signed-in user (BLOCKED until jeff's list arrives)

Do not start until Walter confirms the list contents. If jeff's list doesn't fit the proposed shape below, stop and update the spec first.

**Update:** the Staff PSL sheet arrived and does not fit this shape. See Task 13 and its open questions. Revise the spec and this task before running it.

**Files:**
- Create: `src/timesheet/rates.ts`, test `src/timesheet/rates.test.ts`
- Modify: `src/services/sharepoint.ts`, `provisioning/provision-timesheets.js`, `TimesheetGrid.tsx`, `TimesheetForm.tsx`, `TimesheetWebPart.ts`

**Interfaces:**
- `interface PayRates { Admin?: number; Instructor?: number; Medic?: number }`
- `RATE_KEY: Record<TimesheetType, string> = { Admin: 'rate', Instructor: 'dayRate', Medic: 'dayRate' }`
- `fillDefaultRate(type: TimesheetType, line: Line, rate: number | undefined): Line` sets `line[RATE_KEY[type]] = rate` only when the rate is blank and some other input is set.
- `SpPayRateSource.getMyRates(userId: number): Promise<PayRates | undefined>` reads list `Pay Rates` (`Employee` User field, `AdminRate` Currency, `InstructorPSL` Number, `MedicPSL` Number) with `$filter=EmployeeId eq <userId>`.
- `ITimesheetGridProps.defaultRate?: number`

- [ ] **Step 1: Write the failing tests**

```ts
it('fills a blank rate once the row has input', () =>
  expect(fillDefaultRate('Medic', { date: '2026-11-05' }, 350).dayRate).toBe(350));
it('keeps a rate the user changed', () =>
  expect(fillDefaultRate('Medic', { date: '2026-11-05', dayRate: 450 }, 350).dayRate).toBe(450));
it('leaves an empty row alone', () => expect(fillDefaultRate('Admin', {}, 30).rate).toBeUndefined());
```

- [ ] **Step 2:** Run `npm test`. Expected: FAIL.
- [ ] **Step 3:** Implement `rates.ts` and the source. The grid calls `fillDefaultRate` in `onchange` and writes the value back with `setValueFromCoords`. A PSL value must be one of the column's `options`, or it is ignored.
- [ ] **Step 4:** The form loads the rates once. If there is no row for the current form type, show `No pay rate is set up for you on this timesheet. Contact accounts.` and disable Submit.
- [ ] **Step 5:** Provisioning: create `Pay Rates` with the fields above, with item-level read security so each person sees only their own row, and Payroll Owners has Edit. Fill the rows from jeff's list.
- [ ] **Step 6:** Run `npm test`, expect PASS. Do a manual check: a medic with a PCP level gets `Medic – Level PCP(PSL 9)` on each new row and can change it per shift. Commit: `feat: rates fill from the Pay Rates list`.

### Task 12: Excel templates use the PEAK-only logo (added by Walter mid-run; run right after Task 7)

Today the templates carry older logos: Admin and Instructor show "PEAK emergency training", and Medic shows "PEAK paramedic services" (partly cropped by `<a:srcRect b="18905"/>`). All three must show the same PEAK-only logo the web part uses (`src/webparts/timesheet/assets/peak-logo.jpg`).

**Files:**
- Modify: `Templates/xlsx/Admin Timesheet Template.xlsx`, `Instructor Timesheet Template.xlsx`, `Medic Timesheet Template.xlsx`. Edit the zip directly, as `scripts/Convert-Templates.ps1` does. Do not round-trip through ExcelJS, and never use Excel COM.
- Create: `scripts/Update-TemplateLogo.ps1`, safe to rerun, so the change can be repeated if the logo changes.
- Test: `src/excel/templates.test.ts` (create)
- Modify: `README.md` (re-upload step)

**Interfaces:** none. Leave the `.xlsm` originals in `Templates/` untouched.

- [x] **Step 1: Write the failing test.** For each of the three templates, read the zip (ExcelJS `wb.xlsx.readFile`, or `JSZip` if already installed) and check:
  - the single image's bytes equal `src/webparts/timesheet/assets/peak-logo.jpg`
  - `xl/drawings/drawing1.xml` has no `srcRect`
  - the drawing's `a:ext` `cx/cy` ratio is within 2% of the logo's pixel width/height ratio
- [x] **Step 2:** Run `npm test`. Expected: FAIL.
- [x] **Step 3:** Write the script. For each template:
  - replace `xl/media/image1.jpeg` with the logo bytes
  - in `xl/drawings/drawing1.xml`, remove the `srcRect` element
  - keep the anchor's top-left and width, and set `cy` (and the `to` row offset) so the height matches the logo's aspect ratio
  - check that the logo still fits inside the header rows above the title. If it doesn't, shrink both `cx` and `cy` proportionally instead.

  Run the script.
- [x] **Step 4:** Run `npm test`. Expected: PASS, including the existing definitions and fillTemplate suites.
- [x] **Step 5:** README: "After updating, re-upload the three files in `Templates/xlsx` to the Timesheet Templates library."
- [x] **Step 6:** Commit: `feat: templates use the PEAK-only logo`.

### Task 13: Pay Rates list from the Staff PSL sheet (added by Walter; run before Task 11)

Source: `docs/Staff PSL.pdf` (29 staff). It does not match the shape Task 11 assumed:
- The instructor rate depends on the **course type** (AP, ASU, NUEC-A FULL, NUEC-A BRIDGING, NUEC-I, NUEC-B). It is not one rate per person. For example, Ben Davies (PSL 2/3) is $225 for AP but $150 for NUEC-A FULL.
- Instructor PSL can be split (`2/3`), so it is text, not a number.
- There are no Admin rates.
- Staff are listed by display name, not by account.

This task only creates and seeds the list. Task 11 must be revised (spec first) to use this shape before it runs.

**Files:**
- Modify: `provisioning/provision-timesheets.js`, `README.md`

**List `Pay Rates`** (BaseTemplate 100, on the timesheet web). Fields:

| Display name | Internal name | Type |
|---|---|---|
| Title (existing) | `Title` | Name as printed on the sheet |
| Employee | `Employee` | User (people only) |
| Instructor PSL | `InstructorPSL` | Text (`2/3` allowed; blank if not an instructor) |
| AP | `RateAP` | Currency |
| ASU | `RateASU` | Currency |
| NUEC-A Full | `RateNUECAFull` | Currency |
| NUEC-A Bridging | `RateNUECABridging` | Currency |
| NUEC-I | `RateNUECI` | Currency |
| NUEC-B | `RateNUECB` | Currency |
| Medic PSL | `MedicPSL` | Number (8 to 11) |

- [x] **Step 1:** In the provisioning script, create the list and fields if they are missing, the same way as Pay Period Schedule. Turn on item-level security (`ReadSecurity: 2, WriteSecurity: 2`) and set Members to Read and `Payroll Owners` to Edit, copying the role code used for Timesheet Submissions. Staff must not be able to change their own rates.
- [x] **Step 2:** Seed the 29 rows from the PDF as an inline array (the script runs in the browser console). Skip a row when an item with the same Title already exists, so the script is safe to rerun. Carolyn Nikodym has Medic PSL 9 and no instructor values.
- [x] **Step 3:** Fill `Employee` for each seeded row by matching Title against `web/siteusers?$filter=Title eq '<name>'`. Log every name with no match (`Pay Rates: no account for <name>, set Employee by hand`). Do not guess between several matches. Item-level read only works by Author, so the web part reads by `EmployeeId` (Task 11) and staff get read access through Members.
- [x] **Step 4:** Check the seed against the PDF in Node: 29 rows; Ben Davies NUEC-A Full = 150 and NUEC-B = 225; Ross Goddard NUEC-A Bridging = 225; Richard Bailey all 0; Tytan Wheeler and William Geise PSL 6 at 340.
- [x] **Step 5:** README: "Rerun `provision-timesheets.js`; it creates and seeds Pay Rates. Check the console for staff with no matching account and set their Employee by hand. Accounts maintains rates in that list."
- [x] **Step 6:** Commit: `feat: Pay Rates list seeded from the Staff PSL sheet`.

**Open questions for Walter (answer before Task 11):**
1. PSL 6 is $350 for Jesse Hoevenaars but $340 for Tytan Wheeler and William Geise. In the current Instructor dropdown, 340 is PSL 6 and 350 is PSL 7. Which is right?
2. The Instructor timesheet has no course-type column, only `COURSE #`. How should a row pick AP, ASU, NUEC-A and so on? Add a course-type dropdown, or derive it from the course number?
3. Where do Admin pay rates come from?
