# Timesheet Web Part Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** An SPFx web part where the signed-in user fills in an Admin, Instructor or Medic timesheet in a spreadsheet grid. On submit it records the submission in a SharePoint list, fills in a copy of the real Excel template and emails it from the user's own mailbox to accounts.

**Architecture:** One web part. A single definitions module describes each template (cells, columns, dropdown values). Both the Jspreadsheet CE grid and the ExcelJS template filler are driven from it. Pure TypeScript modules (calc, validation, naming, Excel fill, submission orchestration) are unit-tested with Jest. Thin adapters over the SPFx `SPHttpClient` and `MSGraphClientV3` do the I/O. A PnP PowerShell script provisions the list and the template library.

**Tech Stack:** SPFx (current GA, React web part template), TypeScript, React, Jspreadsheet CE (vanilla core wrapped in a React component), ExcelJS, Microsoft Graph (`/me`, `/me/sendMail`), SharePoint REST via `SPHttpClient`, Jest, PnP.PowerShell.

**Spec:** `docs/superpowers/specs/2026-10-30-timesheet-webpart-design.md` (read it first; section 6 is the cell map).

## Global Constraints

- Timesheet types: `Admin`, `Instructor`, `Medic`. The Expense Report is out of scope.
- Recipient default: `account@peakemergency.com` (web part property `emailTo`).
- Email is sent from the signed-in user via Graph `POST /me/sendMail`, `saveToSentItems: true`. No shared or service mailbox.
- Graph permissions requested in `config/package-solution.json` `webApiPermissionRequests`: `Mail.Send` and `User.Read` (resource `Microsoft Graph`).
- The attachment is `.xlsx`, produced by filling the stored template copy. Never use the grid library's export.
- The attachment matches the template exactly except for these Medic formula changes (rows 8–20 unless stated): `H{r}` = `=G{r}`, `F{r}` = `=MOD(E{r}-D{r},1)*24`, `K20` = `=SUM(I20:J20)`, `J22` = `=SUM(J8:J20)`.
- Pay: Admin = hours × hourly rate; Instructor = days (0.5/1.0) × day rate + admin hours × $25 + travel $100 + per diem $47; Medic = one row per day, line pay = day rate ($250/$350/$450/$550) + travel $100 + per diem $47.
- Name is written as `LAST, FIRST` in uppercase. Course # and client are uppercase.
- Dates display as `dd/mm/yy`.
- List `Timesheet Submissions` columns: Title, TimesheetType (choice Admin/Instructor/Medic), Employee (person), PayPeriodStart, PayPeriodEnd (date only), TotalAmount (currency), LineItems (multi-line plain text, JSON), Status (choice Pending/Sent/Email failed), StatusDetail (multi-line text). The `.xlsx` is stored as an item attachment. Item-level security: users read and edit only their own items.
- Template library `Timesheet Templates` holds `Admin Timesheet Template.xlsx`, `Instructor Timesheet Template.xlsx`, `Medic Timesheet Template.xlsx`.
- File name: `<Type> Timesheet - <LAST, FIRST> - <yyyy-mm-dd>_<yyyy-mm-dd>.xlsx`. Subject: `<Type> Timesheet – <LAST, FIRST> – <dd/mm/yy> to <dd/mm/yy>`.
- Do not change Admin column H.

## Review Focus

1. **Dates shifting by a day in the attachment.** Users are in Mountain Time, and ExcelJS converts JS `Date` to a serial number via UTC. 15/10/26 must open as 15/10/26. Pinned by `fillTemplate writes dates as the same calendar day` (Task 5).
2. **Overnight medic shifts.** 19:00 → 07:00 must show 12.00 hours on screen and in the file, never a negative number. Pinned by `medic overnight shift is 12 hours` (Task 3) and the `F8` formula assertion (Task 5).
3. **Times typed in different forms.** `8:30`, `08:30`, `0830` and `8:30 pm` should parse; `25:00` and `abc` should give a row-specific error, not a silent 0. Pinned by `parseTime` tests (Task 3) and `invalid time is reported with its row` (Task 4).
4. **Gaps between filled rows.** Rows 1 and 5 filled, rows 2–4 blank: each row must land in the same template row it sits in on screen, and blanks stay blank. Pinned by `fillTemplate keeps row positions` (Task 5).
5. **Email failure after the list item exists.** The item must end up with status `Email failed`, the error text, and the file attached, and the user must see the failure. Pinned by `send failure marks item Email failed and keeps file` (Task 6).

---

## File Structure

```
config/package-solution.json            Graph permission requests (modify)
src/timesheet/types.ts                  TimesheetType, ColumnDef, TemplateDef, Line, TimesheetData
src/timesheet/definitions.ts            TEMPLATES: the per-template cell/column map
src/timesheet/values.ts                 parseTime, normalizeDate, toExcelDate
src/timesheet/calc.ts                   computeLine, computeTotals
src/timesheet/validate.ts               validate
src/timesheet/naming.ts                 formatEmployeeName, buildFileName, buildSubject
src/excel/fillTemplate.ts               fillTemplate (ExcelJS)
src/services/submit.ts                  submitTimesheet + SubmissionStore/TemplateSource/Mailer interfaces
src/services/sharepoint.ts              SpSubmissionStore, SpTemplateSource
src/services/graph.ts                   GraphMailer, getCurrentEmployee
src/webparts/timesheet/TimesheetWebPart.ts             properties: emailTo, listTitle, templateLibraryUrl
src/webparts/timesheet/components/TimesheetForm.tsx    type picker, header, grid, totals, submit
src/webparts/timesheet/components/TimesheetGrid.tsx    Jspreadsheet CE wrapper
src/webparts/timesheet/components/TimesheetForm.module.scss
Templates/xlsx/*.xlsx                   macro-free template copies (fixtures + upload source)
scripts/Convert-Templates.ps1           .xlsm → .xlsx via Excel COM
provisioning/Provision-Timesheets.ps1   list + library + template upload
README.md                               build, deploy, admin approval, provisioning
```

Tests sit next to their modules as `*.test.ts`.

---

### Task 1: Scaffold, test runner and template copies

**Files:**
- Create: SPFx solution in the repo root (`yo @microsoft/sharepoint`: solution `peak-timesheets`, web part `Timesheet`, React framework, current folder)
- Create: `scripts/Convert-Templates.ps1`, `Templates/xlsx/{Admin,Instructor,Medic} Timesheet Template.xlsx`
- Create: `src/timesheet/smoke.test.ts`
- Modify: `.gitignore` (keep the scaffold's; make sure `Templates/xlsx/` is **not** ignored)

**Interfaces:**
- Produces: `npm test` runs Jest over `src/**/*.test.ts` in Node, and the three `.xlsx` fixtures used by Tasks 2 and 5.

- [ ] **Step 1: Scaffold.** Run `yo @microsoft/sharepoint` with the answers above and do not commit `node_modules`. Expected: `npm run build` (or the scaffold's equivalent build script) completes.
- [ ] **Step 2: Wire Jest.** If the scaffold already ships Jest (Heft toolchain), use it. Otherwise add `jest`, `ts-jest`, `@types/jest` and a `jest.config.js` (`testEnvironment: 'node'`, `testMatch: ['<rootDir>/src/**/*.test.ts']`), and set `"test": "jest"`. Write `smoke.test.ts` with `expect(1 + 1).toBe(2)`. Run `npm test`. Expected: 1 passed.
- [ ] **Step 3: Write `scripts/Convert-Templates.ps1`.** It opens each `Templates/*Timesheet Template (current).xlsm` (not the Expense Report) with Excel COM, with `$excel.EnableEvents = $false` so the `Workbook_BeforeSave` macro can't cancel the save. It runs `SaveAs(<abs path>\Templates\xlsx\<Type> Timesheet Template.xlsx, 51)`, then closes and quits Excel.
- [ ] **Step 4: Run it.** Run `pwsh scripts/Convert-Templates.ps1`. Expected: three `.xlsx` files in `Templates/xlsx/`, each opening in Excel with the same look as the `.xlsm`.
- [ ] **Step 5: Commit.** `git add -A && git commit -m "chore: scaffold SPFx solution, jest, macro-free template copies"`

---

### Task 2: Template definitions

**Files:**
- Create: `src/timesheet/types.ts`, `src/timesheet/definitions.ts`
- Test: `src/timesheet/definitions.test.ts` (dev dependency: `exceljs`; it's also a runtime dependency from Task 5, so install it here: `npm i exceljs`)

**Interfaces:**
- Produces:
```ts
export type TimesheetType = 'Admin' | 'Instructor' | 'Medic';
export type ColumnKind = 'date' | 'time' | 'upper' | 'number' | 'dropdown' | 'computed';
export interface ColumnDef { key: string; col: string; header: string; kind: ColumnKind; options?: number[]; width: number; }
export interface TemplateDef {
  type: TimesheetType; title: string; sheetName: string; fileName: string;
  nameCell: string; fromCell: string; toCell: string;
  firstRow: number; rowCount: number; headerRow: number;
  columns: ColumnDef[];
  totals: { key: string; cell: string }[];   // on-screen totals = template totals row
  grandTotalKey: string;
}
export type Line = Record<string, string | number | null>; // date 'yyyy-mm-dd', time 'HH:MM', numbers
export interface TimesheetData { type: TimesheetType; name: string; from: string | null; to: string | null; lines: Line[]; }
export const TEMPLATES: Record<TimesheetType, TemplateDef>;
```
- Column keys (used by every later task):
  - **Admin:** `date` A, `start` C, `end` D, `deductions` E [0, 0.5, 1, 1.25, 1.5, 1.75, 2], `hours` F computed, `rate` G number, `amount` I computed. Rows 7–22, header row 6, name A4, from G4, to I4. Totals: `deductions` E25, `hours` F25, `amount` I25. Grand total `amount`.
  - **Instructor:** `date` A, `course` C upper, `days` D [0.5, 1], `dayRate` E [0, 150, 225, 280, 320, 340, 350], `instrTotal` F computed, `adminHours` G [0.25…8 in 0.25 steps], `adminRate` H [25], `adminTotal` I computed, `travel` J [100], `perDiem` K [47], `otherTotal` L computed, `amount` M computed. Rows 8–20, header row 7, name A4, from J4, to M4. Totals: `days` D22, `adminHours` G22, `travel` J22, `perDiem` K22, `amount` M22. Grand total `amount`.
  - **Medic:** `date` A, `client` C upper, `start` D, `end` E, `hours` F computed, `dayRate` G [250, 350, 450, 550], `medicTotal` H computed, `travel` I [100], `perDiem` J [47], `otherTotal` K computed, `amount` L computed. Rows 8–20, header row 7, name A4, from H4, to K4. Totals: `hours` F22, `travel` I22, `perDiem` J22, `amount` L22. Grand total `amount`.
  - `title`: `ADMIN TIMESHEET` / `INSTRUCTOR TIMESHEET` / `MEDIC TIMESHEET`. `sheetName`: `<Type> Timesheet Template`. `fileName`: `<Type> Timesheet Template.xlsx`. `header`: the template's header text with whitespace runs (including line breaks) collapsed to one space. `width`: the template column width × 7 px.

- [ ] **Step 1: Write the failing test.** `definitions match the real templates`. For each type, load `Templates/xlsx/<fileName>` with ExcelJS, then:
  - assert the sheet `sheetName` exists;
  - assert that for every column, the cell `${col}${headerRow}` text (whitespace runs including newlines collapsed to one space, trimmed) equals `header` (Admin headers are on row 6; the Instructor and Medic group headings on row 6 are not columns);
  - assert that for every `computed` column, `${col}${firstRow}` holds a formula;
  - assert `TEMPLATES.Medic.columns.find(c => c.key === 'dayRate')!.options` equals `[250, 350, 450, 550]`.
- [ ] **Step 2: Run.** Run `npm test -- definitions`. Expected: FAIL, cannot find module `./definitions`.
- [ ] **Step 3: Implement `types.ts` and `TEMPLATES` in `definitions.ts`** with the values above.
- [ ] **Step 4: Run.** Run `npm test -- definitions`. Expected: PASS.
- [ ] **Step 5: Commit.** `git add src/timesheet package.json package-lock.json && git commit -m "feat: template definitions tied to the real workbooks"`

---

### Task 3: Value parsing and pay calculation

**Files:**
- Create: `src/timesheet/values.ts`, `src/timesheet/calc.ts`
- Test: `src/timesheet/values.test.ts`, `src/timesheet/calc.test.ts`

**Interfaces:**
- Consumes: `TimesheetType`, `Line`, `TEMPLATES` (Task 2).
- Produces:
```ts
export function parseTime(input: string): string | null;        // → 'HH:MM' 24h, null if invalid
export function normalizeDate(input: string): string | null;    // grid value → 'yyyy-mm-dd'
export function toExcelDate(isoDate: string): Date;             // UTC midnight of that day
export function timeToDayFraction(hhmm: string): number;        // '06:00' → 0.25
export function computeLine(type: TimesheetType, line: Line): Line;   // returns line with computed keys filled
export function computeTotals(type: TimesheetType, lines: Line[]): Record<string, number>; // keys = TEMPLATES[type].totals keys
```
- Computed values on an empty row are `null`, not 0. Hours round to 2 decimals and money to cents.

- [ ] **Step 1: Write failing tests.**
  - `parseTime`: `'8:30'`→`'08:30'`, `'08:30'`→`'08:30'`, `'0830'`→`'08:30'`, `'8:30 pm'`→`'20:30'`, `'12:15 am'`→`'00:15'`, `'25:00'`→`null`, `'abc'`→`null`, `''`→`null`.
  - `normalizeDate`: `'2026-10-15 00:00:00'`→`'2026-10-15'`, `'2026-10-15'`→`'2026-10-15'`, `''`→`null`.
  - `toExcelDate('2026-10-15').toISOString()` → `'2026-10-15T00:00:00.000Z'`.
  - `admin line`: `{date:'2026-10-15', start:'08:00', end:'16:30', deductions:0.5, rate:20}` → `hours` 8, `amount` 160.
  - `instructor line`: `{date:'2026-10-15', days:0.5, dayRate:280, adminHours:2, adminRate:25, travel:100, perDiem:47}` → `instrTotal` 140, `adminTotal` 50, `otherTotal` 147, `amount` 337.
  - `medic overnight shift is 12 hours`: `{date:'2026-10-15', start:'19:00', end:'07:00', dayRate:450, travel:100, perDiem:47}` → `hours` 12, `medicTotal` 450, `otherTotal` 147, `amount` 597.
  - `medic pay ignores hours`: the same line with `start:'08:00', end:'09:00'` → `medicTotal` 450.
  - `computeTotals('Medic', [line, emptyLine])` → `{hours:12, travel:100, perDiem:47, amount:597}`.
- [ ] **Step 2: Run.** Run `npm test -- values calc`. Expected: FAIL, modules not found.
- [ ] **Step 3: Implement** `values.ts` and `calc.ts`. Medic hours = `((end − start) mod 24h)`. Admin hours = `end − start − deductions`, with no wrap (Admin end before start is a validation error in Task 4).
- [ ] **Step 4: Run.** Run `npm test -- values calc`. Expected: PASS.
- [ ] **Step 5: Commit.** `git commit -m "feat: time/date parsing and pay calculation"`

---

### Task 4: Validation and naming

**Files:**
- Create: `src/timesheet/validate.ts`, `src/timesheet/naming.ts`
- Test: `src/timesheet/validate.test.ts`, `src/timesheet/naming.test.ts`

**Interfaces:**
- Consumes: `TimesheetData`, `TEMPLATES`, `parseTime`, `normalizeDate`.
- Produces:
```ts
export function validate(data: TimesheetData): string[];   // [] = ok; messages shown verbatim to the user
export function formatEmployeeName(givenName: string | null, surname: string | null, displayName: string): string;
export function buildFileName(data: TimesheetData): string;
export function buildSubject(data: TimesheetData): string;
```
- Validation messages (exact copy): `Your name is missing.`, `Pick the pay period From date.`, `Pick the pay period To date.`, `The To date can't be before the From date.`, `Enter at least one row.`, `Row {n}: date missing.`, `Row {n}: "{value}" is not a valid time.`, `Row {n}: end time is before start time.` (Admin only). `{n}` is the 1-based grid row. A row counts as entered if any non-computed key has a value.

- [ ] **Step 1: Write failing tests.**
  - Empty Medic data with name `''` → contains all of `Your name is missing.`, `Pick the pay period From date.`, `Pick the pay period To date.`, `Enter at least one row.`
  - `from:'2026-10-15', to:'2026-10-01'` → contains `The To date can't be before the From date.`
  - `invalid time is reported with its row`: Medic line 3 `{date:'2026-10-15', start:'abc'}` → contains `Row 3: "abc" is not a valid time.`
  - Admin line 1 with start `'17:00'`, end `'09:00'` → contains `Row 1: end time is before start time.`
  - A line with `travel: 100` and no date → `Row 1: date missing.`
  - Valid Admin data → `[]`.
  - `formatEmployeeName('Carolyn','Nikodym','Carolyn Nikodym')` → `'NIKODYM, CAROLYN'`; `formatEmployeeName(null,null,"Katie O'Connor")` → `"O'CONNOR, KATIE"`; `formatEmployeeName(null,null,'Madonna')` → `'MADONNA'`.
  - `buildFileName` for Medic, `NIKODYM, CAROLYN`, 2026-10-01 to 2026-10-15 → `'Medic Timesheet - NIKODYM, CAROLYN - 2026-10-01_2026-10-15.xlsx'`. `buildSubject` → `'Medic Timesheet – NIKODYM, CAROLYN – 01/10/26 to 15/10/26'`.
- [ ] **Step 2: Run.** Run `npm test -- validate naming`. Expected: FAIL.
- [ ] **Step 3: Implement** both modules. `formatEmployeeName` falls back to the last word of `displayName` as the surname when `surname` is empty.
- [ ] **Step 4: Run.** Run `npm test -- validate naming`. Expected: PASS.
- [ ] **Step 5: Commit.** `git commit -m "feat: validation, employee name and file naming"`

---

### Task 5: Fill the Excel template

**Files:**
- Create: `src/excel/fillTemplate.ts`
- Test: `src/excel/fillTemplate.test.ts`

**Interfaces:**
- Consumes: `TEMPLATES`, `TimesheetData`, `toExcelDate`, `timeToDayFraction`.
- Produces: `export async function fillTemplate(template: ArrayBuffer, data: TimesheetData): Promise<ArrayBuffer>;`
- Behaviour:
  - Write `name` to `nameCell`, and From/To as `toExcelDate`.
  - For grid row `i` (0-based), write each non-computed column to `${col}${firstRow + i}`: dates via `toExcelDate`, times via `timeToDayFraction`, upper columns uppercased, numbers as numbers. Write nothing for blank values.
  - Never touch computed cells, except the Medic overrides in Global Constraints.
  - Keep existing number formats and styles (assign `.value` only), and set `workbook.calcProperties.fullCalcOnLoad = true`.
  - In the browser build, import ExcelJS as `import('exceljs/dist/exceljs.min.js')` (dynamic, so it stays out of the main bundle and avoids webpack Node polyfill errors). Tests may import `exceljs` directly; keep the import in one small `loadExcelJs()` helper so both work.

- [ ] **Step 1: Write failing tests.** Load the fixture `.xlsx` with `fs.readFileSync`, call `fillTemplate`, then reload the output with ExcelJS.
  - `fillTemplate writes header`: Medic `A4` = `'NIKODYM, CAROLYN'`; `H4` is a Date; merge `A4:E4` still present.
  - `fillTemplate writes dates as the same calendar day`: Medic line date `'2026-10-15'` → `A8` value `toISOString().slice(0,10)` = `'2026-10-15'`, and `numFmt` unchanged from the fixture.
  - `fillTemplate keeps row positions`: Medic lines `[line, {}, {}, {}, line2]` → `C8` and `C12` filled, `C9:C11` empty.
  - `medic formula fixes`: `H8.formula` = `'G8'`, `F8.formula` = `'MOD(E8-D8,1)*24'`, `K20.formula` = `'SUM(I20:J20)'`, `J22.formula` = `'SUM(J8:J20)'`.
  - `other templates untouched`: Instructor `M8.formula` = `'SUM(F8,I8,L8)'`, Admin `F7.formula` = `'((D7-C7)*24-E7)'`. Each sheet's column widths equal the fixture's.
  - `times written as day fractions`: Medic start `'06:00'` → `D8` = 0.25.
  - `uppercase`: client `'acme'` → `C8` = `'ACME'`.
- [ ] **Step 2: Run.** Run `npm test -- fillTemplate`. Expected: FAIL.
- [ ] **Step 3: Implement** `fillTemplate`.
- [ ] **Step 4: Run.** Run `npm test -- fillTemplate`. Expected: PASS. Also save one filled Medic output to `%TEMP%`, open it in Excel, and confirm it looks identical to the template with H8 = the day rate.
- [ ] **Step 5: Commit.** `git commit -m "feat: fill real Excel template with timesheet data"`

---

### Task 6: Submission flow

**Files:**
- Create: `src/services/submit.ts`
- Test: `src/services/submit.test.ts` (hand-written fakes, no mocking library)

**Interfaces:**
- Consumes: `validate`, `computeLine`, `computeTotals`, `fillTemplate`, `buildFileName`, `buildSubject`, `TEMPLATES`.
- Produces:
```ts
export type SubmissionStatus = 'Pending' | 'Sent' | 'Email failed';
export interface SubmissionRecord { title: string; type: TimesheetType; employeeId: number; from: string; to: string; totalAmount: number; lineItemsJson: string; }
export interface SubmissionStore { create(r: SubmissionRecord): Promise<number>; attach(id: number, fileName: string, content: ArrayBuffer): Promise<void>; setStatus(id: number, status: SubmissionStatus, detail?: string): Promise<void>; }
export interface TemplateSource { load(type: TimesheetType): Promise<ArrayBuffer>; }
export interface MailMessage { to: string; subject: string; bodyText: string; fileName: string; content: ArrayBuffer; }
export interface Mailer { send(m: MailMessage): Promise<void>; }
export interface SubmitDeps { store: SubmissionStore; templates: TemplateSource; mailer: Mailer; emailTo: string; employeeId: number; }
export type SubmitResult = { ok: true } | { ok: false; errors: string[] };
export async function submitTimesheet(data: TimesheetData, deps: SubmitDeps): Promise<SubmitResult>;
```
- Order: validate (on errors, return them and touch nothing) → `create` (Pending) → `templates.load` → `fillTemplate` → `attach` → `mailer.send` → `setStatus('Sent')`. Any failure after `create` calls `setStatus(id, 'Email failed', err.message)` and returns `{ok:false, errors:['Your timesheet was saved but the email was not sent: ' + err.message]}`. `lineItemsJson` holds only entered rows, with computed values. The body text is `<Type> timesheet for <name>, <dd/mm/yy> to <dd/mm/yy>. Total: $<amount>.`

- [ ] **Step 1: Write failing tests.**
  - `happy path`: fakes record calls. Expected order `create, load, attach, send, setStatus(Sent)`; the `create` record has `totalAmount` 597 for the Task 3 medic line; `send` gets `to` = `deps.emailTo` and the `buildFileName` file name.
  - `invalid data never touches SharePoint`: name `''` → `ok:false`, fakes have zero calls.
  - `send failure marks item Email failed and keeps file`: mailer throws `Error('Mailbox unavailable')`. Expected: `attach` was called, `setStatus(id,'Email failed','Mailbox unavailable')`, result errors[0] contains `Mailbox unavailable`.
  - `template load failure marks item Email failed`.
- [ ] **Step 2: Run.** Run `npm test -- submit`. Expected: FAIL.
- [ ] **Step 3: Implement** `submitTimesheet`.
- [ ] **Step 4: Run.** Run `npm test -- submit`. Expected: PASS.
- [ ] **Step 5: Commit.** `git commit -m "feat: submission orchestration with failure status"`

---

### Task 7: SharePoint and Graph adapters

**Files:**
- Create: `src/services/sharepoint.ts`, `src/services/graph.ts`
- Modify: `config/package-solution.json`

**Interfaces:**
- Consumes: `SubmissionStore`, `TemplateSource`, `Mailer`, `MailMessage`, `formatEmployeeName`.
- Produces:
```ts
export class SpSubmissionStore implements SubmissionStore { constructor(client: SPHttpClient, webUrl: string, listTitle: string); }
export class SpTemplateSource implements TemplateSource { constructor(client: SPHttpClient, libraryServerRelativeUrl: string); } // loads `${lib}/${TEMPLATES[type].fileName}`
export class GraphMailer implements Mailer { constructor(client: MSGraphClientV3); }
export async function getCurrentEmployee(client: MSGraphClientV3): Promise<string>; // GET /me?$select=givenName,surname,displayName → formatEmployeeName
```
- REST:
  - `create`: POST `_api/web/lists/getbytitle('<list>')/items` with field internal names `Title, TimesheetType, EmployeeId, PayPeriodStart, PayPeriodEnd, TotalAmount, LineItems, Status, StatusDetail`.
  - `attach`: POST `.../items(<id>)/AttachmentFiles/add(FileName='<encoded>')`.
  - `setStatus`: POST with `X-HTTP-Method: MERGE`, `IF-MATCH: *`.
  - Template: GET `_api/web/GetFileByServerRelativePath(decodedurl='<path>')/$value` → `arrayBuffer()`.
  - Mail: `POST /me/sendMail` `{ message: { subject, body:{contentType:'Text', content}, toRecipients:[{emailAddress:{address}}], attachments:[{'@odata.type':'#microsoft.graph.fileAttachment', name, contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes:<base64>}] }, saveToSentItems:true }`.
  - Non-2xx responses throw `Error` with the response's error message.
- Employee id comes from `context.pageContext.legacyPageContext.userId` (passed in by the web part, Task 8).

- [ ] **Step 1: Implement** both files, and add `webApiPermissionRequests: [{resource:'Microsoft Graph', scope:'Mail.Send'}, {resource:'Microsoft Graph', scope:'User.Read'}]` to `package-solution.json`.
- [ ] **Step 2: Build.** Run `npm run build`. Expected: no TypeScript or lint errors. These adapters are verified against a real tenant in Task 10; they have no unit tests.
- [ ] **Step 3: Commit.** `git commit -m "feat: SharePoint and Graph adapters, Graph permission requests"`

---

### Task 8: Web part UI

**Files:**
- Modify: `src/webparts/timesheet/TimesheetWebPart.ts` (+ its manifest's default properties and `loc` strings)
- Create: `src/webparts/timesheet/components/TimesheetForm.tsx`, `TimesheetGrid.tsx`, `TimesheetForm.module.scss`

**Interfaces:**
- Consumes: everything above.
- Web part properties (property pane text fields): `emailTo` (default `account@peakemergency.com`), `listTitle` (default `Timesheet Submissions`), `templateLibraryUrl` (default `<site server-relative url>/Timesheet Templates`).
- `TimesheetGrid` props: `{ def: TemplateDef; lines: Line[]; onChange(lines: Line[]): void }`. It wraps the vanilla `jspreadsheet-ce` (`npm i jspreadsheet-ce`) in a `useRef`/`useEffect` component, not a React wrapper package, so it doesn't depend on React version support. The grid has `def.rowCount` fixed rows (no insert/delete) and one column per `def.columns`. Column types:
  - `date` → `calendar` with format `DD/MM/YY`;
  - `time` → `text`, normalized with `parseTime` on change;
  - `dropdown` → `dropdown` with options formatted like the template (`$250.00`);
  - `upper` → `text`, uppercased on change;
  - `computed` → `readOnly`, filled from `computeLine`.
  - Destroy the instance on unmount and when `def` changes.
- `TimesheetForm`: type picker (three buttons), then a header styled like the template's row 1–4 block: title bar, name (read-only, from `getCurrentEmployee`), PAY PERIOD From/To as `<input type="date">`. Then the grid, a TOTALS row from `computeTotals`, a Submit button (disabled while submitting), and a message area that lists `validate`/`submitTimesheet` errors verbatim or shows `Timesheet sent to <emailTo>.`. After success, clear the lines and dates. Switching type with rows entered asks for confirmation with an inline prompt (not `window.confirm`).
- Styling: copy the template's fonts, header fill and border colours from the fixture workbook (`xl/styles.xml`) into the SCSS, so the screen reads as the same form.

- [ ] **Step 1: Implement** the web part properties and both components.
- [ ] **Step 2: Build.** Run `npm run build` and `npm test`. Expected: build clean, all tests pass.
- [ ] **Step 3: Workbench check.** Run `npm run start` (or `gulp serve`), add the web part to the hosted workbench (`<site>/_layouts/15/workbench.aspx`), and for each type:
  - enter the Task 3 sample row and confirm the line and totals match Task 3's numbers;
  - type `0830` in a time cell and see `08:30`;
  - press Submit with no dates and see the validation messages.
- [ ] **Step 4: Commit.** `git commit -m "feat: timesheet form with template-styled grid"`

---

### Task 9: Provisioning and deployment docs

**Files:**
- Create: `provisioning/Provision-Timesheets.ps1`, `README.md`

**Interfaces:**
- Script parameters: `-SiteUrl` (required), `-AccountsGroup` (required, SharePoint group name), `-ListTitle` (default `Timesheet Submissions`), `-LibraryTitle` (default `Timesheet Templates`).
- The script is idempotent (skips anything that already exists):
  - `Connect-PnPOnline -Interactive`;
  - create the list with the Global Constraints columns, and attachments enabled;
  - `Set-PnPList -ReadSecurity 2 -WriteSecurity 2`;
  - give `-AccountsGroup` the `Edit` role on the list, so they see all items;
  - create the library and upload `Templates/xlsx/*.xlsx`.
- `README.md`:
  - prerequisites (Node version from the SPFx compatibility table, PnP.PowerShell);
  - `npm install`, `npm test`, `npm run start`;
  - how to package the `.sppkg` (the scaffold's ship and package scripts);
  - uploading it to the app catalog;
  - the tenant admin approving `Mail.Send` and `User.Read` in SharePoint admin center → Advanced → API access;
  - running the provisioning script;
  - updating a template (rerun `Convert-Templates.ps1` and replace the file in the library, keeping the file name and cell positions).

- [ ] **Step 1: Write** the script and README.
- [ ] **Step 2: Run** the script against a dev/test site twice. Expected: first run creates the list and library and uploads 3 files; second run reports everything already exists and changes nothing.
- [ ] **Step 3: Commit.** `git commit -m "chore: provisioning script and deployment README"`

---

### Task 10: Tenant acceptance

**Files:** none (fix anything found in the owning task's files, with a test where one applies).

- [ ] **Step 1: Deploy.** Package and deploy to the test tenant, and approve the API permissions. Add the web part to a page.
- [ ] **Step 2: Submit one of each type** with sample rows, including a medic overnight shift. For each one, confirm:
  - the email arrives at the configured address, from the submitting user, with one `.xlsx`;
  - the attachment opens in Excel looking like the template, with values in the section 6 cells and totals equal to the on-screen totals (Medic line total = day rate);
  - a `Timesheet Submissions` item has status `Sent` and the same file attached.
- [ ] **Step 3: Failure path.** Temporarily set `emailTo` to an invalid address such as `not-an-address`, then submit. Expected: an error message on screen and the list item status `Email failed`, with the detail and the file attached. Reset `emailTo` afterwards.
- [ ] **Step 4: Permissions.** As a second, non-accounts user, open the list. Expected: only their own items are visible. As an accounts group member: all items are visible.
- [ ] **Step 5: Commit** any fixes, and tag the release: `git tag v1.0.0`.
