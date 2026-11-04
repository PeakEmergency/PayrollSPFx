# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An SPFx 1.24 web part (`peak-timesheets.sppkg`) for Peak Emergency. The user picks their name from the `Employees` list and fills in an Admin, Instructor or Medic timesheet grid that looks like the Excel templates. On submit it:

1. creates an item in the SharePoint list `Timesheet Submissions` (Status `Pending`),
2. fills the matching `.xlsx` template from the `Timesheet Templates` library with ExcelJS and attaches it to the item,
3. sets Status to `Queued`.

A Power Automate flow ("Timesheet email") then emails the attachment from the `schedule@` mailbox to accounts and sets Status to `Sent` or `Email failed`. The web part never sends mail itself. Setup steps for the list, library, flow and deployment are in `README.md`. Design and plan: `docs/superpowers/specs/` and `docs/superpowers/plans/`.

## Commands

- `npm test`: `heft test --clean`. Jest runs on the compiled `lib-commonjs` output, not on `src`.
- `npm run build`: production test run plus `heft package-solution`, producing `sharepoint/solution/peak-timesheets.sppkg`.
- `npm start`: local workbench.

The repo lives in OneDrive, which can lock files. If a build fails with `EPERM` on `jest-output`, run `rm -rf jest-output` and retry.

## Layout

- `src/timesheet/`: pure logic, no SharePoint. `definitions.ts` maps each template's columns to Excel cells (widths are `round(excelWidth * 7)`), `calc.ts` computes line and grand totals, `validate.ts` checks input, `naming.ts` builds the file name and email subject, `values.ts` parses times and dates.
- `src/excel/fillTemplate.ts`: writes inputs into the template, fixes the Medic formulas, and stores cached formula results so previewers show totals. ExcelJS is loaded with a dynamic `import('exceljs/dist/exceljs.min.js')`.
- `src/services/`: `submit.ts` runs the submit sequence above; `sharepoint.ts` does list and library calls via `SPHttpClient` and names the failing step in errors. No Graph or other API permissions: names come from the `Employees` list (M365 accounts are unreliable), and rates from `Pay Rates` matched by name (`rateKey`).
- `src/expenseReport/`: the Expense Report form, copied from the separate `ExpenseReportSPFx` repo (built with ChatGPT) and kept as is. The "Expense Report" button shows it in place of the timesheet. It saves to the `ExpenseReports` list, reads the `Employees` list for its Instructor dropdown, and sends no email. Both forms stay mounted (hidden) after first open, so switching loses nothing.
- `src/webparts/timesheet/`: web part and React UI. `TimesheetGrid.tsx` wraps Jspreadsheet CE in a ref.
- `provisioning/provision-timesheets.js`: REST script run in the browser console on the site to create the list, fields, permissions and library. Safe to rerun.
- `scripts/Convert-Templates.ps1`: converts `Templates/*.xlsm` to `Templates/xlsx/*.xlsx` by editing the zip (Excel COM hangs on the Medic file).

## Templates

`Templates/*.xlsm` are the original macro-enabled forms. `Templates/xlsx/*.xlsx` are the converted copies used by the tests and uploaded to the `Timesheet Templates` library. The Expense Report template is not used yet.

Named ranges: `RequiredFields` (Admin and Instructor) lists the required header cells. `InstructorList` is broken in two files (Admin points at an external workbook, Medic resolves to `#REF!`).

To inspect a workbook without Excel, treat it as a zip: `unzip -p "<file>" xl/workbook.xml` shows sheets and named ranges; cell data lives in `xl/worksheets/sheetN.xml` and `xl/sharedStrings.xml`.
