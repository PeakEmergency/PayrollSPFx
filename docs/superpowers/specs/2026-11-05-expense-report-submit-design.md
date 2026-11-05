# Expense report submit and email: design

Date: 2026-11-05

## Why

Expense reports are saved to the `ExpenseReports` list but never reach accounts. Timesheets are emailed to accounts by the "Timesheet email" flow; expense reports should be too. The "Expense report email" flow now exists (turned off) and sends every attachment of an `ExpenseReports` item whose `Status` becomes `Queued`. The web part has to produce what that flow sends.

## What the client decided (2026-11-05)

- An instructor can save drafts and come back to them before submitting.
- Submitting locks the report, like a submitted timesheet. More expenses go on a new report.
- The email carries the receipts and the filled Excel expense form.
- One person can submit more than one expense report for the same pay period.

## Goals

- A **Submit** button next to **Save** on the Expense Report form.
- Submit saves the report, uploads pending receipts, fills the Expense Report Excel template, attaches it to the item, fills `Email To`, `Email Subject` and `Email Body`, and sets `Status` to `Queued`. The flow then emails it and sets `Sent` or `Email failed`.
- A submitted report opens read-only, showing its status.

## Non-goals

- No email from the web part itself (the flow sends it, as for timesheets).
- No server-side lock. Members keep Contribute on their own items; the lock is in the form, the same as timesheets.
- No resubmit or unlock of a submitted report. A report that ends in `Email failed` is handled by accounts (they can set `Status` back to `Queued` in the list to make the flow retry).
- No new list or permission changes beyond the columns already added.

## Anti-goals

- A submitted report must never be editable from the form, even after the flow changes the item.
- A failed submit must never leave a report that looks submitted but has no Excel file or no email columns. It must also never lose the user's entries or receipts.
- Several reports in one pay period must not collide (file names, `findClient` duplicate checks).

## Constraints

- SPFx 1.24, React, `SPHttpClient`, no Graph (see `CLAUDE.md`).
- `src/expenseReport/` was copied from another repo; change it as little as the feature needs.
- `Status` on `ExpenseReports` is a Choice with no default: blank means draft. Values: `Pending`, `Queued`, `Sent`, `Email failed`.
- The flow PATCHes the item after sending, which changes its ETag. A submitted report is never saved again, so this does not matter once it is locked.
- Recipient: the web part's existing `emailTo` property, the same one timesheets use.
- ExcelJS is already loaded with a dynamic import for timesheets; reuse that loader.

## How it works

### States

| Status | Form |
|---|---|
| blank | Draft. Save and Submit enabled. |
| `Pending` | Submit started but did not finish (browser closed mid-submit). Read-only, with a **Finish submitting** button that reruns steps 4 and 5 (my call). |
| `Queued`, `Sent`, `Email failed` | Submitted. Read-only, status shown in a banner. |

The saved-reports list in the form shows each report's status (blank shows as "Draft").

### Submit sequence

1. Run the existing save validation. Submit also requires at least one expense row (my call). Receipts are not required.
2. Save the report as Save does today (create or MERGE with ETag), and upload pending receipts. Stop on any failure; nothing is locked yet.
3. Set `Status` = `Pending` (MERGE with the current ETag). From here the form is read-only.
4. Fill the Expense Report template with ExcelJS and attach it to the item as `Expense Report - <NAME> - <from>_<to> - <ID>.xlsx` (the item ID keeps several reports for one pay period apart). On a retry, replace an attachment of that name instead of failing. A receipt the user named the same way is not expected; if one exists, fail with a clear message.
5. MERGE `Email To`, `Email Subject`, `Email Body` and `Status` = `Queued`.
   - Subject: `Expense Report from <NAME> – <dd/mm/yy> to <dd/mm/yy>`
   - Body (HTML, like timesheets): submitted by name and the signed-in user's email, pay period, total, number of receipts.
6. The flow takes over.

**My call:** if step 4 or 5 fails, set `Status` back to blank, delete the Excel attachment if it was added, and show the error. The report stays a draft with all entries and receipts, and the user can Submit again. Timesheets mark failures `Email failed` instead, but a timesheet can be re-entered from scratch and an expense report with receipts should not be.

### Excel template

- Convert `Templates/Expense Report Template (current).xlsm` to `Templates/xlsx/Expense Report Template.xlsx` with `scripts/Convert-Templates.ps1` (it currently only handles the three timesheet file names), and upload it to the `Timesheet Templates` library.
- Map the form fields (instructor name, pay period, the 13 rows: date, start, end, km, rate, mileage total, description, GST, cost, amount, and the totals) to cells by reading `xl/worksheets/sheet1.xml`. Keep the template's formulas and store cached results, as `fillTemplate.ts` does for timesheets, so previewers show totals.
- The hidden `Instructors` sheet feeds a dropdown in the template; write the name straight into the cell.
- Add a test that fills the template and checks the mapped cells, like `src/excel/fillTemplate.test.ts`.

## Left to the builder

- Exact cell addresses and number formats in the template.
- Where the code lives (a new submit function in `src/expenseReport/services/` next to `SharePointReportService`, or alongside `src/services/submit.ts`), as long as `src/expenseReport/` changes stay small.
- Wording of the banner, buttons and errors.
- README updates: Submit, the template upload, turning the flow on.

## Done when

- `npm test` passes, with new tests for the template fill and the submit sequence (success, failure at each step returning to draft).
- On the site: a draft can be saved, reopened and submitted; the item ends with the receipts and the `.xlsx` attached, the email columns filled and `Status` = `Queued`; with the flow on, accounts receive one email with all attachments and the item shows `Sent`; the report reopens read-only; a second report for the same pay period submits on its own.
