# Template changes from jeff's review (November 2026)

## Why

Rosie Trinh (PEAK) walked Walter Johnson through jeff's requested changes to the Admin, Instructor, Medic and Expense Report forms. The goals are:

- Staff can't type the wrong pay period or the wrong rate.
- Nobody can submit for someone else. The name always comes from the signed-in user. Walter advised against jeff's name dropdown for this reason.
- The headings match what accounts expects.

The emailed Excel files must show the same changes as the web form, because accounts works from those files.

## Starting point

This builds on the uncommitted work that folded the Expense Report into this web part (`src/expenseReport/`, plus the "Expense Report" button left of Admin, Instructor and Medic). That work already covers "Expense report becomes its own tab". Commit it before starting this work.

## Goals

### All forms (Admin, Instructor, Medic, Expense Report)

1. **Pay period dropdown.** Replace the From and To date inputs with a "Pay period" dropdown. Choosing a period fills From and To, which are shown but can't be edited.
   - The data comes from a new SharePoint list, **Pay Period Schedule**, with these fields:
     - Year (number)
     - Period number (number, 1–26)
     - From (date)
     - To (date)
     - Payroll Run (date)
   - From and To are full dates, because periods can cross a year boundary. 2027 period 1 runs December 31, 2026 to January 13, 2027.
   - Seed the list with the 2026 and 2027 schedules from `docs/Payperiod Schedule 2026.pdf` and `docs/Payperiod Schedule 2027.pdf` (26 rows each). Extend `provisioning/provision-timesheets.js` to create and seed the list, staying safe to rerun.
   - Example: 2026 period 23 is November 5 to November 18.
2. **Rate from the signed-in user.** When the user starts a row, the rate cell for that row fills in from a new SharePoint list, **Pay Rates**, for the signed-in person and form type. Walter will provide the list contents once jeff sends the pay scale list.
   - The PSL is chosen **per shift**, one per row (Walter). The row starts at the person's level from Pay Rates, and the user can change it for a shift paid at a different level.
   - The PSL cell shows the **level** (e.g. "PSL 4"), not the dollar amount, on screen and in the emailed Excel (Walter). The dollar value behind each level is used for totals. In Excel this needs a level-to-amount lookup table and matching formulas.
3. **Name** stays auto-filled from the signed-in user on every form. There is no name dropdown anywhere.

### Instructor timesheet

4. Rename `DAY(s)` to `DAYS`.
5. Rename the instructional `RATE` to `PAY SCALE LEVEL (PSL)`.
6. The per diem dropdown offers **One day = 47** and **Half day = 23.5**. Today it offers only 47.

### Medic timesheet

7. Remove `TIME START`, `TIME END` and `HOURS(s)`. Add a `DAYS` dropdown with the values 0.5 and 1.0.
8. Rename `RATE` to `PSL`. The medic total becomes `DAYS × PSL rate`. Today it is the flat rate.
9. Use the same per diem dropdown as Instructor (47 / 23.5).

### Admin timesheet

10. Rename `TOTAL HOURS(s)` to `TOTAL HOURS`. Keep it plural.
11. Show `PAY RATE` on two lines (PAY over RATE). `END TIME` is already split this way.

### Expense Report

12. Remove the Instructor dropdown and use the signed-in user's name, as the timesheets do.
13. "Saved reports" lists only the signed-in user's own reports.
14. Use the timesheet logo (`src/webparts/timesheet/assets/peak-logo.jpg`), which reads just "PEAK" with no "paramedic services".
15. Use a date picker for the line dates.
16. In the totals row, unbold the kilometres total and the cost (excl GST) total.
17. Center all table headings.

### Shared look (Walter)

18. The Expense Report's look becomes the common theme. Restyle the Admin, Instructor and Medic timesheets to look like the Expense Report by editing `src/webparts/timesheet/components/TimesheetForm.module.scss`, and the grid styling around `TimesheetGrid.tsx` where needed. Match:
    - fonts and the red title with the logo
    - the name and pay period header
    - the button style
    - table borders, header cells and spacing
    - the red totals row
    - the error and success messages

    Keep the current structure, including the Jspreadsheet grid. This is a restyle, not a rebuild. The source of truth is `src/expenseReport/components/ExpenseReport.module.scss`. This replaces the earlier goal that the web form copies the Excel template look. The emailed `.xlsx` files keep the template look.

### Excel output

19. The `.xlsx` templates in `Templates/xlsx/` get the same heading, column and per diem changes as the web form, with the From and To dates filled from the chosen period. The Medic DAYS column replaces the time columns and its formulas change to match. Upload the updated files to the `Timesheet Templates` library.

## Non-goals

- The Expense Report still doesn't produce Excel or send email. It saves to the `ExpenseReports` list as today.
- No accounts view of everyone's expense reports.
- No admin screen for editing pay periods or pay rates. They are edited directly in the SharePoint lists.

## Anti-goals (failure even if it works)

- A user can enter or change another person's name, rate or pay period dates.
- The emailed Excel shows old headings or columns that don't match the form.

## Constraints

- SPFx 1.24, React 18, ES5 target. See `CLAUDE.md` for commands and the OneDrive `EPERM` workaround.
- The only Graph permission is `User.Read`. Read lists with `SPHttpClient`.
- Jest runs on `lib-commonjs` in the Node environment (`config/jest.config.json`).

## Decisions made in conversation

- The Excel templates change too (Walter).
- The rate fills in automatically when the row is entered, from a list Walter will provide (Walter).
- The Expense Report shows only the user's own reports (Walter).
- The pay period dates are disabled and populated from the new list (Walter).
- PSL is one per shift (per row), and the cell shows the level, not dollars (Walter).
- There are no saved expense reports yet, so nothing needs migrating (Walter). The Expense Report can drop the `Instructors` list and change its stored fields freely.

## Calls made by Claude (check these)

- **No period is preselected.** Submit and Save stay disabled until one is chosen, as with today's blank dates.
- **The dropdown lists periods whose From date is on or before today, newest first**, labelled like `2026 – PP 23 (Nov 5 – Nov 18)`. Future periods are hidden.
- **The timesheet logo is reused** for the Expense Report rather than making a new image.
- **The Payroll Run date is stored but not shown.**
- **Only the main rate comes from Pay Rates.** That is the Admin pay rate, the Instructor instructional PSL and the Medic PSL. The Instructor admin-fee rate (25) and travel (100) stay as they are. The Expense Report km rate stays as the web part setting (0.48 default) and is not per person.
- **Locking happens in the form only.** The "Timesheet email" flow doesn't re-check name, dates or rate against the lists. This matches today's trust model, where the web part builds the list item and the `.xlsx`. The anti-goal means "through the form".
- **The Pay Rates list shape waits for jeff's list.** Proposed shape: one row per person, with Person (person field), Admin pay rate, and the default Instructor PSL and Medic PSL that pre-fill each row. If a person has no row for that form, show a clear "No pay rate set up for you; contact accounts" message and block submit.

## Blocked

- **Item 2, pay rates:** blocked until jeff's pay scale list arrives through Rosie Trinh. Everything else can be built now, with the rate cell staying a dropdown until the list exists.

## Left to the builder

- How to edit the `.xlsx` templates: edit the files, or rewrite headings and formulas at fill time in `fillTemplate.ts`, or both. Excel COM hangs on the Medic file, so prefer zip or ExcelJS edits (see `scripts/Convert-Templates.ps1`).
- How the Expense Report stores the user. There is no old data to keep compatible.
- How "own reports only" is enforced: an `Author` filter, item-level list permissions, or both. Prefer list permissions (ReadSecurity/WriteSecurity = 2, as `provision-timesheets.js` sets on `Timesheet Submissions`) so it isn't only a UI filter. Make sure accounts can still read the list if needed.
- Mapping the period and rate lists into tests (pure functions in `src/timesheet/` where possible).
