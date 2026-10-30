# Timesheet Web Part — Design / Build Prompt

Date: 2026-10-30
Status: draft for approval

This document is written so it can be handed to a builder (human or Claude) as the prompt for building the web part. Decisions marked **(user)** came from the business owner. Decisions marked **(call)** are the designer's defaults and can be overridden.

## 1. Intent

Peak Emergency pays its staff from three macro-enabled Excel timesheets in `Templates/`:

- `Admin Timesheet Template (current).xlsm`
- `Instructor Timesheet Template (current).xlsm`
- `Medic Timesheet Template (current).xlsm`

Today a staff member opens the template for their role, picks their name and pay period, fills in one row per work day, saves, and emails the workbook to **account@peakemergency.com**. Many people work in more than one role, so one person often submits several timesheet types.

Build a **SharePoint Framework (SPFx) web part** that replaces the manual workbook handling, while the accounts team keeps receiving the same Excel file by email.

## 2. Goals

1. The **logged-in user** fills in a timesheet in the browser. Their name comes from their Microsoft 365 profile, not a dropdown. **(user)**
2. The user picks which timesheet to fill in: Admin, Instructor or Medic. **(user, confirmed on playback)**
3. The entry screen is a spreadsheet grid using **Jspreadsheet CE** (MIT licence), styled to **look and feel like the Excel template** for that role: same column order, headings, header block, totals row, colours and borders as closely as practical. **(user)**
4. On submit, the web part produces an Excel attachment that **matches the template exactly** — same sheet, layout, cells, formatting — by filling in a copy of the real template, not by using the grid library's own export. **(user)** Library: **ExcelJS**. **(call)**
5. The attachment is emailed **from the user's own Outlook mailbox** (Microsoft Graph `/me/sendMail`) to account@peakemergency.com. **(user)**
6. Every submission is **recorded in a new SharePoint list**. **(user)**

## 3. Non-goals

- The Expense Report template (`Expense Report Template (current).xlsm`). Out of scope for now. **(user)**
- Approval workflow, manager sign-off, payroll system integration. **(call)**
- Saving drafts, or editing/resubmitting a timesheet after it is sent. A correction is a new submission. **(call)**
- Fixing the Admin column H data validation (it references `#REF!`). Leave it alone. **(user)**

## 4. Anti-goals (would count as failure even if it "works")

- The emailed workbook differs from what accounts receives today (different layout, missing formatting, moved cells).
- The email comes from a shared/service mailbox instead of the person submitting.
- A submission is emailed but not recorded, or recorded with no way to tell whether the email went out.

## 5. Pay rules **(user)**

| Timesheet | Pay basis |
|---|---|
| Admin | **By the hour.** Hours × hourly pay rate. |
| Instructor | **Daily rate.** Days (0.5 = half day, 1.0 = full day) × day rate. Plus admin hours × $25, travel $100, per diem $47. |
| Medic | **Daily rate.** One row = one day; line pay = the PSL-level day rate. Start/end times are recorded but do not change pay. Plus travel $100, per diem $47. |

The Medic template currently computes `H = F × G` (hours × rate). This is wrong under the daily-rate rule. **In the Medic attachment, column H for rows 8–20 must equal the day rate (G)** — e.g. write `=G8` in H8 — instead of hours × rate. The other Medic formula bugs listed in section 6 are also fixed in the attachment. Everything else in the templates stays as-is. **(user: daily rate and fix the bugs; call: exact formulas)**

## 6. Template cell map (source of truth)

The templates themselves are the source of truth; re-read them for styling. Each workbook has one visible sheet named after the template, protected (no password needed to write cells with ExcelJS). Dates display as `dd/mm/yy`.

### Admin — sheet `Admin Timesheet Template`

- Header: name `A4` (merged `A4:E4`), pay period From `G4`, To `I4`.
- Lines: rows **7–22** (16 rows).
  - `A` date · `C` start time · `D` end time · `E` deductions (hours; list `0.0, .5, 1.0, 1.25, 1.50, 1.75, 2.0`)
  - `F` total hours `=((D-C)*24-E)` · `G` pay rate (typed, $/hour) · `I` amount `=(G*F)`
- Totals row 25: `E25` deductions, `F25` hours, `I25` amount.

### Instructor — sheet `Instructor Timesheet Template`

- Header: name `A4` (merged `A4:E4`), From `J4`, To `M4`.
- Lines: rows **8–20** (13 rows).
  - `A` date · `C` course # (uppercase)
  - Instructional fees: `D` days (`0.5`, `1.0`) · `E` day rate (list from hidden `XFC2:XFC8`: 0, 150, 225, 280, 320, 340, 350) · `F` `=D*E`
  - Administrative fees: `G` hours (0.25–8.00 in 0.25 steps) · `H` rate (`25`) · `I` `=IF(OR(G="",H=""),"",G*H)`
  - Other fees: `J` travel time (`100`) · `K` per diem (`47`) · `L` `=SUM(J:K)`
  - `M` amount `=SUM(F,I,L)`
- Totals row 22: `D22`, `G22`, `J22`, `K22`, `M22`.

### Medic — sheet `Medic Timesheet Template`

- Header: name `A4` (merged `A4:E4`), From `H4`, To `K4`.
- Lines: rows **8–20** (13 rows).
  - `A` date · `C` client (uppercase) · `D` time start · `E` time end · `F` hours `=(E-D)*24`
  - `G` day rate (list from hidden `O6:O9` = `N6:N9`): $250 NUEC-A/AP (PSL 8), $350 PCP (PSL 9), $450 ACP (PSL 10), $550 CCP/MD (PSL 11) — role labels are in hidden `M6:M9`
  - `H` total — **override to `=G` per section 5**
  - Other fees: `I` travel time ($100) · `J` per diem ($47) · `K` `=SUM(I:J)`
  - `L` amount `=SUM(H,I,J)`
- Totals row 22: `F22`, `I22`, `J22`, `L22`.
- **Formula fixes to apply in the Medic attachment (user):**
  - `J22`: template has `=SUM(J8:K20)` (double-counts per diem and adds travel). Write `=SUM(J8:J20)`.
  - `K20`: template has no formula. Write `=SUM(I20:J20)`.
  - `F8:F20`: template `=(E-D)*24` goes negative for overnight shifts. Write `=MOD(E-D,1)*24` so the file matches the on-screen hours.

### Behaviour the macros enforce today (replicate in the web part)

- Name, pay period From and To are required. To cannot be before From.
- Pay period dates are free-entry (the templates use a pop-up calendar, not a fixed list of periods). Line dates are not checked against the pay period, same as today. **(call, from the templates)**
- Name stored in uppercase. Course # / client stored in uppercase.
- Fill the name as `LAST, FIRST` in uppercase, matching the existing lists (e.g. `NIKODYM, CAROLYN`). **(call)**

## 7. How it works

1. User opens the page, picks a timesheet type.
2. Grid shows the template's header block and a fixed number of line rows (16 Admin, 13 Instructor/Medic), with the same dropdown choices and live-calculated totals.
3. Validation on submit: required header fields, To ≥ From, at least one line with a date.
4. Submit:
   1. Create the SharePoint list item with status **Pending** (data saved first, so nothing is lost if email fails). **(call)**
   2. Load the template for that type, write header + lines into the cells in section 6, apply the Medic override, keep all formulas/formatting, set the workbook to recalculate on open.
   3. Send via Graph `POST /me/sendMail` with the file as a `fileAttachment`, `saveToSentItems: true`.
   4. Attach the generated file to the list item and set status **Sent** (or **Email failed** with the error text). The user sees a clear success or failure message.

## 8. Decided defaults **(call)** — override if wrong

- **Template storage:** one `.xlsx` copy of each template (macros removed via Excel "Save As .xlsx") in a SharePoint document library, e.g. `Timesheet Templates`. The web part loads it at submit time, so accounts can update a template without a redeploy. The library URL / file names are web part properties.
- **Attachment format:** `.xlsx` (no macros — they only helped data entry). File name: `<Type> Timesheet - <LAST, FIRST> - <yyyy-mm-dd>_<yyyy-mm-dd>.xlsx`.
- **Email:** To = web part property, default `account@peakemergency.com`. Subject: `<Type> Timesheet – <LAST, FIRST> – <From> to <To>`. Short body with the grand total.
- **SharePoint list** `Timesheet Submissions`: Title (file name), TimesheetType (choice), Employee (person), PayPeriodStart, PayPeriodEnd (dates), TotalAmount (currency), LineItems (multi-line JSON of the grid rows), Status (choice: Pending / Sent / Email failed), StatusDetail (text), plus the `.xlsx` as an item attachment. List item-level permissions: users read and edit only their own items; accounts/admins see all.
- **Provisioning:** the builder supplies a repeatable way to create the list and library (PnP PowerShell script or PnP provisioning template) — not created ad hoc in the web part.
- **Mailboxes:** every submitter has a licensed Exchange Online mailbox, so no fallback sender is needed. **(user)**
- **Graph permission:** `Mail.Send` (delegated) requested in `package-solution.json` `webApiPermissionRequests`; a tenant admin approves it once in SharePoint admin center → API access.
- **Overnight medic shifts** (end time earlier than start): show hours as `end − start + 24h` on screen. Pay is unaffected (daily rate).

## 9. Left to the builder

- SPFx version (use the current GA release and its supported React/Node versions), project structure, state management, styling approach.
- Jspreadsheet CE version and React wrapper; how to mimic template styling in the grid.
- ExcelJS details for loading a template and preserving styles, merged cells, data validation and protection; base64 encoding for the Graph attachment.
- Error handling detail, loading states, tests.

## 10. Acceptance

- For each of Admin, Instructor and Medic: fill in sample rows, submit, and confirm
  - email arrives at the configured address, sent from the submitting user, with one `.xlsx` attachment;
  - opening the attachment in Excel shows the original template layout and formatting with the entered values in the cells from section 6, and totals that match the web part's on-screen totals (Medic: line total = day rate);
  - a `Timesheet Submissions` item exists with status Sent and the same file attached.
- Missing name/dates or To < From blocks submit with a clear message.
- A user cannot see other users' submissions in the list; accounts can.
