# peak-timesheets

SharePoint Framework web part for Peak Emergency timesheets. The user picks Admin, Instructor or Medic, picks their name from the `Employees` list, fills in a grid laid out like the Excel template, and submits. On submit the web part:

1. records the submission in the `Timesheet Submissions` list (status `Pending`), with the email's To, Subject and Body ready in hidden columns;
2. fills a copy of the real template from the `Timesheet Templates` library and attaches it to the list item;
3. sets the item to `Queued` (or to `Email failed` with the error text if step 2 fails).

A Power Automate flow then emails `Queued` items from the shared `schedule@peakemergencytraining.onmicrosoft.com` mailbox to accounts (default `accounting@peakemergency.com`) and sets `Sent` or `Email failed`. Submitters don't need a mailbox of their own; the body names the real sender and their email.

Design: `docs/superpowers/specs/2026-10-30-timesheet-webpart-design.md`.

## Prerequisites

- SPFx 1.24.0. Node.js 22 LTS (`>=22.14.0 <23`), see the [SPFx compatibility table](https://learn.microsoft.com/sharepoint/dev/spfx/compatibility).
- A SharePoint tenant app catalog.

Keep the repo out of OneDrive sync if you can: OneDrive locks files in `node_modules`, `jest-output` and `lib*`, which makes `npm install` very slow and can make builds fail with `EPERM`.

## Develop

```bash
npm install
npm test          # Heft build + Jest (unit tests for parsing, pay, validation, Excel fill, submission flow)
npm run start     # serve for the hosted workbench: <site>/_layouts/15/workbench.aspx
```

## Package and deploy

```bash
npm run build     # production build + package-solution
```

The package is `sharepoint/solution/peak-timesheets.sppkg`.

1. Upload the `.sppkg` to the tenant app catalog and deploy it.
2. Provision the list and the template library on the site that hosts the web part. Open any page of that site as a site owner, press F12, and paste `provisioning/provision-timesheets.js` into the console (set `ACCOUNTS_GROUP` at the top first). It uses SharePoint REST with your browser login, and is safe to rerun. It creates:
   - `Timesheet Submissions`, with attachments on and item-level security (users read and edit only their own items). The site Members group gets **Contribute** on this list, not Edit, because Edit includes *Manage Lists*, which bypasses item-level security. The accounts group (a SharePoint group) gets **Edit**, so it sees all items. People who submit timesheets must be site Members.
   - `Timesheet Templates`. Upload the three `.xlsx` files from `Templates/xlsx` into it. The script prints the library URL; put it in the web part's template library property if it is not `<site>/Timesheet Templates`.
   - `Pay Period Schedule`, seeded with the 2026 PP 1 to 2027 PP 26 periods. Rerun `provision-timesheets.js` to create and seed it. Add future years as rows in that list.
   - `Pay Rates`, seeded with the 29 staff from `docs/Staff PSL.pdf` (Instructor PSL, the six course rates AP, ASU, NUEC-A Full, NUEC-A Bridging, NUEC-I, NUEC-B, and Medic PSL) and the three hourly admin staff (Admin Rate). Salaried staff have no Admin Rate and don't submit admin timesheets. The timesheet fills each row's rate from the row of the person picked in the name dropdown, matched by name (`Rosie Trinh` in Pay Rates matches `TRINH, ROSIE` in Employees; a nickname in brackets is ignored): Admin Rate, the Medic PSL amount, or the instructor rate for the course type read from the course # (`AP-…`, `ASU-…`, `NUEC-A-…F`, `NUEC-A-…B`, `NUEC-I-…`, `NUEC-B-…`). A course # in any other format can't be submitted. Someone with no rate for a form sees "No pay rate is set up for NAME on this timesheet" and can't submit. Picking another name refills the rate on every entered row. Site Members can read every row (any of them can see everyone's rates); only the accounts group can edit. The **Employee** column is no longer used. Accounts keeps the rates up to date in that list.
3. Add the app to the site, then add the **Timesheet** web part to a page. Its properties are the recipient email, the list title and the template library URL (leave this blank for `<site>/Timesheet Templates`).

## Expense reports list

The Expense Report button uses the `ExpenseReports` list. Create it by pasting `provisioning/provision-expense-reports.js` into the browser console on `https://peakemergency.sharepoint.com` as a site owner (safe to rerun). It sets item-level security (users read and edit only their own reports), gives site Members Contribute and the `Payroll Owners` group Edit. Each user sees only reports they created. The name on the report is picked from the `Employees` list.

## Employees list

The name dropdown on every timesheet and the expense report lists the `Employees` list on the site (formerly `Instructors`): **Title** is the name as `LAST, FIRST`, **InstructorKey** a unique key (`instructor-NN`, saved with expense reports), and **Active** hides people who have left when set to No. Add new staff there, and give them a `Pay Rates` row with the same name.

## Email flow (Power Automate)

`schedule@` is a licensed user mailbox, so the flow's **Office 365 Outlook** connection signs in as `schedule@peakemergencytraining.onmicrosoft.com` and sends as itself; no Send As delegation is needed. The **SharePoint** connections can use any account with Edit on the list.

1. **Trigger:** SharePoint *When an item is created or modified*. Site `https://peakemergency.sharepoint.com`, list `Timesheet Submissions`. In the trigger's **Settings → Trigger conditions** add:
   `@equals(triggerOutputs()?['body/Status/Value'], 'Queued')`
2. **Get attachments:** SharePoint *Get attachments*, same site and list, Id = `ID` from the trigger.
3. **Get attachment content:** SharePoint *Get attachment content*, Id = `ID`, File Identifier = expression `first(body('Get_attachments'))?['Id']`.
4. **Send:** Office 365 Outlook *Send an email (V2)*, using the connection signed in as `schedule@`.
   - To: `Email To`; Subject: `Email Subject`; Body: `Email Body` (all from the trigger)
   - Attachments Name: expression `first(body('Get_attachments'))?['DisplayName']`; Content: `Attachment Content` from step 3
   - Set the attachment in the **classic designer** (turn off *New designer*). The new designer silently saves the attachment content as `null`. Check with **⋯ → Peek code** that `ContentBytes` is `@body('Get_attachment_content')`.
5. **Mark sent:** SharePoint *Send an HTTP request to SharePoint*. Method `PATCH`, Uri `_api/web/lists/getbytitle('Timesheet Submissions')/items(@{triggerBody()?['ID']})`, Headers `Accept` and `Content-Type` = `application/json;odata=nometadata`, `IF-MATCH` = `*`. Body: `{"Status": "Sent", "StatusDetail": ""}`
6. **Mark failed:** a second *Send an HTTP request to SharePoint* like step 5, with Body `{"Status": "Email failed", "StatusDetail": "Email flow failed: see the flow run history"}`. Open its **⋯ → Configure run after** and tick *has failed*, *has timed out* and *is skipped* on the step 4 action (skipped means step 2 or 3 failed), and set step 5 to run after step 4 *is successful* only.

The flow updates the item, which fires the trigger again, but the trigger condition skips anything not `Queued`.

### Expense report email flow

A second flow, "Expense report email" (made with *Save As* from "Timesheet email" and created turned off), built the same way on the `ExpenseReports` list. Rerun `provisioning/provision-expense-reports.js` first: it adds the same `Status`, `Status Detail`, `Email To`, `Email Subject` and `Email Body` columns. `Status` has no default, so saved reports stay blank until something sets them to `Queued`.

Not wired up yet: the web part does not set `Queued` or fill the email columns on expense reports. When it should (every save or a separate "send" step) is waiting on the client. To test the flow, fill the three email columns on a report by hand and set `Status` to `Queued`.

The list has item-level security, so the flow's SharePoint connection must be an account in `Payroll Owners` (or a site owner); a Member account only sees its own reports.

Differences from the timesheet flow, step by step:

1. **Trigger:** as step 1 above, list `ExpenseReports`, same trigger condition.
2. **Attachments variable:** *Initialize variable*, Name `Attachments`, Type Array, Value `[]`.
3. **Get attachments:** as step 2 above, list `ExpenseReports`. A report can have several receipts (and later the filled Excel form), so all of them are sent.
4. **Apply to each** over `body('Get_attachments')`, containing:
   - *Get attachment content*: Id = `ID` from the trigger, File Identifier = `items('Apply_to_each')?['Id']`
   - *Append to array variable* `Attachments`, Value (one expression; the classic designer rejects typed JSON mixed with tokens): `addProperty(addProperty(json('{}'), 'Name', items('Apply_to_each')?['DisplayName']), 'ContentBytes', body('Get_attachment_content')?['$content'])`
5. **Send:** as step 4 above, but switch the Attachments field to array input (the **T** icon in the classic designer) and set it to `variables('Attachments')`. Check with **Peek code** that `Attachments` is `@variables('Attachments')`.
6. **Mark sent / Mark failed:** as steps 5 and 6 above, Uri `_api/web/lists/getbytitle('ExpenseReports')/items(@{triggerBody()?['ID']})`. Configure *Mark failed* to run after *Send* has failed, timed out or is skipped.

A report with no receipts sends with no attachments.

## Updating a template

The web part writes values into fixed cells (see section 6 of the spec and `src/timesheet/definitions.ts`), so a template change must keep the sheet name, the cell positions and the file name.

1. Save the new macro-enabled workbook as `Templates/<Type> Timesheet Template (current).xlsm`.
2. Run `pwsh scripts/Convert-Templates.ps1`. It strips the macros and writes `Templates/xlsx/<Type> Timesheet Template.xlsx`.
3. Run `npm test`. The definitions test fails if headers or formula cells have moved.
4. Replace the file in the `Timesheet Templates` library, keeping the same name.

## Updating the template logo

Run `powershell -File scripts/Update-TemplateLogo.ps1`. It puts `src/webparts/timesheet/assets/peak-logo.jpg` into the three `Templates/xlsx` files, uncropped, at the logo's aspect ratio. After updating, re-upload the three files in `Templates/xlsx` to the Timesheet Templates library.
