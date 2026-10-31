# peak-timesheets

SharePoint Framework web part for Peak Emergency timesheets. The signed-in user picks Admin, Instructor or Medic, fills in a grid laid out like the Excel template, and submits. On submit the web part:

1. records the submission in the `Timesheet Submissions` list (status `Pending`), with the email's To, Subject and Body ready in hidden columns;
2. fills a copy of the real template from the `Timesheet Templates` library and attaches it to the list item;
3. sets the item to `Queued` (or to `Email failed` with the error text if step 2 fails).

A Power Automate flow then emails `Queued` items from the shared `schedule@peakemergencytraining.onmicrosoft.com` mailbox to accounts (default `accounting@peakemergency.com`) and sets `Sent` or `Email failed`. Submitters don't need a mailbox of their own; the body names the real sender and their email.

Design: `docs/superpowers/specs/2026-10-30-timesheet-webpart-design.md`.

## Prerequisites

- SPFx 1.24.0. Node.js 22 LTS (`>=22.14.0 <23`), see the [SPFx compatibility table](https://learn.microsoft.com/sharepoint/dev/spfx/compatibility).
- A SharePoint tenant app catalog, and a tenant admin to approve the Graph permissions.

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
2. Approve the API access: SharePoint admin center → **Advanced** → **API access**. Approve `Microsoft Graph` · `User.Read` (used to read the submitter's name).
3. Provision the list and the template library on the site that hosts the web part. Open any page of that site as a site owner, press F12, and paste `provisioning/provision-timesheets.js` into the console (set `ACCOUNTS_GROUP` at the top first). It uses SharePoint REST with your browser login, and is safe to rerun. It creates:
   - `Timesheet Submissions`, with attachments on and item-level security (users read and edit only their own items). The site Members group gets **Contribute** on this list, not Edit, because Edit includes *Manage Lists*, which bypasses item-level security. The accounts group (a SharePoint group) gets **Edit**, so it sees all items. People who submit timesheets must be site Members.
   - `Timesheet Templates`. Upload the three `.xlsx` files from `Templates/xlsx` into it. The script prints the library URL; put it in the web part's template library property if it is not `<site>/Timesheet Templates`.
4. Add the app to the site, then add the **Timesheet** web part to a page. Its properties are the recipient email, the list title and the template library URL (leave this blank for `<site>/Timesheet Templates`).

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

## Updating a template

The web part writes values into fixed cells (see section 6 of the spec and `src/timesheet/definitions.ts`), so a template change must keep the sheet name, the cell positions and the file name.

1. Save the new macro-enabled workbook as `Templates/<Type> Timesheet Template (current).xlsm`.
2. Run `pwsh scripts/Convert-Templates.ps1`. It strips the macros and writes `Templates/xlsx/<Type> Timesheet Template.xlsx`.
3. Run `npm test`. The definitions test fails if headers or formula cells have moved.
4. Replace the file in the `Timesheet Templates` library, keeping the same name.
