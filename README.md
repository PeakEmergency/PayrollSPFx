# peak-timesheets

SharePoint Framework web part for Peak Emergency timesheets. The signed-in user picks Admin, Instructor or Medic, fills in a grid laid out like the Excel template, and submits. On submit the web part:

1. records the submission in the `Timesheet Submissions` list (status `Pending`);
2. fills a copy of the real template from the `Timesheet Templates` library and attaches it to the list item;
3. emails the `.xlsx` from the user's own mailbox (Microsoft Graph `/me/sendMail`) to accounts (default `account@peakemergency.com`);
4. sets the item to `Sent`, or to `Email failed` with the error text if anything after step 1 fails.

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
2. Approve the API access: SharePoint admin center → **Advanced** → **API access**. Approve `Microsoft Graph` · `Mail.Send` and `Microsoft Graph` · `User.Read`.
3. Provision the list and the template library on the site that hosts the web part. Open any page of that site as a site owner, press F12, and paste `provisioning/provision-timesheets.js` into the console (set `ACCOUNTS_GROUP` at the top first). It uses SharePoint REST with your browser login, and is safe to rerun. It creates:
   - `Timesheet Submissions`, with attachments on and item-level security (users read and edit only their own items). The site Members group gets **Contribute** on this list, not Edit, because Edit includes *Manage Lists*, which bypasses item-level security. The accounts group (a SharePoint group) gets **Edit**, so it sees all items. People who submit timesheets must be site Members.
   - `Timesheet Templates`. Upload the three `.xlsx` files from `Templates/xlsx` into it. The script prints the library URL; put it in the web part's template library property if it is not `<site>/Timesheet Templates`.
4. Add the app to the site, then add the **Timesheet** web part to a page. Its properties are the recipient email, the list title and the template library URL (leave this blank for `<site>/Timesheet Templates`).

## Updating a template

The web part writes values into fixed cells (see section 6 of the spec and `src/timesheet/definitions.ts`), so a template change must keep the sheet name, the cell positions and the file name.

1. Save the new macro-enabled workbook as `Templates/<Type> Timesheet Template (current).xlsm`.
2. Run `pwsh scripts/Convert-Templates.ps1`. It strips the macros and writes `Templates/xlsx/<Type> Timesheet Template.xlsx`.
3. Run `npm test`. The definitions test fails if headers or formula cells have moved.
4. Replace the file in the `Timesheet Templates` library, keeping the same name.
