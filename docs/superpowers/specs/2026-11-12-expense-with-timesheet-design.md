# Expense reports sent with the timesheet: design

Date: 2026-11-12

## Why

Accounts get timesheets and expense reports as separate emails and have to pair them up by hand. They want a person's expense report to come attached to that person's timesheet for the same pay period.

## What the client decided (2026-11-12)

- The match is automatic: same name, same pay period. The employee does not pick anything.
- A matching expense report goes with the timesheet even if it was already emailed on its own. Accounts may get it twice, which is fine.
- An expense report submitted after its timesheet is sent on its own, as it is today. Nothing is added to the timesheet that went out earlier.
- The separate expense report email stays as it is.

## Goals

- When a timesheet is submitted, find the submitted expense reports with the same name and pay period and copy their attachments (the filled Excel form and the receipts) onto the `Timesheet Submissions` item before it is set to `Queued`.
- The "Timesheet email" flow sends every attachment on the item, not only the first.
- The timesheet email body says how many expense reports came with it.

## Non-goals

- No change to the expense report form, its submit, or its flow.
- No link column between the two lists. The attachments are copies.
- No resend of timesheets already sent.
- No match on drafts. Only expense reports with Status `Queued`, `Sent` or `Email failed` count (they have their filled Excel attached).

## Anti-goals

- A timesheet must never fail to submit because of the expense report lookup. If the lookup or a copy fails, the timesheet is still sent without it, and the user is told the expense report was not included.
- Another person's expense report must never be attached. The match is exact on name and on both dates.

## How the match works

- Timesheet: `data.name` (the `Employees` Title, `LAST, FIRST`), `data.from`, `data.to` (ISO dates from the pay period picker).
- Expense report: `InstructorName` (same `Employees` Title), `FromDate`, `ToDate` columns. Both forms use the same `PayPeriodPicker`, so the dates match exactly when the pay period is the same.
- Query `ExpenseReports` for `InstructorName eq name` and `Status` in the three values above, then compare the date part of `FromDate` and `ToDate` in code (date-only fields come back as ISO date-times).
- Several matching reports: attach them all.
- `ExpenseReports` has item-level security, so the submitter only sees reports they created. A report someone else filled in for them is not found. Accepted (my call): in practice people fill in their own.

## Submit sequence (timesheet)

Today: create item (Pending), fill template, attach, set Queued.

New: create item (Pending), fill template, attach the timesheet, **find matching expense reports and copy each of their attachments onto the item**, set Queued.

- Copy = read the file bytes from the expense report attachment's `ServerRelativeUrl`, then `AttachmentFiles/add` on the timesheet item, keeping the file name.
- Two expense reports can carry a receipt with the same file name. A timesheet item cannot hold two attachments with one name, so prefix the second and later copies with the report Id (`<Id> - <name>`). My call.
- Lookup or copy failure: log it, skip the remaining expense files, still set Queued, and show a warning after submit ("Your timesheet was sent, but the expense report could not be included: ..."). `submitTimesheet` returns this as a warning, not an error.
- Email body gets a line `Expense reports attached: N` when N > 0.

## Flow change (manual, in Power Automate)

"Timesheet email" today sends `first(body('Get_attachments'))`. Rebuild steps 3 and 4 the way "Expense report email" does it: initialize an `Attachments` array, loop over `Get attachments`, get each file's content, append `{Name, ContentBytes}` with the whole content object, and send `variables('Attachments')`. Update `README.md` to match. This change must be live before the new web part version is deployed, or accounts get only the timesheet file.

## Where the code goes

- `src/services/submit.ts`: new optional dependency `expenses?: { find(name, from, to): Promise<{ id: number; files: { name: string; url: string }[] }[]>; copy(itemId, file, asName): Promise<void> }`, used between attach and Queued. Tests in `submit.test.ts` cover: no match, one match, two matches with a clashing file name, lookup failure still queues with a warning.
- `src/services/sharepoint.ts`: the SharePoint version of that dependency, using `SPHttpClient`.
- `src/webparts/timesheet/`: pass the dependency in, show the warning.
- Version bump in `config/package-solution.json`.

## Calls made after the builder check (2026-11-12, as built)

- **Match on `InstructorKey`, not the name text.** The timesheet's name is looked up in the `Employees` list to get the key. `InstructorName` is a copy taken when the report was saved, so a rename (Jay to Justin, Katie to Kathryn) would break a name match. If the name is not in the list, it falls back to `InstructorName`.
- **20 MB cap.** The expense files are read before the timesheet item is created. If they total over 20 MB, none are attached and the submitter gets a warning. This means one big report never makes the whole timesheet email fail.
- **All or nothing.** If a copy fails partway, the copies already added are deleted, the body line is removed, and the timesheet goes alone with a warning.
- **Two timesheets for one pay period** (say Instructor and Medic): the expense report goes with each of them. This is a default; the client hasn't decided.
- `SubmissionStore` gained `deleteAttachment` and `setEmailBody`. `SubmitResult` gained an optional `warning`, which the form shows after the success message.

## Left to the builder

- Exact wording of the warning and the email body line.
- Whether the copy streams bytes via `SPHttpClient` or `fetch`; either works on the same site.
