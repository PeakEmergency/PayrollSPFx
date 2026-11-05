// Creates the Timesheet Submissions list, the Timesheet Templates library and the seeded Pay Period Schedule and Pay Rates lists with SharePoint REST.
// Run it in the browser console (F12) on a page of the target site, signed in as a site owner.
// Safe to rerun: anything that already exists is left alone.
// Afterwards, drag the three files from Templates/xlsx into the Timesheet Templates library.
await (async () => {
  const ACCOUNTS_GROUP = 'Payroll Owners'; // SharePoint group that sees every submission
  const LIST_TITLE = 'Timesheet Submissions';
  const LIBRARY_TITLE = 'Timesheet Templates';

  const verbose = 'application/json;odata=verbose';
  // Modern pages and the workbench have no _spPageContextInfo, so take the site from the page address
  // (everything before /_layouts/, /SitePages/, /Lists/ ...) and let contextinfo return the exact web URL.
  const here = (location.origin + location.pathname)
    .split(/\/(?:_layouts|SitePages|Lists|Shared%20Documents|Forms)\//i)[0]
    .replace(/\/$/, '');
  const ctxRes = await fetch(`${here}/_api/contextinfo`, { method: 'POST', headers: { Accept: verbose } });
  if (!ctxRes.ok) throw new Error(`Couldn't read the site from ${here}: ${ctxRes.status}. Run this on a page of the target site.`);
  const { WebFullUrl: web, FormDigestValue: digest } = (await ctxRes.json()).d.GetContextWebInformation;
  console.log(`Site: ${web}`);

  async function sp(path, { method = 'GET', body, headers = {} } = {}) {
    const res = await fetch(`${web}/_api/${path}`, {
      method,
      headers: { Accept: verbose, 'Content-Type': verbose, 'X-RequestDigest': digest, ...headers },
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`${method} ${path}: ${res.status} ${await res.text()}`);
    const text = await res.text();
    return text ? JSON.parse(text).d : null;
  }
  const list = `web/lists/getbytitle('${LIST_TITLE}')`;

  // --- Submissions list ---
  if (await sp(list)) console.log(`List '${LIST_TITLE}' exists`);
  else {
    await sp('web/lists', {
      method: 'POST',
      body: { __metadata: { type: 'SP.List' }, Title: LIST_TITLE, BaseTemplate: 100, EnableAttachments: true, EnableVersioning: true }
    });
    console.log(`Created list '${LIST_TITLE}'`);
  }

  const fields = [
    ['TimesheetType', '<Field Type="Choice" Name="TimesheetType" StaticName="TimesheetType" DisplayName="Timesheet Type" Required="TRUE"><CHOICES><CHOICE>Admin</CHOICE><CHOICE>Instructor</CHOICE><CHOICE>Medic</CHOICE></CHOICES></Field>'],
    ['Employee', '<Field Type="User" Name="Employee" StaticName="Employee" DisplayName="Employee" UserSelectionMode="PeopleOnly" Required="TRUE" />'],
    ['PayPeriodStart', '<Field Type="DateTime" Name="PayPeriodStart" StaticName="PayPeriodStart" DisplayName="Pay Period Start" Format="DateOnly" />'],
    ['PayPeriodEnd', '<Field Type="DateTime" Name="PayPeriodEnd" StaticName="PayPeriodEnd" DisplayName="Pay Period End" Format="DateOnly" />'],
    ['TotalAmount', '<Field Type="Currency" Name="TotalAmount" StaticName="TotalAmount" DisplayName="Total Amount" LCID="1033" Decimals="2" />'],
    ['LineItems', '<Field Type="Note" Name="LineItems" StaticName="LineItems" DisplayName="Line Items" RichText="FALSE" NumLines="6" />'],
    ['Status', '<Field Type="Choice" Name="Status" StaticName="Status" DisplayName="Status"><Default>Pending</Default><CHOICES><CHOICE>Pending</CHOICE><CHOICE>Queued</CHOICE><CHOICE>Sent</CHOICE><CHOICE>Email failed</CHOICE></CHOICES></Field>'],
    ['StatusDetail', '<Field Type="Note" Name="StatusDetail" StaticName="StatusDetail" DisplayName="Status Detail" RichText="FALSE" NumLines="3" />'],
    // Read by the Power Automate flow that sends the email.
    ['EmailTo', '<Field Type="Text" Name="EmailTo" StaticName="EmailTo" DisplayName="Email To" />'],
    ['EmailSubject', '<Field Type="Text" Name="EmailSubject" StaticName="EmailSubject" DisplayName="Email Subject" MaxLength="255" />'],
    ['EmailBody', '<Field Type="Note" Name="EmailBody" StaticName="EmailBody" DisplayName="Email Body" RichText="FALSE" NumLines="6" />']
  ];
  const hiddenFromView = ['LineItems', 'EmailTo', 'EmailSubject', 'EmailBody'];
  for (const [name, xml] of fields) {
    // getbyinternalnameortitle answers 400 (not 404) for a missing field, so filter instead.
    if ((await sp(`${list}/fields?$filter=InternalName eq '${name}'&$select=Id`)).results.length) { console.log(`  field ${name} exists`); continue; }
    // AddToDefaultContentType (1) | AddFieldInternalNameHint (8) | AddFieldToDefaultView (16).
    const options = hiddenFromView.includes(name) ? 9 : 25;
    await sp(`${list}/fields/createfieldasxml`, {
      method: 'POST',
      body: { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: xml, Options: options } }
    });
    console.log(`  added field ${name}`);
  }

  // Lists made before the email flow lack the Queued status.
  await sp(`${list}/fields/getbyinternalnameortitle('Status')`, {
    method: 'POST',
    headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    body: { __metadata: { type: 'SP.FieldChoice' }, Choices: { __metadata: { type: 'Collection(Edm.String)' }, results: ['Pending', 'Queued', 'Sent', 'Email failed'] } }
  });
  console.log('Status choices: Pending, Queued, Sent, Email failed');

  // --- Item-level security: users read and edit only their own items ---
  await sp(list, {
    method: 'POST',
    headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    body: { __metadata: { type: 'SP.List' }, ReadSecurity: 2, WriteSecurity: 2, EnableAttachments: true }
  });
  console.log('Item-level security on');

  // Anyone with Manage Lists (the Edit role) bypasses item-level security.
  // Unique permissions: site Members drop to Contribute on this list, the accounts group gets Edit.
  const roleId = async (name) => (await sp(`web/roledefinitions/getbyname('${name}')`)).Id;
  const edit = await roleId('Edit');
  const contribute = await roleId('Contribute');
  const { HasUniqueRoleAssignments } = await sp(`${list}?$select=HasUniqueRoleAssignments`);
  if (!HasUniqueRoleAssignments) {
    await sp(`${list}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`, { method: 'POST' });
    const members = await sp('web/AssociatedMemberGroup');
    await sp(`${list}/roleassignments/removeroleassignment(principalid=${members.Id},roledefid=${edit})`, { method: 'POST' });
    await sp(`${list}/roleassignments/addroleassignment(principalid=${members.Id},roledefid=${contribute})`, { method: 'POST' });
    console.log(`List permissions: '${members.Title}' set to Contribute`);
  }
  const [accounts] = (await sp(`web/sitegroups?$filter=Title eq '${ACCOUNTS_GROUP}'&$select=Id`)).results;
  if (!accounts) throw new Error(`SharePoint group '${ACCOUNTS_GROUP}' not found on this site`);
  await sp(`${list}/roleassignments/addroleassignment(principalid=${accounts.Id},roledefid=${edit})`, { method: 'POST' });
  console.log(`List permissions: '${ACCOUNTS_GROUP}' has Edit`);

  // --- Template library ---
  const lib = `web/lists/getbytitle('${LIBRARY_TITLE}')`;
  if (await sp(lib)) console.log(`Library '${LIBRARY_TITLE}' exists`);
  else {
    await sp('web/lists', {
      method: 'POST',
      body: { __metadata: { type: 'SP.List' }, Title: LIBRARY_TITLE, BaseTemplate: 101, EnableVersioning: true }
    });
    console.log(`Created library '${LIBRARY_TITLE}'`);
  }
  // --- Pay Period Schedule list: 52 biweekly periods starting 2026-01-01 (2026 PP 1 to 2027 PP 26) ---
  const PERIOD_LIST = 'Pay Period Schedule';
  const periods = `web/lists/getbytitle('${PERIOD_LIST}')`;
  if (await sp(periods)) console.log(`List '${PERIOD_LIST}' exists`);
  else {
    await sp('web/lists', { method: 'POST', body: { __metadata: { type: 'SP.List' }, Title: PERIOD_LIST, BaseTemplate: 100 } });
    console.log(`Created list '${PERIOD_LIST}'`);
  }
  const dateField = (name, label) => [name, `<Field Type="DateTime" Name="${name}" StaticName="${name}" DisplayName="${label}" Format="DateOnly" />`];
  const periodFields = [
    ['Year', '<Field Type="Number" Name="Year" StaticName="Year" DisplayName="Year" Decimals="0" Required="TRUE" />'],
    ['PeriodNumber', '<Field Type="Number" Name="PeriodNumber" StaticName="PeriodNumber" DisplayName="Period Number" Decimals="0" Required="TRUE" />'],
    dateField('FromDate', 'From Date'),
    dateField('ToDate', 'To Date'),
    dateField('PayrollRun', 'Payroll Run')
  ];
  for (const [name, xml] of periodFields) {
    if ((await sp(`${periods}/fields?$filter=InternalName eq '${name}'&$select=Id`)).results.length) { console.log(`  field ${name} exists`); continue; }
    await sp(`${periods}/fields/createfieldasxml`, {
      method: 'POST',
      body: { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: xml, Options: 25 } }
    });
    console.log(`  added field ${name}`);
  }
  const existing = (await sp(`${periods}/items?$select=Year,PeriodNumber&$top=5000`)).results.map((r) => `${r.Year}-${r.PeriodNumber}`);
  // Noon UTC keeps a date-only field on the same calendar day in any North American time zone.
  const iso = (base, days) => new Date(base + days * 86400000).toISOString().slice(0, 10) + 'T12:00:00Z';
  const base = Date.UTC(2026, 0, 1);
  const itemType = (await sp(`${periods}?$select=ListItemEntityTypeFullName`)).ListItemEntityTypeFullName;
  let seeded = 0;
  for (let k = 0; k < 52; k++) {
    const Year = k < 26 ? 2026 : 2027;
    const PeriodNumber = (k % 26) + 1;
    if (existing.includes(`${Year}-${PeriodNumber}`)) continue;
    await sp(`${periods}/items`, {
      method: 'POST',
      body: {
        __metadata: { type: itemType },
        Title: `${Year}-${String(PeriodNumber).padStart(2, '0')}`,
        Year,
        PeriodNumber,
        FromDate: iso(base, 14 * k),
        ToDate: iso(base, 14 * k + 13),
        PayrollRun: iso(base, 14 * k + 14)
      }
    });
    seeded++;
  }
  console.log(`Pay periods: ${seeded} added, ${52 - seeded} already there. Add future years as rows in '${PERIOD_LIST}'.`);

  // --- Pay Rates list: one row per person, seeded from docs/Staff PSL.pdf ---
  // Members read every row (the timesheet picks the person by name from the Employees list); only Payroll Owners edit.
  const RATE_LIST = 'Pay Rates';
  const rates = `web/lists/getbytitle('${RATE_LIST}')`;
  if (await sp(rates)) console.log(`List '${RATE_LIST}' exists`);
  else {
    await sp('web/lists', { method: 'POST', body: { __metadata: { type: 'SP.List' }, Title: RATE_LIST, BaseTemplate: 100, EnableVersioning: true } });
    console.log(`Created list '${RATE_LIST}'`);
  }
  const money = (name, label) => [name, `<Field Type="Currency" Name="${name}" StaticName="${name}" DisplayName="${label}" LCID="1033" Decimals="2" />`];
  const rateCols = ['RateAP', 'RateASU', 'RateNUECAFull', 'RateNUECABridging', 'RateNUECI', 'RateNUECB'];
  const rateFields = [
    ['Employee', '<Field Type="User" Name="Employee" StaticName="Employee" DisplayName="Employee" UserSelectionMode="PeopleOnly" />'],
    ['InstructorPSL', '<Field Type="Text" Name="InstructorPSL" StaticName="InstructorPSL" DisplayName="Instructor PSL" MaxLength="10" />'],
    money('RateAP', 'AP'),
    money('RateASU', 'ASU'),
    money('RateNUECAFull', 'NUEC-A Full'),
    money('RateNUECABridging', 'NUEC-A Bridging'),
    money('RateNUECI', 'NUEC-I'),
    money('RateNUECB', 'NUEC-B'),
    ['MedicPSL', '<Field Type="Number" Name="MedicPSL" StaticName="MedicPSL" DisplayName="Medic PSL" Decimals="0" Min="8" Max="11" />'],
    money('AdminRate', 'Admin Rate')
  ];
  for (const [name, xml] of rateFields) {
    if ((await sp(`${rates}/fields?$filter=InternalName eq '${name}'&$select=Id`)).results.length) { console.log(`  field ${name} exists`); continue; }
    await sp(`${rates}/fields/createfieldasxml`, {
      method: 'POST',
      body: { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: xml, Options: 25 } }
    });
    console.log(`  added field ${name}`);
  }

  // [Name, Instructor PSL, AP, ASU, NUEC-A Full, NUEC-A Bridging, NUEC-I, NUEC-B, Medic PSL]
  const same = (name, psl, rate, medic) => [name, psl, rate, rate, rate, rate, rate, rate, medic];
  const staff = [
    ['Ben Davies', '2/3', 225, 225, 150, 150, 225, 225, 8],
    ['Carolyn Nikodym', '', null, null, null, null, null, null, 9],
    same('Dallas Flockhart', '5', 320, 10),
    same('Doug Massel', '5', 320, 9),
    same('Evan Korpan', '3', 225, 10),
    same("Erin O'Gorman", '2', 150, 10),
    same('Fiona Pickett', '5', 320, 8),
    same('Jacob Hutton', '5', 320, 9),
    ['Justin Darbyshire', '2/3', 225, 225, 150, 150, 150, 225, 8],
    same('Jesse Hoevenaars', '7', 350, 9), // sheet said 6; client confirmed PSL 7 (2026-11-03)
    same('John Baril', '5', 320, 9),
    same('Josie Hartlin', '4', 280, 8),
    same("Kathryn O'Connor", '5', 320, 8),
    same('Kathryn Robinson', '4', 280, 8),
    same('Kennedy Schaefer', '2', 150, 8),
    same('Lindsay Eastwood', '5', 320, 9),
    same('Lucy McCart', '2', 150, 8),
    same('Lynette Hleck', '5', 320, 9),
    same('Maximilian Burko', '4', 280, 9),
    same('Melissa Huntley', '5', 320, 8),
    same('Mike Bowman', '5', 320, 9),
    same('Riley Adams', '2', 150, 8),
    same('Richard Bailey', '1', 0, 8),
    ['Ross Goddard', '2/3', 225, 225, 150, 225, 225, 225, 9],
    same('Shelly Hicks', '4', 280, 9),
    same('Simon Piney', '5', 320, 8),
    same('Tabea Stephan', '3', 225, 8),
    same('Tytan Wheeler', '6', 340, 8),
    same('William Geise', '6', 340, 8)
  ];
  // Hourly admin staff (client, 2026-11-03): [Name, Admin Rate]. Lucy McCart, Carolyn Nikodym and Jeffrey Burko
  // are salaried and submit no admin timesheet, so they have no Admin Rate.
  const adminStaff = [
    ['Emma Holden-Hindley', 30],
    ['Jennifer Treend', 25],
    ['Rosie Trinh', 24]
  ];

  // Finds the account by exact display name through the people picker (covers people who never visited the site).
  async function findUser(name) {
    const r = await sp('SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser', {
      method: 'POST',
      body: { queryParams: { __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' }, QueryString: name, PrincipalSource: 15, PrincipalType: 1, AllowEmailAddresses: false, MaximumEntitySuggestions: 10 } }
    });
    const hits = JSON.parse(r.ClientPeoplePickerSearchUser).filter((e) => (e.DisplayText || '').toLowerCase() === name.toLowerCase());
    if (hits.length !== 1) return { error: hits.length ? `${hits.length} accounts match` : 'no account found' };
    return { id: (await sp('web/ensureuser', { method: 'POST', body: { logonName: hits[0].Key } })).Id };
  }

  const rateType = (await sp(`${rates}?$select=ListItemEntityTypeFullName`)).ListItemEntityTypeFullName;
  const haveRates = (await sp(`${rates}/items?$select=Title&$top=5000`)).results.map((r) => r.Title.toLowerCase());
  let rateAdded = 0;
  const rows = [
    ...staff.map(([name, psl, ...rest]) => {
      const fields = { InstructorPSL: psl, MedicPSL: rest[6] };
      rateCols.forEach((col, i) => { if (rest[i] !== null) fields[col] = rest[i]; });
      return [name, fields];
    }),
    ...adminStaff.map(([name, rate]) => [name, { AdminRate: rate }])
  ];
  for (const [name, fields] of rows) {
    if (haveRates.includes(name.toLowerCase())) continue;
    const body = { __metadata: { type: rateType }, Title: name, ...fields };
    const user = await findUser(name);
    if (user.id) body.EmployeeId = user.id;
    else console.warn(`Pay Rates: ${user.error} for ${name}, set Employee by hand and rerun`);
    await sp(`${rates}/items`, { method: 'POST', body });
    rateAdded++;
  }
  console.log(`Pay rates: ${rateAdded} added, ${rows.length - rateAdded} already there`);

  // List: Members read only, Payroll Owners edit.
  const read = await roleId('Read');
  const members = await sp('web/AssociatedMemberGroup');
  if (!(await sp(`${rates}?$select=HasUniqueRoleAssignments`)).HasUniqueRoleAssignments) {
    await sp(`${rates}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`, { method: 'POST' });
    await sp(`${rates}/roleassignments/removeroleassignment(principalid=${members.Id},roledefid=${edit})`, { method: 'POST' });
    await sp(`${rates}/roleassignments/addroleassignment(principalid=${members.Id},roledefid=${read})`, { method: 'POST' });
  }
  await sp(`${rates}/roleassignments/addroleassignment(principalid=${accounts.Id},roledefid=${edit})`, { method: 'POST' });

  // Items inherit the list: the timesheet's Employee dropdown loads anyone's rates, so every Member reads every row
  // (client, 2026-11-04). Earlier runs gave each row its own permissions; this undoes that.
  let reset = 0;
  for (const item of (await sp(`${rates}/items?$select=Id,HasUniqueRoleAssignments&$top=5000`)).results) {
    if (!item.HasUniqueRoleAssignments) continue;
    await sp(`${rates}/items(${item.Id})/resetroleinheritance`, { method: 'POST' });
    reset++;
  }
  console.log(`Pay Rates permissions set. ${reset} row(s) reset to the list permissions.`);

  const folder = (await sp(`${lib}/RootFolder?$select=ServerRelativeUrl`)).ServerRelativeUrl;
  console.log(`Done. Upload Templates/xlsx/*.xlsx to ${location.origin}${folder}`);
  console.log(`Template library URL for the web part: ${folder}`);
})().catch((e) => console.error('Provisioning failed:', e.message));
