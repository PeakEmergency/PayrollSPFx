// Creates the Timesheet Submissions list and the Timesheet Templates library with SharePoint REST.
// Run it in the browser console (F12) on a page of the target site, signed in as a site owner.
// Safe to rerun: anything that already exists is left alone.
// Afterwards, drag the three files from Templates/xlsx into the Timesheet Templates library.
(async () => {
  const ACCOUNTS_GROUP = 'Payroll Owners'; // SharePoint group that sees every submission
  const LIST_TITLE = 'Timesheet Submissions';
  const LIBRARY_TITLE = 'Timesheet Templates';

  const web = _spPageContextInfo.webAbsoluteUrl;
  const verbose = 'application/json;odata=verbose';
  const digest = (await (await fetch(`${web}/_api/contextinfo`, { method: 'POST', headers: { Accept: verbose } })).json()).d
    .GetContextWebInformation.FormDigestValue;

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
    ['Status', '<Field Type="Choice" Name="Status" StaticName="Status" DisplayName="Status"><Default>Pending</Default><CHOICES><CHOICE>Pending</CHOICE><CHOICE>Sent</CHOICE><CHOICE>Email failed</CHOICE></CHOICES></Field>'],
    ['StatusDetail', '<Field Type="Note" Name="StatusDetail" StaticName="StatusDetail" DisplayName="Status Detail" RichText="FALSE" NumLines="3" />']
  ];
  for (const [name, xml] of fields) {
    if (await sp(`${list}/fields/getbyinternalnameortitle('${name}')`)) { console.log(`  field ${name} exists`); continue; }
    // AddToDefaultContentType (1) | AddFieldInternalNameHint (8) | AddFieldToDefaultView (16); LineItems stays out of the view.
    const options = name === 'LineItems' ? 9 : 25;
    await sp(`${list}/fields/createfieldasxml`, {
      method: 'POST',
      body: { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: xml, Options: options } }
    });
    console.log(`  added field ${name}`);
  }

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
  const accounts = await sp(`web/sitegroups/getbyname('${ACCOUNTS_GROUP}')`);
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
  const folder = (await sp(`${lib}/RootFolder?$select=ServerRelativeUrl`)).ServerRelativeUrl;
  console.log(`Done. Upload Templates/xlsx/*.xlsx to ${location.origin}${folder}`);
  console.log(`Template library URL for the web part: ${folder}`);
})().catch((e) => console.error('Provisioning failed:', e.message));
