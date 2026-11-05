/*
 * Run from the developer console on https://peakemergency.sharepoint.com.
 * Uses the current SharePoint browser session and SharePoint REST endpoints.
 * Idempotent: creates the list/fields and sets permissions; preserves existing data.
 * Users see and edit only their own reports (item-level security); site Members get Contribute and
 * the 'Payroll Owners' group gets Edit (Edit bypasses item-level security).
 * It never deletes lists or items. Rerun after fixing any reported schema conflict.
 */
await (async () => {
  'use strict';

  const site = 'https://peakemergency.sharepoint.com';
  if (location.origin !== site) {
    throw new Error(`Open ${site} before running this script. Current origin: ${location.origin}`);
  }
  const ACCOUNTS_GROUP = 'Payroll Owners';
  const api = `${site}/_api`;
  const jsonType = 'application/json;odata=verbose';
  let digest;
  let digestExpires = 0;

  const unwrap = response => response?.d ?? response;
  const asArray = value => value?.results ?? value?.value ?? [];
  const getDigest = async () => {
    if (digest && Date.now() < digestExpires) return digest;
    const response = await fetch(`${api}/contextinfo`, {
      method: 'POST', credentials: 'same-origin', headers: { Accept: jsonType }
    });
    const body = await response.json();
    if (!response.ok) throw new Error(`contextinfo failed (${response.status}): ${JSON.stringify(body)}`);
    const context = unwrap(body).GetContextWebInformation;
    digest = context.FormDigestValue;
    digestExpires = Date.now() + Math.max(0, Number(context.FormDigestTimeoutSeconds) - 60) * 1000;
    return digest;
  };

  const request = async (url, { method = 'GET', body, extraHeaders = {} } = {}) => {
    const headers = { Accept: jsonType, ...extraHeaders };
    const isContextInfo = url.endsWith('/contextinfo');
    if (method !== 'GET') {
      if (!isContextInfo) headers['X-RequestDigest'] = await getDigest();
      headers['Content-Type'] = jsonType;
    }
    const response = await fetch(url, {
      method, credentials: 'same-origin', headers,
      ...(body === undefined ? {} : { body: JSON.stringify(body) })
    });
    const text = await response.text();
    let result;
    try { result = text ? JSON.parse(text) : undefined; } catch { result = text; }
    if (!response.ok) {
      const message = result?.error?.message?.value ?? result?.error?.message ?? text;
      throw new Error(`${method} ${url} failed (${response.status}): ${message}`);
    }
    return unwrap(result);
  };

  const listsUrl = `${api}/web/lists`;
  const listPath = id => `${listsUrl}(guid'${id}')`;
  async function ensureList(title, description) {
    const all = asArray(await request(`${listsUrl}?$select=Title,BaseTemplate,Id,ListItemEntityTypeFullName,EnableAttachments,EnableVersioning,ReadSecurity,WriteSecurity&$top=5000`));
    const matches = all.filter(item => item.Title === title);
    if (matches.length > 1) throw new Error(`Multiple lists are titled ${title}; resolve them before provisioning.`);
    let list = matches[0];
    if (list && Number(list.BaseTemplate) !== 100) throw new Error(`${title} exists with BaseTemplate ${list.BaseTemplate}; expected a standard custom list (100).`);
    if (!list) {
      list = await request(listsUrl, {
        method: 'POST',
        body: {
          __metadata: { type: 'SP.List' }, Title: title, Description: description,
          BaseTemplate: 100, AllowContentTypes: false, ContentTypesEnabled: false
        }
      });
    }
    list = await request(`${listPath(list.Id)}?$select=Title,BaseTemplate,Id,ListItemEntityTypeFullName,EnableAttachments,EnableVersioning,ReadSecurity,WriteSecurity`);
    return list;
  }

  async function getFields(list) {
    return asArray(await request(`${listPath(list.Id)}/Fields?$select=Id,Title,InternalName,TypeAsString,Indexed,EnforceUniqueValues,Required,RichText,AppendOnly,DisplayFormat&$top=5000`));
  }

  async function mergeField(list, field, values) {
    await request(`${listPath(list.Id)}/Fields(guid'${field.Id}')`, {
      method: 'POST',
      extraHeaders: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
      body: { __metadata: { type: 'SP.Field' }, ...values }
    });
  }

  async function ensureField(list, fields, spec) {
    let field = fields.find(item => item.InternalName === spec.name);
    if (field && field.TypeAsString !== spec.type) {
      throw new Error(`${list.Title}.${spec.name} is ${field.TypeAsString}; expected ${spec.type}. No schema migration was attempted.`);
    }
    if (!field) {
      await request(`${listPath(list.Id)}/Fields/CreateFieldAsXml`, {
        method: 'POST',
        body: { parameters: {
          __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
          SchemaXml: spec.xml, Options: 4
        } }
      });
      fields = await getFields(list);
      field = fields.find(item => item.InternalName === spec.name);
    }
    if (!field) throw new Error(`SharePoint did not return ${list.Title}.${spec.name} after creation.`);

    if (spec.indexed && field.Indexed !== true) {
      await mergeField(list, field, { Indexed: true });
      fields = await getFields(list);
      field = fields.find(item => item.InternalName === spec.name);
    }
    if (spec.required && field.Required !== true) {
      await mergeField(list, field, { Required: true });
      fields = await getFields(list);
      field = fields.find(item => item.InternalName === spec.name);
    }
    if (spec.unique && field.EnforceUniqueValues !== true) {
      // Index first: SharePoint requires unique columns to be indexed.
      if (field.Indexed !== true) throw new Error(`${list.Title}.${spec.name} could not be indexed before unique-value enforcement.`);
      await mergeField(list, field, { EnforceUniqueValues: true });
      fields = await getFields(list);
      field = fields.find(item => item.InternalName === spec.name);
    }
    if (spec.displayFormat !== undefined && Number(field.DisplayFormat) !== spec.displayFormat) {
      await mergeField(list, field, { DisplayFormat: spec.displayFormat });
    }
    return field;
  }

  async function ensureListSettings(list, values) {
    const updates = {};
    for (const [key, value] of Object.entries(values)) if (list[key] !== value) updates[key] = value;
    if (Object.keys(updates).length) {
      await request(listPath(list.Id), {
        method: 'POST', extraHeaders: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
        body: { __metadata: { type: 'SP.List' }, ...updates }
      });
      return request(`${listPath(list.Id)}?$select=Title,BaseTemplate,Id,ListItemEntityTypeFullName,EnableAttachments,EnableVersioning,ReadSecurity,WriteSecurity`);
    }
    return list;
  }

  let reports = await ensureList('ExpenseReports', 'Peak expense reports; each item stores a versioned 13-row report payload.');
  // Item-level security (Read/WriteSecurity 2): users read and edit only their own reports.
  // SharePoint refuses item-level security on a list with unique-value columns, so ClientId
  // is indexed but not unique (ClientIds are random per report; the web part rejects duplicates).
  const oldClientId = (await getFields(reports)).find(item => item.InternalName === 'ClientId');
  if (oldClientId && oldClientId.EnforceUniqueValues === true) await mergeField(reports, oldClientId, { EnforceUniqueValues: false });
  reports = await ensureListSettings(reports, { EnableAttachments: true, EnableVersioning: true, ReadSecurity: 2, WriteSecurity: 2 });
  const reportSpecs = [
    { name: 'ClientId', type: 'Text', required: true, indexed: true,
      xml: "<Field Type='Text' Name='ClientId' StaticName='ClientId' DisplayName='ClientId' MaxLength='255' Required='TRUE' Indexed='TRUE' />" },
    { name: 'InstructorKey', type: 'Text', required: true, indexed: true,
      xml: "<Field Type='Text' Name='InstructorKey' StaticName='InstructorKey' DisplayName='InstructorKey' MaxLength='255' Required='TRUE' Indexed='TRUE' />" },
    { name: 'InstructorName', type: 'Text', required: true,
      xml: "<Field Type='Text' Name='InstructorName' StaticName='InstructorName' DisplayName='InstructorName' MaxLength='255' Required='TRUE' />" },
    { name: 'FromDate', type: 'DateTime', required: true, displayFormat: 0,
      xml: "<Field Type='DateTime' Name='FromDate' StaticName='FromDate' DisplayName='FromDate' Format='DateOnly' Required='TRUE' />" },
    { name: 'ToDate', type: 'DateTime', required: true, displayFormat: 0,
      xml: "<Field Type='DateTime' Name='ToDate' StaticName='ToDate' DisplayName='ToDate' Format='DateOnly' Required='TRUE' />" },
    { name: 'SchemaVersion', type: 'Number', required: true,
      xml: "<Field Type='Number' Name='SchemaVersion' StaticName='SchemaVersion' DisplayName='SchemaVersion' Decimals='0' Required='TRUE' />" },
    { name: 'Payload', type: 'Note', required: true,
      xml: "<Field Type='Note' Name='Payload' StaticName='Payload' DisplayName='Payload' RichText='FALSE' AppendOnly='FALSE' NumLines='6' Required='TRUE' />" },
    // Read and set by the "Expense report email" flow, same as on Timesheet Submissions.
    // No default: a report has no Status until it is queued for email.
    { name: 'Status', type: 'Choice',
      xml: "<Field Type='Choice' Name='Status' StaticName='Status' DisplayName='Status'><CHOICES><CHOICE>Pending</CHOICE><CHOICE>Queued</CHOICE><CHOICE>Sent</CHOICE><CHOICE>Email failed</CHOICE></CHOICES></Field>" },
    { name: 'StatusDetail', type: 'Note',
      xml: "<Field Type='Note' Name='StatusDetail' StaticName='StatusDetail' DisplayName='Status Detail' RichText='FALSE' NumLines='3' />" },
    { name: 'EmailTo', type: 'Text',
      xml: "<Field Type='Text' Name='EmailTo' StaticName='EmailTo' DisplayName='Email To' />" },
    { name: 'EmailSubject', type: 'Text',
      xml: "<Field Type='Text' Name='EmailSubject' StaticName='EmailSubject' DisplayName='Email Subject' MaxLength='255' />" },
    { name: 'EmailBody', type: 'Note',
      xml: "<Field Type='Note' Name='EmailBody' StaticName='EmailBody' DisplayName='Email Body' RichText='FALSE' NumLines='6' />" }
  ];
  for (const spec of reportSpecs) await ensureField(reports, await getFields(reports), spec);
  const reportFields = await getFields(reports);
  const modified = reportFields.find(field => field.InternalName === 'Modified');
  if (!modified) throw new Error('Built-in Modified field was not found on ExpenseReports.');
  if (modified.Indexed !== true) await mergeField(reports, modified, { Indexed: true });

  // Anyone with Manage Lists (the Edit role) bypasses item-level security.
  // Unique permissions: site Members drop to Contribute on this list, the accounts group gets Edit.
  const roleId = async name => (await request(`${api}/web/roledefinitions/getbyname('${name}')`)).Id;
  const edit = await roleId('Edit');
  const contribute = await roleId('Contribute');
  const listUrl = listPath(reports.Id);
  const { HasUniqueRoleAssignments } = await request(`${listUrl}?$select=HasUniqueRoleAssignments`);
  if (!HasUniqueRoleAssignments) {
    await request(`${listUrl}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`, { method: 'POST' });
    const members = await request(`${api}/web/AssociatedMemberGroup`);
    await request(`${listUrl}/roleassignments/removeroleassignment(principalid=${members.Id},roledefid=${edit})`, { method: 'POST' });
    await request(`${listUrl}/roleassignments/addroleassignment(principalid=${members.Id},roledefid=${contribute})`, { method: 'POST' });
    console.info(`List permissions: '${members.Title}' set to Contribute`);
  }
  const [accounts] = asArray(await request(`${api}/web/sitegroups?$filter=Title eq '${ACCOUNTS_GROUP}'&$select=Id`));
  if (!accounts) throw new Error(`SharePoint group '${ACCOUNTS_GROUP}' not found on this site`);
  await request(`${listUrl}/roleassignments/addroleassignment(principalid=${accounts.Id},roledefid=${edit})`, { method: 'POST' });

  const finalReports = await request(`${listUrl}?$select=Title,EnableAttachments,EnableVersioning,ReadSecurity,WriteSecurity`);
  if (finalReports.EnableAttachments !== true || finalReports.EnableVersioning !== true || finalReports.ReadSecurity !== 2 || finalReports.WriteSecurity !== 2) {
    throw new Error('ExpenseReports was created but attachments, versioning or item-level security could not be verified.');
  }
  console.info('SharePoint REST provisioning complete.', { site, reportsList: reports.Title, itemLevelSecurity: true, accountsGroup: ACCOUNTS_GROUP });
})().catch(error => {
  console.error('SharePoint REST provisioning stopped. No lists or items were deleted.', error);
});
