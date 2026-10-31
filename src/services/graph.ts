import { MSGraphClientV3 } from '@microsoft/sp-http';
import { formatEmployeeName } from '../timesheet/naming';

export async function getCurrentEmployee(client: MSGraphClientV3): Promise<string> {
  const me = await client.api('/me').select('givenName,surname,displayName').get();
  return formatEmployeeName(me.givenName, me.surname, me.displayName || '');
}
