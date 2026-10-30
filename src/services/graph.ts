import { MSGraphClientV3 } from '@microsoft/sp-http';
import { formatEmployeeName } from '../timesheet/naming';
import { MailMessage, Mailer } from './submit';

function toBase64(content: ArrayBuffer): string {
  const bytes = new Uint8Array(content);
  let binary = '';
  // Chunked: String.fromCharCode(...bytes) overflows the stack on large files.
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + 0x8000)));
  }
  return btoa(binary);
}

/** Sends from the signed-in user's own mailbox (Mail.Send delegated). */
export class GraphMailer implements Mailer {
  public constructor(private readonly client: MSGraphClientV3) {}

  public async send(m: MailMessage): Promise<void> {
    await this.client.api('/me/sendMail').post({
      message: {
        subject: m.subject,
        body: { contentType: 'Text', content: m.bodyText },
        toRecipients: [{ emailAddress: { address: m.to } }],
        attachments: [
          {
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: m.fileName,
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            contentBytes: toBase64(m.content)
          }
        ]
      },
      saveToSentItems: true
    });
  }
}

export async function getCurrentEmployee(client: MSGraphClientV3): Promise<string> {
  const me = await client.api('/me').select('givenName,surname,displayName').get();
  return formatEmployeeName(me.givenName, me.surname, me.displayName || '');
}
