const EMAIL = /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/;

/** undefined means "no copy". Returns '' when there is nothing wrong. */
export const copyProblem = (copy: string | undefined): string =>
  copy === undefined || EMAIL.test(copy.trim()) ? '' : 'Enter a valid email address to send a copy to.';

/** The flow's Outlook "Send an email (V2)" takes several To addresses separated by ';'. */
export const recipients = (to: string, copy?: string): string => (copy && copy.trim() ? `${to};${copy.trim()}` : to);
