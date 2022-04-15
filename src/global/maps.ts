export type EmailDataToSend = {
  sendReplyEmail: boolean;
  send: boolean;
  emailThreadId: string;
  isReplyorNewEmail: 'new' | 'reply';
  emailSubject: string;
  personFrom: string;
  inResponseToEmailMessageId: string;
  date: GoogleAppsScript.Base.Date;
  emailFrom: string;
  emailSendTo: string;
  emailBody: string;
  emailThreadPermaLink: string;
};

/** list of emails that should never received an autoresponse */
export const doNotSendMailAutoMap = new Map<string, number>();

/** list of emails or domains that are ignored */
export const doNotTrackMap = new Map<string, boolean>();

/** allow list of domains or emails */
export const alwaysAllowMap = new Map<string, boolean>();

/** list of all emailId (threadId) and row number in automation sheet */
export const emailmessagesIdMap = new Map<string, number>();

/** list of all emails by email address as key, and replyToEmail as boolean with emailSubject and body */
export const emailsToAddToPendingSheet = new Map<string, EmailDataToSend>();

/** map of all emails in pending to send sheet, key is "send email to", value is true,  */
export const pendingEmailsToSendMap = new Map<string, true>();
