import { SENT_SHEET_HEADERS } from '../variables/publicvariables';

export type EmailDataToSend = {
  send: boolean;
  emailThreadId: string;
  inResponseToEmailMessageId: string;
  isReplyorNewEmail: 'new' | 'reply';
  date: GoogleAppsScript.Base.Date;
  emailFrom: string;
  emailSendTo: string;
  emailSubject: string;
  emailBody: string;
  domain: string;
  personFrom: string;
  phoneNumbers: string;
  salary: string;
  emailThreadPermaLink: string;
};

/** list of emails that should never received an autoresponse */
export const doNotSendMailAutoMap = new Map<string, number>();

/** list of emails or domains that are ignored */
export const doNotTrackMap = new Map<string, boolean>();

/** allow list of domains or emails */
export const alwaysAllowMap = new Map<string, boolean>();

/** list of all email thread ids (threadId) and row number / messageId in automation sheet */
export const emailThreadIdsMap = new Map<string, { rowNumber: number; emailMessageId: string }>();

/** list of existing linkedin by email thread ids (threadId) and row number / messageId in automation sheet */
export const emailThreadsIdAppliedLinkedInMap = new Map<string, { rowNumber: number; emailMessageId: string }>();

/** list of existing calendar events by event ids (eventId) and row number in calendar sheet */
export const calendarEventsMap = new Map<string, { rowNumber: number }>();

/** list of all emails by email address as key, and replyToEmail as boolean with emailSubject and body */
export const emailsToAddToPendingSheetMap = new Map<string, EmailDataToSend>();

/** map of all emails in pending to send sheet, key is "send email to", value is true,  */
export const pendingEmailsToSendMap = new Map<string, true>();

export type ValidRowToWriteInSentSheet = [
  emailThreadId: string,
  inResponseToEmailMessageId: string,
  isReplyorNewEmail: 'new' | 'reply',
  date: GoogleAppsScript.Base.Date,
  emailFrom: string,
  emailSendTo: string,
  emailSubject: string,
  emailBody: string,
  domain: string,
  personFrom: string,
  phoneNumbers: string,
  salary: string,
  emailThreadPermaLink: string,
  deleteDraft: boolean,
  draftId: string,
  draftSentMessageId: string,
  draftMessageDate: GoogleAppsScript.Base.Date,
  draftMessageSubject: string,
  draftMessageFrom: string,
  draftMessageTo: string,
  draftMessageBody: string,
  viewDraftInGmail: string,
  manuallyMoveDraftToSent: boolean,
  sentThreadId: string,
  sentEmailMessageId: string,
  sentEmailMessageDate: GoogleAppsScript.Base.Date,
  sentThreadPermaLink: string
];
/** map of all sent emails in the sent email responses sheet with the sent message ID as the key  */
export const sentEmailsBySentMessageIdMap = new Map<
  string,
  {
    rowObject: Record<typeof SENT_SHEET_HEADERS[number], ValidRowToWriteInSentSheet[number]>;
    rowArray: ValidRowToWriteInSentSheet;
  }
>();

/** map of all sent emails in the sent email responses sheet with the domain as the key  */
export const sentEmailsByDomainMap = new Map<
  string,
  {
    rowObject: Record<typeof SENT_SHEET_HEADERS[number], ValidRowToWriteInSentSheet[number]>;
    rowArray: ValidRowToWriteInSentSheet;
  }
>();

/** map of all rows in the follow up  sheet, key is "received message id", value is true,  */
export const followUpSheetMessageIdMap = new Map<string, true>();
