/** Label to create for and attach to email messages */
export const LABEL_NAME = 'auto-responder-label';

/** Label for archived email threads */
export const ARCHIVE_LABEL_NAME = 'archived-auto-responder-label';

/** Label for send / moved emails drafts/messaged for follow up responses GMAIL search */
export const FOLLOW_UP_LABEL_NAME = 'sent-email-auto-responder-label';

/** Regex For Bounced Messages */
export const BOUNCED_MESSAGES_REGEX = new RegExp('mailer-daemon@googlemail.com|@bounce|@bounced', 'gi');

/** Name of the spreadsheet to create */
export const SPREADSHEET_NAME = 'Email Autoresponder - Job Tracker';

/** Name of the sheet for the emails found results */
export const AUTOMATED_RECEIVED_SHEET_NAME = 'Automated Received Emails';
export const AUTOMATED_RECEIVED_SHEET_HEADERS = [
  'Email Thread Id',
  'Email Message Id',
  'Date',
  'From',
  'ReplyTo',
  'Person / Company Name',
  'US Phone Number',
  'Subject',
  'Body Emails',
  'Body',
  'Salary',
  'Thread Permalink',
  'Has Email Response',
  'Archive Thread Id',
  'Warning: Delete Thread Id',
  'Remove Gmail Label',
] as const;

/** Name of the sheet for the sent auto response email data */
export const PENDING_EMAILS_TO_SEND_SHEET_NAME = 'Pending Emails To Send';
export const PENDING_EMAILS_TO_SEND_SHEET_HEADERS = [
  'Send',
  'Email Thread Id',
  'In Response To Email Message Id',
  'Is Reply Or New Email',
  'Date',
  'From',
  'Send Email To',
  'Person / Company Name',
  'Subject',
  'Body',
  'Thread Permalink',
  'deleteDraft',
  'draftId',
  'draftSentMessageId',
  'draftMessageDate',
  'draftMessageSubject',
  'draftMessageFrom',
  'draftMessageTo',
  'draftMessageBody',
  'viewDraftInGmail',
  'manuallyMoveDraftToSent',
] as const;

/** Name of the sheet for the sent auto response email data */
export const SENT_SHEET_NAME = 'Sent Automated Responses';
export const SENT_SHEET_HEADERS = [
  'Email Thread Id',
  'In Response To Email Message Id',
  'Is Reply Or New Email',
  'Date',
  'From',
  'Send Email To',
  'Person / Company Name',
  'Subject',
  'Body',
  'Thread Permalink',
  'deleteDraft',
  'draftId',
  'draftSentMessageId',
  'draftMessageDate',
  'draftMessageSubject',
  'draftMessageFrom',
  'draftMessageTo',
  'draftMessageBody',
  'viewDraftInGmail',
  'manuallyMoveDraftToSent',
  'sentThreadId',
  'sentEmailMessageId',
  'sentEmailMessageDate',
  'sentThreadPermaLink',
] as const;

/** Name of the sheet for the emails found results */
export const FOLLOW_UP_EMAILS_SHEET_NAME = 'Follow Up Emails Received List';
export const FOLLOW_UP_EMAILS__SHEET_HEADERS = [
  'From Domain',
  'Email Thread Id',
  'Email Message Id',
  'Date',
  'From',
  'ReplyTo',
  'Person / Company Name',
  'Subject',
  'Body Emails',
  'Body',
  'Salary',
  'Thread Permalink',
  'Has Email Response',
] as const;

/** Name of the sheet for the any bounced emails */
export const BOUNCED_SHEET_NAME = 'Bounced Responses';
export const BOUNCED_SHEET_HEADERS = [
  'Email Thread Id',
  'Sent Email Message Id',
  'In Response Email Message Id',
  'Date',
  'From',
  'Person / Company Name',
  'ReplyTo',
  'Subject',
  'Body',
  'Thread Permalink',
] as const;

/** Name of the sheet for the any bounced emails */
export const ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME = 'Always Autorespond List';
export const ALWAYS_RESPOND_DOMAIN_LIST_SHEET_HEADERS = ['Email or Domain'] as const;
export const ALWAYS_RESPOND_LIST_INITIAL_DATA = [['inmail-hit-reply@linkedin.com']];

/** Name of the sheet for the any bounced emails */
export const DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME = 'Do Not Track List';
export const DO_NOT_TRACK_DOMAIN_LIST_SHEET_HEADERS = ['Email or Domain'] as const;
export const DO_NOT_TRACK_DOMAIN_LIST_INITIAL_DATA = [['noreply@linkedin.com']];

/** Name of email domains not to send autoreplies */
export const DO_NOT_EMAIL_AUTO_SHEET_NAME = 'Do Not Autorespond List';
export const DO_NOT_EMAIL_AUTO_SHEET_HEADERS = ['Email / Domain', 'Date', 'Sent Previous Emails Count'] as const;
export const DO_NOT_EMAIL_AUTO_INITIAL_DATA = [
  ['noreply@', new Date(), 0],
  ['no-reply@', new Date(), 0],
  ['mailer-daemon@googlemail.com', new Date(), 0],
  ['@bounce', new Date(), 0],
  ['@bounced', new Date(), 0],
];

/** Name of the sheet for the archived found results */
export const ARCHIVED_THREADS_SHEET_NAME = 'Archived Email Threads';
export const ARCHIVED_THREADS_SHEET_HEADERS = [
  'Email Thread Id',
  'Email Message Id',
  'Date',
  'From',
  'ReplyTo',
  'Person / Company Name',
  'Subject',
  'Body Emails',
  'Body',
  'Salary',
  'Thread Permalink',
  'Has Email Response',
  'Archive Thread Id',
  'Warning: Delete Thread Id',
  'Remove Gmail Label',
] as const;

export const allSheets = [
  AUTOMATED_RECEIVED_SHEET_NAME,
  PENDING_EMAILS_TO_SEND_SHEET_NAME,
  SENT_SHEET_NAME,
  FOLLOW_UP_EMAILS_SHEET_NAME,
  BOUNCED_SHEET_NAME,
  ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME,
  DO_NOT_EMAIL_AUTO_SHEET_NAME,
  DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME,
  ARCHIVED_THREADS_SHEET_NAME,
] as const;
