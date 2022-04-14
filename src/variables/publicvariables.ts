/** function to trigger in apps script */
export const FUNCTION_NAME = 'AutoResponder';

/** Label to create for and attach to email messages */
export const LABEL_NAME = 'auto-responder-label';

/** Regex For Bounced Messages */
export const BOUNCED_MESSAGES_REGEX = new RegExp('mailer-daemon@googlemail.com|@bounce|@bounced', 'gi');

/** Name of the spreadsheet to create */
export const SPREADSHEET_NAME = 'Email Autoresponder - Job Tracker';

/** Name of the sheet for the emails found results */
export const AUTOMATED_SHEET_NAME = 'Automated Results List';
export const AUTOMATED_SHEET_HEADERS = [
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
];

/** Name of the sheet for the sent auto response email data */
export const PENDING_EMAILS_TO_SEND_SHEET_NAME = 'Pending Emails To Send';
export const PENDING_EMAILS_TO_SEND_HEADERS = [
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
];

/** Name of the sheet for the sent auto response email data */
export const SENT_SHEET_NAME = 'Sent Automated Responses';
export const SENT_SHEET_NAME_HEADERS = [
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
  'sentThreadId',
  'sentEmailMessageId',
  'sentEmailMessageDate',
  'sentThreadPermaLink',
];

/** Name of the sheet for the emails found results */
export const FOLLOW_UP_EMAILS_SHEET_NAME = 'Follow Up Emails Received List';
export const FOLLOW_UP_EMAILS_HEADERS = [
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
];

/** Name of the sheet for the any bounced emails */
export const BOUNCED_SHEET_NAME = 'Bounced Responses';
export const BOUNCED_SHEET_NAME_HEADERS = [
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
];

/** Name of the sheet for the any bounced emails */
export const ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME = 'Always Autorespond List';
export const ALWAYS_RESPOND_DOMAIN_LIST_HEADERS = ['Email or Domain'];
export const ALWAYS_RESPOND_LIST_INITIAL_DATA = [['inmail-hit-reply@linkedin.com']];

/** Name of the sheet for the any bounced emails */
export const DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME = 'Do Not Track List';
export const DO_NOT_TRACK_DOMAIN_LIST_HEADERS = ['Email or Domain'];
export const DO_NOT_TRACK_DOMAIN_LIST_INITIAL_DATA = [['noreply@linkedin.com']];

/** Name of email domains not to send autoreplies */
export const DO_NOT_EMAIL_AUTO_SHEET_NAME = 'Do Not Autorespond List';
export const DO_NOT_EMAIL_AUTO_SHEET_HEADERS = ['Email / Domain', 'Date', 'Sent Previous Emails Count'];
export const DO_NOT_EMAIL_AUTO_INITIAL_DATA = [
  ['noreply@', new Date(), 0],
  ['no-reply@', new Date(), 0],
  ['mailer-daemon@googlemail.com', new Date(), 0],
  ['@bounce', new Date(), 0],
  ['@bounced', new Date(), 0],
];

export const allSheets = [
  AUTOMATED_SHEET_NAME,
  PENDING_EMAILS_TO_SEND_SHEET_NAME,
  SENT_SHEET_NAME,
  FOLLOW_UP_EMAILS_SHEET_NAME,
  BOUNCED_SHEET_NAME,
  ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME,
  DO_NOT_EMAIL_AUTO_SHEET_NAME,
  DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME,
] as const;
