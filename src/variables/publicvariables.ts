/** function to trigger in apps script */
export const FUNCTION_NAME = 'AutoResponder';

/** Label to create for and attach to email messages */
export const LABEL_NAME = 'hasAutoResponse';

/** Regex For Bounced Messages */
export const BOUNCED_MESSAGES_REGEX = new RegExp('mailer-daemon@googlemail.com|@bounce|@bounced', 'gi');

/** Name of the spreadsheet to create */
export const SPREADSHEET_NAME = 'Job Tracker';

/** Name of the sheet for the emails found results */
export const AUTOMATED_SHEET_NAME = 'Automated Results List';
export const AUTOMATED_SHEET_HEADERS = [
  'Email Id',
  'Date',
  'From',
  'ReplyTo',
  'Body Emails',
  'Body',
  'Salary',
  'Email Permalink',
  'Has Email Response',
];

/** Name of the sheet for the sent auto response email data */
export const SENT_SHEET_NAME = 'Sent Automated Responses';
export const SENT_SHEET_NAME_HEADERS = ['Email Id', 'Date', 'From', 'ReplyTo', 'Body', 'Email Permalink'];

/** Name of the sheet for the any bounced emails */
export const BOUNCED_SHEET_NAME = 'Bounced Responses';
export const BOUNCED_SHEET_NAME_HEADERS = ['Email Id', 'Date', 'From', 'ReplyTo', 'Body', 'Email Permalink'];

/** Name of email domains not to send autoreplies */
export const DO_NOT_EMAIL_AUTO_SHEET_NAME = 'Do Not Autorespond List';
export const DO_NOT_EMAIL_AUTO_SHEET_HEADERS = ['Email / Domain', 'Date', 'Sent Previous Emails Count'];
export const DO_NOT_EMAIL_AUTO_INITIAL_DATA = [
  ['noreply@', new Date()],
  ['no-reply@', new Date()],
  ['mailer-daemon@googlemail.com', new Date()],
  ['@bounce', new Date()],
  ['@bounced', new Date()],
];
