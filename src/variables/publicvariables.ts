/** function to trigger in apps script */
export const FUNCTION_NAME = 'AutoResponder';

/** Label to create for and attach to email messages */
export const LABEL_NAME = 'hasAutoResponse';

/** Name of the spreadsheet to create */
export const SPREADSHEET_NAME = 'Job Tracker';

/** Name of the sheet for the results */
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
/** Name of the sheet for the results */

/** Name of the sheet for the results */
export const SENT_SHEET_NAME = 'Sent Automated Responses';
export const SENT_SHEET_NAME_HEADERS = ['Email Id', 'Date', 'From', 'ReplyTo', 'Body', 'Email Permalink'];
