/** Label to create for and attach to email messages */
export const RECEIVED_MESSAGES_LABEL_NAME = 'auto-responder-received-label';

/** Label for archived email threads */
export const RECEIVED_MESSAGES_ARCHIVE_LABEL_NAME = 'auto-responder-archived-label';

/** Label for send / moved emails drafts/messaged for follow up responses GMAIL search */
export const SENT_MESSAGES_LABEL_NAME = 'auto-responder-sent-email-label';

/** Label for archiving and excluding sent emails in the GMAIL search */
export const SENT_MESSAGES_ARCHIVE_LABEL_NAME = 'auto-responder-sent-email-archived-label';

/** Label to apply a 'auto-responder-follow-up-label' in GMAIL manually for organization purposes */
export const FOLLOW_UP_MESSAGES_LABEL_NAME = 'auto-responder-follow-up-label';

/** Label for send / moved emails drafts/messaged for follow up responses GMAIL search */
export const LINKEDIN_APPLIED_LABEL_NAME = 'auto-responder-linkedin-applied-label';

/** Label for send / moved emails drafts/messaged for follow up responses GMAIL search */
export const LINKEDIN_APPLIED_ARCHIVE_LABEL_NAME = 'auto-responder-linkedin-applied-archived-label';

/** Regex For Bounced Messages */
export const BOUNCED_MESSAGES_REGEX = new RegExp('mailer-daemon@googlemail.com|@bounce|@bounced', 'gi');

/** Name of the spreadsheet to create */
export const SPREADSHEET_NAME = 'Email Autoresponder - Job Tracker';

/** Name of the sheet for the emails found results */
export const AUTOMATED_RECEIVED_SHEET_NAME = 'Automated Received Emails';
export const AUTOMATED_RECEIVED_SHEET_HEADERS = [
  'Email Thread Id',
  'Email Message Id',
  'Date of Email',
  'From Email',
  'ReplyTo Email',
  'Email Subject',
  'Body Emails',
  'Email Body',
  'Domain',
  'Person / Company Name',
  'US Phone Number',
  'Salary',
  'Thread Permalink',
  'Has Email Response',
  'Last Sent Email To This Domain',
  'Last Sent Email Thread Id To This Domain',
  'Last Sent Email Subject To This Domain',
  'Last Sent Person / Company From This Domain',
  'Manually Move To Follow Up Emails',
  'Manually Create Pending Email',
  'Archive Thread Id',
  'Warning: Delete Thread Id',
  'Remove Gmail Label',
] as const;
export const AUTOMATED_RECEIVED_SHEET_PROTECTION_DESCRIPTION = `${AUTOMATED_RECEIVED_SHEET_NAME} Sheet Protection`;

/** Name of the sheet for the sent auto response email data */
export const PENDING_EMAILS_TO_SEND_SHEET_NAME = 'Pending Emails To Send';
export const PENDING_EMAILS_TO_SEND_SHEET_HEADERS = [
  'Send',
  'Email Thread Id',
  'In Response To Email Message Id',
  'Is Reply Or New Email',
  'Date of Received Email',
  'From Email',
  'Send Email To',
  'Received Email Subject',
  'Received Email Body',
  'Domain',
  'Person / Company Name',
  'US Phone Number',
  'Salary',
  'Thread Permalink',
  'Delete / Discard Draft',
  'Draft Id',
  'Draft Sent Message Id',
  'Draft Creation Date',
  'Draft Subject',
  'Draft From Email',
  'Draft Message To Email',
  'Draft Email Body',
  'View Draft In Gmail',
  'Manually Move Draft To Sent Sheet',
] as const;
export const PENDING_EMAILS_TO_SEND_SHEET_PROTECTION_DESCRIPTION = `${PENDING_EMAILS_TO_SEND_SHEET_NAME} Sheet Protection`;

/** Name of the sheet for the sent auto response email data */
export const SENT_SHEET_NAME = 'Sent Email Responses';
export const SENT_SHEET_HEADERS = [
  'Email Thread Id',
  'In Response To Email Message Id',
  'Is Reply Or New Email',
  'Date of Received Email',
  'From Email',
  'Send Email To',
  'Received Email Subject',
  'Received Email Body',
  'Domain',
  'Person / Company Name',
  'US Phone Number',
  'Salary',
  'Thread Permalink',
  'Delete / Discard Draft',
  'Draft Id',
  'Draft Sent Message Id',
  'Draft Creation Date',
  'Draft Subject',
  'Draft From Email',
  'Draft Message To Email',
  'Draft Email Body',
  'View Draft In Gmail',
  'Manually Move Draft To Sent Sheet',
  'Sent Thread Id',
  'Sent Email Message Id',
  'Sent Email Message Date',
  'Sent Thread PermaLink',
] as const;
export const SENT_SHEET_PROTECTION_DESCRIPTION = `${SENT_SHEET_NAME} Sheet Protection`;

/** Name of the sheet for the emails found results */
export const FOLLOW_UP_EMAILS_SHEET_NAME = 'Follow Up Emails Received List';
export const FOLLOW_UP_EMAILS__SHEET_HEADERS = [
  'Email Thread Id',
  'Email Message Id',
  'Date of Received Email',
  'From Email',
  'Send Email To',
  'Received Email Subject',
  'Body Emails',
  'Received Email Body',
  'Domain',
  'Person / Company Name',
  'US Phone Number',
  'Salary',
  'Thread Permalink',
  'View In Gmail',
  'Response To Sent Thread Id',
  'Response To Sent Email Message Id',
  'Response To Sent Email Message Date',
  'Response To Sent Thread PermaLink',
  'Last Email Message Date In Thread',
  'Archive Thread Id',
  'Warning: Delete Thread Id',
  'Remove Gmail Label',
  'Add Follows Up Gmail Label',
  'Manual: Replied To Email',
  'Manual: Replied To Date',
] as const;
export const FOLLOW_UP_EMAILS_SHEET_PROTECTION_DESCRIPTION = `${FOLLOW_UP_EMAILS_SHEET_NAME} Sheet Protection`;

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
export const ALWAYS_RESPOND_LIST_INITIAL_DATA = [['inmail-hit-reply@linkedin.com', '@gmail.com']];

/** Name of the sheet for the any bounced emails */
export const DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME = 'Do Not Track List';
export const DO_NOT_TRACK_DOMAIN_LIST_SHEET_HEADERS = ['Email or Domain'] as const;
export const DO_NOT_TRACK_DOMAIN_LIST_INITIAL_DATA = [
  ['noreply@linkedin.com'],
  ['jobs-listings@linkedin.com'],
  ['jobs-noreply@linkedin.com'],
];

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
  'Date of Email',
  'From Email',
  'ReplyTo Email',
  'Email Subject',
  'Body Emails',
  'Email Body',
  'Domain',
  'Person / Company Name',
  'US Phone Number',
  'Salary',
  'Thread Permalink',
  'Has Email Response',
  'Archive Thread Id',
  'Warning: Delete Thread Id',
  'Remove Gmail Label',
] as const;

/** Name of the sheet for the archived found results */
export const ARCHIVED_FOLLOW_UP_SHEET_NAME = 'Archived Follow Up Threads';
export const ARCHIVED_FOLLOW_UP_SHEET_HEADERS = FOLLOW_UP_EMAILS__SHEET_HEADERS;

export const LINKEDIN_APPLIED_JOBS_SHEET_NAME = 'Applied LinkedIn Jobs';
export const LINKEDIN_APPLIED_JOBS_SHEET_HEADERS = [
  'Email Thread Id',
  'Email Message Id',
  'Date of Email',
  'Email Subject',
  'Email Body',
  'Thread Permalink',
  'Company Name',
  'Job Position',
  'Point of Contact',
  'View Job On LinkedIn',
  'Link To Other Email Threads',
  'Archive Thread Id',
  'Warning: Delete Thread Id',
  'Remove LinkedIn Jobs Gmail Label',
] as const;
export const LINKEDIN_APPLIED_JOBS_SHEET_PROTECTION_DESCRIPTION = `${LINKEDIN_APPLIED_JOBS_SHEET_NAME} Sheet Protection`;

export const CALENDAR_EVENTS_SHEET_NAME = 'Calendar Events';
export const CALENDAR_EVENTS_SHEET_HEADERS = [
  'Calendar Event Id',
  'Event Last Updated Time',
  'Event Created Time',
  'Event Start Time',
  'Event End Time',
  'Event Title',
  'Event Description',
  'Guest Company Name',
  'Guest Domain',
  'Guest Name',
  'Guest Email',
  'Guest Phone Number',
  'Guest Status',
  'Number Of Guests',
  'Event URL',
] as const;
export const CALENDAR_EVENTS_SHEET_PROTECTION_DESCRIPTION = `${CALENDAR_EVENTS_SHEET_NAME} Sheet Protection`;

export const LINKED_JOB_SEARCH_EMAILS = {
  applied: 'jobs-listings@linkedin.com',
  viewed: 'jobs-noreply@linkedin.com',
} as const;

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
  ARCHIVED_FOLLOW_UP_SHEET_NAME,
  LINKEDIN_APPLIED_JOBS_SHEET_NAME,
  CALENDAR_EVENTS_SHEET_NAME,
] as const;
