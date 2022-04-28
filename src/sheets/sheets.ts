import { ValidRowToWriteInCalendarSheet } from '../calendar/calendar';
import {
  createOrSentTemplateEmail,
  DraftAttributeArray,
  EmailReceivedSheetRowItem,
  getEmailByThreadAndAddToMap,
  getToEmailArray,
  ValidAppliedLinkedInSheetRow,
  ValidFollowUpSheetRowItem,
} from '../email/email';
import {
  alwaysAllowMap,
  doNotSendMailAutoMap,
  emailsToAddToPendingSheetMap,
  emailThreadIdsMap,
  ValidRowToWriteInSentSheet,
} from '../global/maps';
import { getSingleUserPropValue, getUserProps, setUserProps } from '../properties-service/properties-service';
import {
  createTriggerForAutoResponsingToEmails,
  deleteAllExistingProjectTriggers,
  deleteAllTriggersWithMatchingFunctionName,
} from '../trigger/trigger';
import { getAtDomainFromEmailAddress, initialGlobalMap } from '../utils/utils';
import {
  allSheets,
  ALWAYS_RESPOND_DOMAIN_LIST_SHEET_HEADERS,
  ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME,
  ALWAYS_RESPOND_LIST_INITIAL_DATA,
  ARCHIVED_THREADS_SHEET_HEADERS,
  ARCHIVED_THREADS_SHEET_NAME,
  RECEIVED_MESSAGES_ARCHIVE_LABEL_NAME,
  AUTOMATED_RECEIVED_SHEET_HEADERS,
  AUTOMATED_RECEIVED_SHEET_NAME,
  BOUNCED_SHEET_HEADERS,
  BOUNCED_SHEET_NAME,
  DO_NOT_EMAIL_AUTO_INITIAL_DATA,
  DO_NOT_EMAIL_AUTO_SHEET_HEADERS,
  DO_NOT_EMAIL_AUTO_SHEET_NAME,
  DO_NOT_TRACK_DOMAIN_LIST_INITIAL_DATA,
  DO_NOT_TRACK_DOMAIN_LIST_SHEET_HEADERS,
  DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME,
  FOLLOW_UP_EMAILS_SHEET_NAME,
  FOLLOW_UP_EMAILS__SHEET_HEADERS,
  PENDING_EMAILS_TO_SEND_SHEET_HEADERS,
  PENDING_EMAILS_TO_SEND_SHEET_NAME,
  SENT_MESSAGES_LABEL_NAME,
  SENT_SHEET_HEADERS,
  SENT_SHEET_NAME,
  AUTOMATED_RECEIVED_SHEET_PROTECTION_DESCRIPTION,
  PENDING_EMAILS_TO_SEND_SHEET_PROTECTION_DESCRIPTION,
  FOLLOW_UP_EMAILS_SHEET_PROTECTION_DESCRIPTION,
  RECEIVED_MESSAGES_LABEL_NAME,
  ARCHIVED_FOLLOW_UP_SHEET_HEADERS,
  ARCHIVED_FOLLOW_UP_SHEET_NAME,
  FOLLOW_UP_MESSAGES_LABEL_NAME,
  SENT_MESSAGES_ARCHIVE_LABEL_NAME,
  LINKEDIN_APPLIED_JOBS_SHEET_NAME,
  LINKEDIN_APPLIED_JOBS_SHEET_HEADERS,
  LINKEDIN_APPLIED_JOBS_SHEET_PROTECTION_DESCRIPTION,
  CALENDAR_EVENTS_SHEET_HEADERS,
  CALENDAR_EVENTS_SHEET_NAME,
} from '../variables/publicvariables';

export type SheetNames =
  | typeof AUTOMATED_RECEIVED_SHEET_NAME
  | typeof SENT_SHEET_NAME
  | typeof FOLLOW_UP_EMAILS_SHEET_NAME
  | typeof BOUNCED_SHEET_NAME
  | typeof ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME
  | typeof DO_NOT_EMAIL_AUTO_SHEET_NAME
  | typeof DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME
  | typeof PENDING_EMAILS_TO_SEND_SHEET_NAME
  | typeof ARCHIVED_THREADS_SHEET_NAME
  | typeof ARCHIVED_FOLLOW_UP_SHEET_NAME
  | typeof LINKEDIN_APPLIED_JOBS_SHEET_NAME
  | typeof CALENDAR_EVENTS_SHEET_NAME;

export type SheetHeaders =
  | typeof AUTOMATED_RECEIVED_SHEET_HEADERS
  | typeof SENT_SHEET_HEADERS
  | typeof FOLLOW_UP_EMAILS__SHEET_HEADERS
  | typeof BOUNCED_SHEET_HEADERS
  | typeof ALWAYS_RESPOND_DOMAIN_LIST_SHEET_HEADERS
  | typeof DO_NOT_EMAIL_AUTO_SHEET_HEADERS
  | typeof DO_NOT_TRACK_DOMAIN_LIST_SHEET_HEADERS
  | typeof PENDING_EMAILS_TO_SEND_SHEET_HEADERS
  | typeof ARCHIVED_THREADS_SHEET_HEADERS
  | typeof ARCHIVED_FOLLOW_UP_SHEET_HEADERS
  | typeof LINKEDIN_APPLIED_JOBS_SHEET_HEADERS
  | typeof CALENDAR_EVENTS_SHEET_HEADERS;

const tabColors = ['blue', 'green', 'red', 'purple', 'orange', 'yellow', 'black', 'teal', 'gold', 'grey'] as const;

export let activeSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
export let activeSheet: GoogleAppsScript.Spreadsheet.Sheet;

type ReplyToUpdateType = [row: number, emailMessage: GoogleAppsScript.Gmail.GmailMessage[]][];
export const repliesToUpdateArray: ReplyToUpdateType = [];

export function WarningResetSheetsAndSpreadsheet() {
  try {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (activeSpreadsheet.getSheetByName(PENDING_EMAILS_TO_SEND_SHEET_NAME)) {
      sendOrMoveManuallyOrDeleteDraftsInPendingSheet({ type: 'delete' }, { deleteAll: true });
    }

    deleteAllExistingProjectTriggers();

    const tempSheet = activeSpreadsheet.insertSheet();

    PropertiesService.getUserProperties().deleteAllProperties();

    const deletedSheetsMap = activeSpreadsheet.getSheets().reduce((acc: Record<SheetNames, true>, sheet) => {
      if (allSheets.includes(sheet.getName() as SheetNames)) {
        acc[sheet.getName() as SheetNames] = true;
        activeSpreadsheet.deleteSheet(sheet);
      }
      return acc;
    }, {} as Record<SheetNames, true>);

    checkExistsOrCreateSpreadsheet(deletedSheetsMap).then((_) => {
      activeSpreadsheet.deleteSheet(tempSheet);
    });
  } catch (error) {
    console.error(error as any);
  }
}

function filteredRecordOfMissingSheets() {
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = activeSS.getSheets();
  const sheetNames = sheets.map((sheet) => sheet.getName());
  return allSheets.reduce((acc: Partial<Record<SheetNames, true>>, ourSheetName) => {
    if (sheetNames.includes(ourSheetName)) return acc;
    acc[ourSheetName] = true;
    return acc;
  }, {});
}

export async function checkExistsOrCreateSpreadsheet(deletedSheetsMap?: Record<SheetNames, true>): Promise<'done'> {
  return new Promise((resolve, _reject) => {
    let { spreadsheetId } = getSpreadSheetId();
    const missingSheetsMap =
      deletedSheetsMap && Object.keys(deletedSheetsMap).length > 0 ? deletedSheetsMap : filteredRecordOfMissingSheets();
    if (spreadsheetId) {
      try {
        SpreadsheetApp.openById(spreadsheetId);
        // const file = DriveApp.getFileById(spreadsheetId);
        // if (file.isTrashed()) throw Error('File in trash, creating a new sheet');
      } catch (error) {
        PropertiesService.getUserProperties().deleteAllProperties();
        spreadsheetId = undefined;
      }
    }

    if (!spreadsheetId || Object.keys(missingSheetsMap).length !== 0) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

      setUserProps({
        spreadsheetId: spreadsheet.getId(),
      });

      if (missingSheetsMap['Applied LinkedIn Jobs']) {
        createSheet(spreadsheet, LINKEDIN_APPLIED_JOBS_SHEET_NAME, LINKEDIN_APPLIED_JOBS_SHEET_HEADERS, {
          tabColor: 'gold',
        });
      }
      if (missingSheetsMap['Calendar Events']) {
        createSheet(spreadsheet, CALENDAR_EVENTS_SHEET_NAME, CALENDAR_EVENTS_SHEET_HEADERS, {
          tabColor: 'green',
        });
      }
      if (missingSheetsMap['Automated Received Emails']) {
        createSheet(spreadsheet, AUTOMATED_RECEIVED_SHEET_NAME, AUTOMATED_RECEIVED_SHEET_HEADERS, {
          tabColor: 'blue',
        });
      }
      if (missingSheetsMap['Pending Emails To Send']) {
        createSheet(spreadsheet, PENDING_EMAILS_TO_SEND_SHEET_NAME, PENDING_EMAILS_TO_SEND_SHEET_HEADERS, {
          tabColor: 'gold',
        });
      }
      if (missingSheetsMap['Sent Email Responses']) {
        createSheet(spreadsheet, SENT_SHEET_NAME, SENT_SHEET_HEADERS, { tabColor: 'green' });
      }
      if (missingSheetsMap['Follow Up Emails Received List']) {
        createSheet(spreadsheet, FOLLOW_UP_EMAILS_SHEET_NAME, FOLLOW_UP_EMAILS__SHEET_HEADERS, { tabColor: 'black' });
      }
      if (missingSheetsMap['Bounced Responses']) {
        createSheet(spreadsheet, BOUNCED_SHEET_NAME, BOUNCED_SHEET_HEADERS, { tabColor: 'red' });
      }
      if (missingSheetsMap['Always Autorespond List']) {
        createSheet(spreadsheet, ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME, ALWAYS_RESPOND_DOMAIN_LIST_SHEET_HEADERS, {
          tabColor: 'teal',
          initData: ALWAYS_RESPOND_LIST_INITIAL_DATA,
        });
      }
      if (missingSheetsMap['Do Not Autorespond List']) {
        createSheet(spreadsheet, DO_NOT_EMAIL_AUTO_SHEET_NAME, DO_NOT_EMAIL_AUTO_SHEET_HEADERS, {
          tabColor: 'purple',
          initData: DO_NOT_EMAIL_AUTO_INITIAL_DATA,
        });
      }
      if (missingSheetsMap['Do Not Track List']) {
        createSheet(spreadsheet, DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME, DO_NOT_TRACK_DOMAIN_LIST_SHEET_HEADERS, {
          tabColor: 'orange',
          initData: DO_NOT_TRACK_DOMAIN_LIST_INITIAL_DATA,
        });
      }
      if (missingSheetsMap['Archived Email Threads']) {
        createSheet(spreadsheet, ARCHIVED_THREADS_SHEET_NAME, ARCHIVED_THREADS_SHEET_HEADERS, {
          tabColor: 'grey',
        });
      }
      if (missingSheetsMap['Archived Follow Up Threads']) {
        createSheet(spreadsheet, ARCHIVED_FOLLOW_UP_SHEET_NAME, ARCHIVED_FOLLOW_UP_SHEET_HEADERS, {
          tabColor: 'black',
        });
      }
    }
    resolve('done');
  });
}

type Options = { tabColor: typeof tabColors[number]; initData: any[][]; unprotectColumnLetters?: string | string[] };

function createSheet(
  activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: SheetNames,
  headersValues: SheetHeaders,
  options: Partial<Options>
) {
  const { initData = [], tabColor = 'green', unprotectColumnLetters = undefined } = options;
  const sheet = activeSS.insertSheet();
  sheet.setName(sheetName);
  setSheetInProps(sheetName, sheet);
  sheet.setTabColor(tabColor);

  writeHeaders(sheet, headersValues);

  initData.length > 0 && setInitialSheetData(sheet, headersValues, initData);

  const numColsToDelete = sheet.getMaxColumns() - headersValues.length;
  numColsToDelete > 0 && sheet.deleteColumns(headersValues.length + 1, numColsToDelete);

  const startRowIndex = initData.length > 50 ? initData.length : 50;
  sheet.deleteRows(startRowIndex, sheet.getMaxRows() - startRowIndex);

  sheet.autoResizeColumns(1, headersValues.length);
  setSheetProtection(sheet, `${sheetName} Sheet Protection`, unprotectColumnLetters);
}

export function setColumnToCheckBox(sheet: GoogleAppsScript.Spreadsheet.Sheet, columnNumberForCheckbox: number) {
  sheet.getRange(2, columnNumberForCheckbox, sheet.getMaxRows() - 1).insertCheckboxes();
}

function getSpreadSheetId() {
  const { spreadsheetId } = getUserProps(['spreadsheetId']);
  return { spreadsheetId };
}

export function setActiveSpreadSheet(ssApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp) {
  try {
    const { spreadsheetId } = getSpreadSheetId();
    if (!spreadsheetId) throw Error('No Spreadsheet Id');
    const spreadsheet = ssApp.openById(spreadsheetId);
    ssApp.setActiveSpreadsheet(spreadsheet);
    activeSpreadsheet = spreadsheet;
  } catch (error) {
    console.error(error as any);
  }
}

function findSheetById(sheetId: number | string) {
  if (activeSpreadsheet == null) return;
  const sheet = activeSpreadsheet.getSheets().find((sheet) => sheet.getSheetId().toString() === sheetId);
  return sheet;
}

function setSheetInProps(sheetName: string, sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  setUserProps({ [sheetName]: sheet.getSheetId().toString() });
}

function getSheet(sheetName: SheetNames) {
  try {
    if (activeSpreadsheet == null) {
      setActiveSpreadSheet(SpreadsheetApp);
      if (activeSheet == null) {
        throw Error('No active spreadsheet in getsheet method');
      }
    }
    const [sheetNameProp, sheetIdProp] = Object.entries(getUserProps([sheetName]))[0];

    const sheet = activeSpreadsheet.getSheetByName(sheetName);
    if (!sheet || !sheetNameProp) {
      const activeSheet = activeSpreadsheet.insertSheet(sheetName);
      setSheetInProps(sheetName, activeSheet);
      return activeSheet;
    } else {
      const sheet = findSheetById(sheetIdProp);
      return sheet;
    }
  } catch (error) {
    console.error(error as any);
    return;
  }
}

export function getSheetByName(sheetName: SheetNames) {
  try {
    if (activeSpreadsheet == null) {
      setActiveSpreadSheet(SpreadsheetApp);
      if (activeSpreadsheet == null) {
        throw Error('No active spreadsheet in method getSheetByName');
      }
    }
    const [_, sheetValueId] = Object.entries(getUserProps([sheetName]))[0];
    const sheet = findSheetById(sheetValueId);
    if (sheet) return sheet;
    const sheetByName = activeSpreadsheet.getSheetByName(sheetName);
    if (!sheetByName) throw Error(`Could Not Find Sheet with Name ${sheetName}`);
    return sheetByName;
  } catch (error) {
    console.error(error as any);
    return null;
  }
}

export function getAndSetActiveSheetByName(sheetName: SheetNames) {
  try {
    if (activeSpreadsheet == null) throw Error('No Active Spreadsheet');
    const sheet = getSheet(sheetName);
    if (sheet) {
      activeSpreadsheet.setActiveSheet(sheet);
      activeSheet = sheet;
      return sheet;
    }
    throw Error(`Cannot Find The Sheet`);
  } catch (error) {
    if (error instanceof Error) console.error(error.message);
    else console.log(error as string | object);
    return;
  }
}

export function getHeaders(sheet: GoogleAppsScript.Spreadsheet.Sheet, headersValues: string[]) {
  const headers = sheet.getRange(1, 1, 1, headersValues.length);
  return headers.getValues();
}

function writeHeaders(sheet: GoogleAppsScript.Spreadsheet.Sheet, headerValues: SheetHeaders) {
  const headers = sheet.getRange(1, 1, 1, headerValues.length);
  const headerRow = [headerValues];
  headers.setValues(headerRow as unknown as string[][]);
  sheet.setFrozenRows(1);
  headers.setFontWeight('bold');
  sheet.getRange(1, headerValues.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

function unprotectedRangesException(
  columnLetters: string | string[],
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  protectedSheet: GoogleAppsScript.Spreadsheet.Protection
) {
  let unprotectedColumns: GoogleAppsScript.Spreadsheet.Range[] | GoogleAppsScript.Spreadsheet.Range;
  if (typeof columnLetters === 'string') {
    unprotectedColumns = sheet.getRange(`${columnLetters}2:${columnLetters}`);
    protectedSheet.setUnprotectedRanges([unprotectedColumns]);
  }
  if (Array.isArray(columnLetters)) {
    unprotectedColumns = columnLetters.map((columnLetter) => sheet.getRange(`${columnLetter}2:${columnLetter}`));
    protectedSheet.setUnprotectedRanges(unprotectedColumns);
  }
}

function setSheetProtection(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  description: string,
  columnLetters?: string | string[]
) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);

  let protectedSheet: GoogleAppsScript.Spreadsheet.Protection;

  if (protections.length === 0) {
    protectedSheet = sheet.protect().setWarningOnly(true).setDescription(description);
    protections.push(protectedSheet);
  }

  if (!columnLetters) return;

  protections.forEach((protection) => {
    unprotectedRangesException(columnLetters, sheet, protection);
  });
}

function setInitialSheetData(sheet: GoogleAppsScript.Spreadsheet.Sheet, headers: SheetHeaders, initialData: any[][]) {
  try {
    const existingData = sheet.getRange(2, 1, headers.length, headers.length).getValues();
    const hasInitialData = existingData.every((row, index) => row[0] === initialData[index][0]);

    if (!hasInitialData) {
      initialData.forEach((row) => {
        sheet.appendRow(row);
      });
    }
  } catch (error) {
    console.error(error as any);
  }
}

export function getAllDataFromSheet(sheetName: SheetNames) {
  try {
    const sheet = getSheetByName(sheetName);
    if (!sheet) throw Error(`Cannot find sheet: ${sheetName}`);

    return sheet.getDataRange().getValues().slice(1);
  } catch (error) {
    console.error(error as any);
    return null;
  }
}

export function getNumRowsAndColsFromArray(arrayOfRows: any[][]) {
  return {
    numRows: arrayOfRows.length,
    numCols: arrayOfRows[0].length,
  };
}

export function setValuesInRangeAndSortSheet(
  numRows: number,
  numCols: number,
  arrayOfRows: any[][],
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  sortOptions?: { sortByCol: number; asc: boolean },
  range = { startRow: 2, startCol: 1 }
) {
  const dataRange = sheet.getRange(range.startRow, range.startCol, numRows, numCols);
  dataRange.setValues(arrayOfRows);
  sortOptions && sheet.sort(sortOptions.sortByCol, sortOptions.asc);
}

export function addRowsToTopOfSheet(numRows: number, sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  sheet.insertRowsBefore(2, numRows);
}

export function formatRowHeight(sheetName: SheetNames) {
  const sheet = getSheetByName(sheetName);
  if (!sheet) throw Error(`Could Not Find ${sheetName} in function call ${formatRowHeight.name}`);
  const numRows = sheet.getDataRange().getNumRows() - 1;
  //@ts-expect-error
  numRows > 0 && sheet.setRowHeightsForced(2, sheet.getDataRange().getNumRows() - 1, 21);
}

export function addToRepliesArray(index: number, emailMessages: GoogleAppsScript.Gmail.GmailMessage[]) {
  repliesToUpdateArray.push([index + 2, emailMessages]);
}

export function updateRepliesColumn(rowsToUpdate: ReplyToUpdateType) {
  try {
    const automatedSheet = getSheetByName(`${AUTOMATED_RECEIVED_SHEET_NAME}`);
    if (!automatedSheet) throw Error('Cannot find automated sheet to updated the replies column');
    const dataValues = automatedSheet.getDataRange().getValues();
    rowsToUpdate.forEach(([row, emailMessage]) => {
      const numCols = dataValues[row].length;
      automatedSheet.getRange(row, numCols).setValue(getToEmailArray(emailMessage));
    });
  } catch (error) {
    console.error(error as any);
  }
}

type ValidRowToWriteInPendingSheet = [
  send: boolean,
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
  phoneNumber: string,
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
  manuallyMoveDraftToSent: boolean
];
export function writeEmailsToPendingSheet() {
  const pendingEmailsSheet = getSheetByName(PENDING_EMAILS_TO_SEND_SHEET_NAME);
  if (!pendingEmailsSheet) {
    throw Error(`Could Not Find Pending Emails To Send Sheet`);
  }
  const emailObjects = Array.from(emailsToAddToPendingSheetMap.values());
  const validSentRows: ValidRowToWriteInPendingSheet[] = emailObjects.map(
    ({
      send,
      emailThreadId,
      inResponseToEmailMessageId,
      isReplyorNewEmail,
      date,
      emailFrom,
      emailSendTo,
      emailSubject,
      emailBody,
      domain,
      personFrom,
      phoneNumbers,
      salary,
      emailThreadPermaLink,
    }): ValidRowToWriteInPendingSheet => {
      const [
        draftId,
        draftSentMessageId,
        draftMessageDate,
        draftMessageSubject,
        draftMessageFrom,
        draftMessageTo,
        draftMessageBody,
      ] =
        isReplyorNewEmail === 'new'
          ? (createOrSentTemplateEmail({
              type: 'newDraftEmail',
              recipient: emailSendTo,
              personFrom: personFrom === emailSendTo ? '' : personFrom,
              subject: emailSubject,
            }) as DraftAttributeArray)
          : (createOrSentTemplateEmail({
              type: 'replyDraftEmail',
              personFrom: personFrom === emailSendTo ? '' : personFrom,
              gmailMessageId: inResponseToEmailMessageId,
            }) as DraftAttributeArray);
      return [
        send,
        emailThreadId,
        inResponseToEmailMessageId,
        isReplyorNewEmail,
        date,
        emailFrom,
        emailSendTo,
        emailSubject,
        emailBody,
        domain,
        personFrom,
        phoneNumbers,
        salary,
        emailThreadPermaLink,
        false,
        draftId,
        draftSentMessageId,
        draftMessageDate,
        draftMessageSubject,
        draftMessageFrom,
        draftMessageTo,
        draftMessageBody,
        `https://mail.google.com/mail/u/0/#drafts/${draftSentMessageId}`,
        false,
      ];
    }
  );
  if (validSentRows.length === 0) return;
  const columnsObject = findColumnNumbersOrLettersByHeaderNames({
    sheetName: PENDING_EMAILS_TO_SEND_SHEET_NAME,
    headerName: ['Date of Received Email', 'Send', 'Delete / Discard Draft', 'Manually Move Draft To Sent Sheet'],
  });
  const sendCol = columnsObject.Send;
  const deleteDiscardDraftCol = columnsObject['Delete / Discard Draft'];
  const manuallyMoveCol = columnsObject['Manually Move Draft To Sent Sheet'];
  const dateOfReceivedEmail = columnsObject['Date of Received Email'];
  if (!sendCol || !deleteDiscardDraftCol || !manuallyMoveCol || !dateOfReceivedEmail)
    throw Error(
      `Could Not Find Column Headers with ${findColumnNumbersOrLettersByHeaderNames.name} in ${sendOrMoveManuallyOrDeleteDraftsInPendingSheet.name} to set sheet protection`
    );
  const { numCols, numRows } = getNumRowsAndColsFromArray(validSentRows);
  addRowsToTopOfSheet(numRows, pendingEmailsSheet);
  setValuesInRangeAndSortSheet(
    numRows,
    numCols,
    validSentRows,
    pendingEmailsSheet,
    { asc: false, sortByCol: dateOfReceivedEmail.colNumber },
    { startCol: 1, startRow: 2 }
  );
  setCheckedValueForEachRow(
    emailObjects.map(({ send }) => [send]),
    pendingEmailsSheet,
    sendCol.colNumber
  );
  setCheckedValueForEachRow(
    emailObjects.map((_row) => [false]),
    pendingEmailsSheet,
    deleteDiscardDraftCol.colNumber
  );
  setCheckedValueForEachRow(
    emailObjects.map((_row) => [false]),
    pendingEmailsSheet,
    manuallyMoveCol.colNumber
  );
  setSheetProtection(pendingEmailsSheet, PENDING_EMAILS_TO_SEND_SHEET_PROTECTION_DESCRIPTION, [
    sendCol.colLetter,
    deleteDiscardDraftCol.colLetter,
    manuallyMoveCol.colLetter,
  ]);
  formatRowHeight(PENDING_EMAILS_TO_SEND_SHEET_NAME);
}

function setCheckedValueForEachRow(
  arrayRows: [isChecked: boolean][],
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  columnNumber: number
) {
  arrayRows.forEach(([isChecked], index) => {
    const dataRange = sheet.getRange(2 + index, columnNumber);
    dataRange.insertCheckboxes();
    isChecked && dataRange.check();
  });
}

type SendDraftsOptions = { type: 'send' | 'manuallyMove' | 'delete' };

type DeleteDraftOptions = {
  deleteAll?: boolean;
};

export function setAllPendingDraftsSendCheckBox(truthy: boolean = true) {
  const pendingEmailsSheet = getSheetByName(`${PENDING_EMAILS_TO_SEND_SHEET_NAME}`);
  if (!pendingEmailsSheet) throw Error('No Pending Emails Found, Cannot run ' + setAllPendingDraftsSendCheckBox.name);
  const lastRow = pendingEmailsSheet.getLastRow() - 1;
  const trueRows = Array.from({ length: lastRow }, (_row) => [truthy]);
  if (lastRow <= 0) return;
  pendingEmailsSheet.getRange(2, 1, lastRow).setValues(trueRows);
}

export function sendDraftsIfAutoResponseUserOptionIsOn() {
  const isAutoResOn = getSingleUserPropValue('isAutoResOn');
  const onOrOff = isAutoResOn === 'On' ? true : false;
  if (onOrOff) {
    if (createTriggerForAutoResponsingToEmails() === true) {
      setAllPendingDraftsSendCheckBox(true);
      return;
    }
    sendOrMoveManuallyOrDeleteDraftsInPendingSheet({ type: 'send' }, {});
    return;
  }
  if (onOrOff === false) {
    deleteAllTriggersWithMatchingFunctionName(sendDraftsIfAutoResponseUserOptionIsOn.name);
    setAllPendingDraftsSendCheckBox(false);
  }
}

export function manuallyMoveToFollowUpSheet() {
  const followUpSheet = getSheetByName(FOLLOW_UP_EMAILS_SHEET_NAME);
  const automatedReceivedSheet = getSheetByName(AUTOMATED_RECEIVED_SHEET_NAME);
  const automatedReceivedSheetData = getAllDataFromSheet(AUTOMATED_RECEIVED_SHEET_NAME) as EmailReceivedSheetRowItem[];
  const userLabel = getSingleUserPropValue('labelToSearch');

  if (!followUpSheet) throw Error(`Cannot get ${followUpSheet} SHEET for ${manuallyMoveToFollowUpSheet.name}`);
  if (!automatedReceivedSheet)
    throw Error(`Cannot get ${AUTOMATED_RECEIVED_SHEET_NAME} SHEET for ${manuallyMoveToFollowUpSheet.name}`);
  if (!automatedReceivedSheetData)
    throw Error(`Cannot get ${AUTOMATED_RECEIVED_SHEET_NAME} DATA for ${manuallyMoveToFollowUpSheet.name}`);

  const automatedReceivedColumnHeaders = getAllHeaderColNumsAndLetters<typeof AUTOMATED_RECEIVED_SHEET_HEADERS>({
    sheetName: AUTOMATED_RECEIVED_SHEET_NAME,
  });
  const manuallyMoveToFollowUpEmailCol = automatedReceivedColumnHeaders['Manually Move To Follow Up Emails'];
  const emailThreadIdCol = automatedReceivedColumnHeaders['Email Thread Id'];
  const emailMessageIdCol = automatedReceivedColumnHeaders['Email Message Id'];
  const dateOfEmailCol = automatedReceivedColumnHeaders['Date of Email'];
  const fromEmailCol = automatedReceivedColumnHeaders['From Email'];
  const replyToEmailCol = automatedReceivedColumnHeaders['ReplyTo Email'];
  const emailSubjectCol = automatedReceivedColumnHeaders['Email Subject'];
  const bodyEmailsCol = automatedReceivedColumnHeaders['Body Emails'];
  const emailBodyCol = automatedReceivedColumnHeaders['Email Body'];
  const domainCol = automatedReceivedColumnHeaders['Domain'];
  const personFromCol = automatedReceivedColumnHeaders['Person / Company Name'];
  const phoneNumbersCol = automatedReceivedColumnHeaders['US Phone Number'];
  const salaryCol = automatedReceivedColumnHeaders['Salary'];
  const threadPermalinkCol = automatedReceivedColumnHeaders['Thread Permalink'];

  let rowNumberToDelete: number = 2;
  const rowsToWriteAndDelete = automatedReceivedSheetData.reduce(
    (acc, row) => {
      const manuallyMoveToFollowUpEmail = row[manuallyMoveToFollowUpEmailCol.colNumber - 1];

      if (manuallyMoveToFollowUpEmail) {
        const emailThreadId = row[emailThreadIdCol.colNumber - 1] as string;
        const emailMessageId = row[emailMessageIdCol.colNumber - 1] as string;

        const gmailMessage = GmailApp.getMessageById(emailMessageId);
        const gmailThread = gmailMessage.getThread();
        const lastMessageDate = gmailThread.getLastMessageDate();

        const dateOfEamil = row[dateOfEmailCol.colNumber - 1] as GoogleAppsScript.Base.Date;
        const fromEmail = row[fromEmailCol.colNumber - 1] as string;
        const replyToEmail = row[replyToEmailCol.colNumber - 1] as string;
        const emailSubject = row[emailSubjectCol.colNumber - 1] as string;
        const bodyEmails = row[bodyEmailsCol.colNumber - 1] as string;
        const emailBody = row[emailBodyCol.colNumber - 1] as string;
        const domain = row[domainCol.colNumber - 1] as string;
        const personFrom = row[personFromCol.colNumber - 1] as string;
        const phoneNumbers = row[phoneNumbersCol.colNumber - 1] as string;
        const salary = row[salaryCol.colNumber - 1] as string;
        const emailThreadPermalink = row[threadPermalinkCol.colNumber - 1] as string;

        const validFollowUpSheetRowItem: ValidFollowUpSheetRowItem = [
          emailThreadId,
          emailMessageId,
          dateOfEamil,
          fromEmail,
          replyToEmail,
          emailSubject,
          bodyEmails,
          emailBody,
          domain,
          personFrom,
          phoneNumbers,
          salary,
          emailThreadPermalink,
          `https://mail.google.com/mail/u/0/#label/auto-responder-sent-email-label/${emailMessageId}`,
          '',
          '',
          new Date(),
          '',
          lastMessageDate,
          false,
          false,
          false,
          false,
          false,
          undefined,
        ];

        const sentMessageLabel = GmailApp.getUserLabelByName(SENT_MESSAGES_LABEL_NAME);
        const labels = gmailThread.getLabels();
        const foundSentMsgLabel = labels.find((label) => label.getName() === SENT_MESSAGES_LABEL_NAME);
        if (!foundSentMsgLabel) gmailThread.addLabel(sentMessageLabel);
        const receivedMsgLabel = labels.find((label) => label.getName() === RECEIVED_MESSAGES_LABEL_NAME);
        if (receivedMsgLabel) gmailThread.removeLabel(receivedMsgLabel);
        const userLabelToRemove = labels.find((label) => label.getName() === userLabel);
        if (userLabelToRemove) gmailThread.removeLabel(userLabelToRemove);

        acc.rowsToWrite.push(validFollowUpSheetRowItem);
        acc.rowsToDelete.push(rowNumberToDelete);
        rowNumberToDelete--;
      }
      rowNumberToDelete++;
      return acc;
    },
    { rowsToWrite: [], rowsToDelete: [] } as { rowsToWrite: ValidFollowUpSheetRowItem[]; rowsToDelete: number[] }
  );
  rowsToWriteAndDelete.rowsToDelete.forEach((rowNumber) => {
    automatedReceivedSheet.deleteRow(rowNumber);
  });
  writeMessagesToFollowUpEmailsSheet(rowsToWriteAndDelete.rowsToWrite);
}

export function sendOrMoveManuallyOrDeleteDraftsInPendingSheet(
  { type }: SendDraftsOptions,
  { deleteAll = false }: DeleteDraftOptions
) {
  try {
    const pendingSheetEmailData = getAllDataFromSheet(
      PENDING_EMAILS_TO_SEND_SHEET_NAME
    ) as ValidRowToWriteInPendingSheet[];
    const pendingSheet = getSheetByName(PENDING_EMAILS_TO_SEND_SHEET_NAME);
    const sentEmailsSheet = getSheetByName(SENT_SHEET_NAME);

    if (!sentEmailsSheet) throw Error(`Cannot send emails, no sent sheet found`);
    if (!pendingSheetEmailData) throw Error(`Cannot send emails, no pending emails sheet data found`);
    if (!pendingSheet) throw Error(`Cannot send emails, no pending email sheet found`);

    if (type === 'manuallyMove' || type === 'send') {
      initialGlobalMap('doNotSendMailAutoMap');
      initialGlobalMap('alwaysAllowMap');
    }

    let rowNumber: number = 2;

    const rowsForSentSheet = pendingSheetEmailData.reduce(
      (
        acc: ValidRowToWriteInSentSheet[],
        [
          send,
          _emailThreadId,
          _inResponseToEmailMessageId,
          isReplyorNewEmail,
          _date,
          _emailFrom,
          emailSendTo,
          _emailSubject,
          _emailBody,
          _domain,
          _personFrom,
          _phoneNumbers,
          _salary,
          _emailThreadPermaLink,
          deleteDraft,
          draftId,
          _draftSentMessageId,
          _draftMessageDate,
          _draftMessageSubject,
          _draftMessageFrom,
          _draftMessageTo,
          _draftMessageBody,
          _viewDraftInGmail,
          manuallyMoveDraftToSent,
        ]
      ) => {
        let draftData:
          | [
              sentThreadId: string,
              sentDraftMessageId: string,
              sendDraftDate: GoogleAppsScript.Base.Date,
              sentThreadPermaLink: string
            ]
          | undefined;
        try {
          if (type === 'send' && send === true) {
            const { getDate, getEmailMessageId, getId, getPermalink } = sendDraftOrGetMessageFromDraft(
              { type },
              draftId
            );
            draftData = [getId().toString(), getEmailMessageId().toString(), getDate(), getPermalink()];
          } else if (type === 'manuallyMove' && manuallyMoveDraftToSent === true) {
            const { getDate, getEmailMessageId, getId, getPermalink } = sendDraftOrGetMessageFromDraft(
              { type },
              draftId
            );
            draftData = [getId().toString(), getEmailMessageId().toString(), getDate(), getPermalink()];
          } else if (type === 'delete' && (deleteDraft === true || deleteAll === true)) {
            GmailApp.getDraft(draftId).deleteDraft();
          }
        } catch (error) {
          console.error(error as any);
        } finally {
          if (
            (type === 'send' && send === true) ||
            (type === 'manuallyMove' && manuallyMoveDraftToSent === true) ||
            (type === 'delete' && (deleteDraft === true || deleteAll === true))
          ) {
            pendingSheet.deleteRow(rowNumber);
            rowNumber--;
            if (draftData) {
              const rowData = [
                _emailThreadId,
                _inResponseToEmailMessageId,
                isReplyorNewEmail,
                _date,
                _emailFrom,
                emailSendTo,
                _emailSubject,
                _emailBody,
                _domain,
                _personFrom,
                _phoneNumbers,
                _salary,
                _emailThreadPermaLink,
                deleteDraft,
                draftId,
                _draftSentMessageId,
                _draftMessageDate,
                _draftMessageSubject,
                _draftMessageFrom,
                _draftMessageTo,
                _draftMessageBody,
                _viewDraftInGmail,
                manuallyMoveDraftToSent,
              ];
              addSentEmailsToDoNotReplyMap(emailSendTo);
              acc.push([...rowData, ...draftData] as ValidRowToWriteInSentSheet);
            }
          }
        }

        rowNumber++;
        return acc;
      },
      []
    );

    if (rowsForSentSheet.length > 0) {
      writeToSentEmailsSheet(rowsForSentSheet);
      writeDomainsListToDoNotRespondSheet();
    }
    setProtectionForCheckboxesInPendingSheet();
  } catch (error) {
    console.error(error as any);
  }
}

function setProtectionForCheckboxesInPendingSheet() {
  const pendingSheet = getSheetByName(PENDING_EMAILS_TO_SEND_SHEET_NAME);
  if (!pendingSheet) throw Error(`Cannot find pending sheet in ${setProtectionForCheckboxesInPendingSheet.name}`);

  const pendingSheetHeaders = getAllHeaderColNumsAndLetters({
    sheetName: PENDING_EMAILS_TO_SEND_SHEET_NAME,
  });
  const sendCol = pendingSheetHeaders.Send;
  const deleteDiscardDraftCol = pendingSheetHeaders['Delete / Discard Draft'];
  const manuallyMoveCol = pendingSheetHeaders['Manually Move Draft To Sent Sheet'];

  setSheetProtection(pendingSheet, PENDING_EMAILS_TO_SEND_SHEET_PROTECTION_DESCRIPTION, [
    sendCol.colLetter,
    deleteDiscardDraftCol.colLetter,
    manuallyMoveCol.colLetter,
  ]);
}

function addToDoNotSendMailAutoMap(domainOrEmail: string) {
  const count = doNotSendMailAutoMap.get(domainOrEmail);
  if (count == null) {
    doNotSendMailAutoMap.set(domainOrEmail, 0);
  }
  if (typeof count === 'number') {
    doNotSendMailAutoMap.set(domainOrEmail, count + 1);
  }
}

function addEmailAndDomainIfNotAlwaysAllow(emailAddress: string) {
  const domain = getAtDomainFromEmailAddress(emailAddress);

  if (!alwaysAllowMap.has(domain)) {
    addToDoNotSendMailAutoMap(domain);
  }

  if (!alwaysAllowMap.has(emailAddress)) {
    addToDoNotSendMailAutoMap(emailAddress);
  }
}

function addSentEmailsToDoNotReplyMap(sentEmails: string | string[]) {
  if (typeof sentEmails === 'string') {
    addEmailAndDomainIfNotAlwaysAllow(sentEmails);
  }

  if (Array.isArray(sentEmails)) {
    sentEmails.forEach((email) => {
      addEmailAndDomainIfNotAlwaysAllow(email);
    });
  }
}

function sendDraftOrGetMessageFromDraft({ type }: SendDraftsOptions, draftId: string) {
  const {
    getThread,
    getId: getEmailMessageId,
    getDate,
  } = type === 'send' ? GmailApp.getDraft(draftId).send() : GmailApp.getDraft(draftId).getMessage();
  const { getPermalink, getId, addLabel } = getThread();
  if (type === 'send' || type === 'manuallyMove') {
    let followUpLabelForSentEmails = GmailApp.getUserLabelByName(SENT_MESSAGES_LABEL_NAME);
    if (!followUpLabelForSentEmails) {
      followUpLabelForSentEmails = GmailApp.createLabel(SENT_MESSAGES_LABEL_NAME);
    }
    addLabel(followUpLabelForSentEmails);
  }
  return { getDate, getId, getEmailMessageId, getPermalink };
}

export function writeToSentEmailsSheet(rowsForSentSheet: ValidRowToWriteInSentSheet[]) {
  if (rowsForSentSheet.length === 0) return;
  const sentEmailsSheet = getSheetByName(SENT_SHEET_NAME);

  if (!sentEmailsSheet) throw Error(`Cannot find ${SENT_SHEET_NAME} in ${writeToSentEmailsSheet.name}`);
  const { numCols, numRows } = getNumRowsAndColsFromArray(rowsForSentSheet);

  addRowsToTopOfSheet(numRows, sentEmailsSheet);

  const sentSheetColumnHeaders = getAllHeaderColNumsAndLetters<typeof SENT_SHEET_HEADERS>({
    sheetName: SENT_SHEET_NAME,
  });
  const sentDateEmailMessageCol = sentSheetColumnHeaders['Sent Email Message Date'];
  const emailThreadIdInSentSheetCol = sentSheetColumnHeaders['Email Thread Id'];

  const automatedReceivedColumnHeaders = getAllHeaderColNumsAndLetters<typeof AUTOMATED_RECEIVED_SHEET_HEADERS>({
    sheetName: AUTOMATED_RECEIVED_SHEET_NAME,
  });
  const receivedEmailThreadCol = automatedReceivedColumnHeaders['Email Thread Id'];

  setValuesInRangeAndSortSheet(numRows, numCols, rowsForSentSheet, sentEmailsSheet, {
    asc: false,
    sortByCol: sentDateEmailMessageCol.colNumber,
  });

  writeLinkInCellsFromSheetComparison(
    { sheetToWriteToName: SENT_SHEET_NAME, colNumToWriteTo: emailThreadIdInSentSheetCol.colNumber },
    { sheetToLinkFromName: AUTOMATED_RECEIVED_SHEET_NAME, colNumToLinkFrom: receivedEmailThreadCol.colNumber }
  );
  formatRowHeight(SENT_SHEET_NAME);
}

export function writeLinkInCellsFromSheetComparison(
  sheetOne: { sheetToWriteToName: SheetNames; colNumToWriteTo: number },
  sheetTwo: { sheetToLinkFromName: SheetNames; colNumToLinkFrom: number }
) {
  const sheetToWriteTo = getSheetByName(sheetOne.sheetToWriteToName);
  if (!sheetToWriteTo)
    throw Error(`Cannot find ${sheetOne.sheetToWriteToName} in ${writeLinkInCellsFromSheetComparison.name}`);

  const sheetToLinkFrom = getSheetByName(sheetTwo.sheetToLinkFromName);
  if (!sheetToLinkFrom)
    throw Error(`Cannot find ${sheetTwo.sheetToLinkFromName} in ${writeLinkInCellsFromSheetComparison.name}`);

  const { colNumToWriteTo } = sheetOne;
  const { colNumToLinkFrom } = sheetTwo;

  const sheetToWriteRange = sheetToWriteTo.getRange(2, colNumToWriteTo, sheetToWriteTo.getLastRow() - 1);
  const sheetToLinkFromRange = sheetToLinkFrom.getRange(2, colNumToLinkFrom, sheetToLinkFrom.getLastRow() - 1);

  const sheetToLinkFromMap = new Map();

  sheetToLinkFromRange.getValues().forEach(([cellValue], index) => {
    sheetToLinkFromMap.set(cellValue, index + 2);
  });

  sheetToWriteRange.getValues().forEach(([cellValue], index) => {
    const rowNumberFromLinkFromMap = sheetToLinkFromMap.get(cellValue);
    if (rowNumberFromLinkFromMap) {
      const rangeToBeLinked = sheetToLinkFrom.getRange(rowNumberFromLinkFromMap, colNumToLinkFrom).getA1Notation();
      const rangeToAddLink = sheetToWriteTo.getRange(index + 2, colNumToWriteTo);
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(cellValue)
        .setLinkUrl('#gid=' + sheetToLinkFrom.getSheetId() + '&range=' + rangeToBeLinked)
        .build();
      rangeToAddLink.setRichTextValue(richText);
    }
  });
}

export function writeDomainsListToDoNotRespondSheet() {
  try {
    const doNotRespondSheet = getSheetByName(DO_NOT_EMAIL_AUTO_SHEET_NAME);
    if (!doNotRespondSheet) throw Error(`Could find the Do Not Response List to write the array list to`);
    const existingData = getAllDataFromSheet(DO_NOT_EMAIL_AUTO_SHEET_NAME);
    if (!existingData) throw Error('Could not get existing data from Do Not Respond List');

    existingData.forEach(([domainInSheet, _dateInSheet, countInSheet], rowIndex) => {
      const count = doNotSendMailAutoMap.get(domainInSheet);
      if (typeof count !== 'number' && !count) return;
      if (countInSheet !== count) {
        doNotRespondSheet.getRange(rowIndex + 2, 3).setValue(count);
      }
      doNotSendMailAutoMap.delete(domainInSheet);
    });
    const arrayFromMap = Array.from(doNotSendMailAutoMap.entries());
    const createNewDomainsToAppend = arrayFromMap.map(([domain, count]) => [domain, new Date(), count]);
    createNewDomainsToAppend.forEach((row) => doNotRespondSheet.appendRow(row));
  } catch (error) {
    console.error(error as any);
  }
}

export function archiveOrDeleteSelectEmailThreadIds({ type }: { type: 'archive' | 'delete' | 'remove gmail label' }) {
  const automatedReceivedSheet = getSheetByName(AUTOMATED_RECEIVED_SHEET_NAME);
  // const pendingEmailsSheet = type !== 'delete' && getSheetByName(PENDING_EMAILS_TO_SEND_SHEET_NAME);
  // const sentEmailsSheet = type !== 'delete' && getSheetByName(SENT_SHEET_NAME);

  const archivedEmailsSheet = type === 'archive' && getSheetByName(ARCHIVED_THREADS_SHEET_NAME);

  if (!automatedReceivedSheet) throw Error(`Could Not Find ${AUTOMATED_RECEIVED_SHEET_NAME} Sheet`);
  // if (type !== 'delete' && !pendingEmailsSheet) throw Error(`Could Not Find Pending Emails To Send Sheet`);
  // if (type !== 'delete' && !sentEmailsSheet) throw Error(`Could Not Find Sent Automated Responses Sheet`);
  if (type === 'archive' && !archivedEmailsSheet) throw Error(`Could Not Find Archived Emails Sheet`);

  const automatedReceivedSheetData = automatedReceivedSheet.getDataRange().getValues().slice(1);

  const columnsObject = findColumnNumbersOrLettersByHeaderNames({
    sheetName: AUTOMATED_RECEIVED_SHEET_NAME,
    headerName: [
      'Email Thread Id',
      'Warning: Delete Thread Id',
      'Archive Thread Id',
      'Remove Gmail Label',
      'Email Message Id',
    ],
  });
  const emailThreadIdCol = columnsObject['Email Thread Id'];
  const emailMessageIdCol = columnsObject['Email Message Id'];
  const archiveThreadIdCol = columnsObject['Archive Thread Id'];
  const deleteThreadIdCol = columnsObject['Warning: Delete Thread Id'];
  const removeGmailLabelIdCol = columnsObject['Remove Gmail Label'];

  if (!archiveThreadIdCol || !deleteThreadIdCol || !removeGmailLabelIdCol || !emailThreadIdCol || !emailMessageIdCol)
    throw Error(
      `Cannot find col / header to delete, archive, or remove label in function ${
        archiveOrDeleteSelectEmailThreadIds.name
      } with arguments ${archiveOrDeleteSelectEmailThreadIds.arguments.toString()}`
    );

  let rowNumber: number = 2;
  automatedReceivedSheetData.forEach((row) => {
    const emailThreadId = row[emailThreadIdCol.colNumber - 1];
    const emailMessageId = row[emailMessageIdCol.colNumber - 1];
    const archiveCheckBox = row[archiveThreadIdCol.colNumber - 1];
    const deleteCheckBox = row[deleteThreadIdCol.colNumber - 1];
    const removeLabelCheckbox = row[removeGmailLabelIdCol.colNumber - 1];

    if (
      (type === 'archive' && archiveCheckBox === true) ||
      (type === 'delete' && deleteCheckBox === true) ||
      (type === 'remove gmail label' && removeLabelCheckbox === true)
    ) {
      emailThreadIdsMap.set(emailThreadId, { rowNumber, emailMessageId });
      if (type === 'archive') {
        let archiveGmailLabel = GmailApp.getUserLabelByName(RECEIVED_MESSAGES_ARCHIVE_LABEL_NAME);
        if (!archiveGmailLabel) {
          archiveGmailLabel = GmailApp.createLabel(RECEIVED_MESSAGES_ARCHIVE_LABEL_NAME);
        }
        GmailApp.getThreadById(emailThreadId).addLabel(archiveGmailLabel);
        addRowsToTopOfSheet(1, archivedEmailsSheet as GoogleAppsScript.Spreadsheet.Sheet);
        setValuesInRangeAndSortSheet(1, row.length, [row], archivedEmailsSheet as GoogleAppsScript.Spreadsheet.Sheet);
      }
      if (type === 'delete') {
        try {
          GmailApp.getThreadById(emailThreadId).moveToTrash();
        } catch (error) {
          console.error(error as any, 'Could Not Find That Email By Thread Id or Move It To Trash');
        }
      }
      if (type === 'remove gmail label') {
        const userLabel = getSingleUserPropValue('labelToSearch');
        if (!userLabel) throw Error('No User Label To Search Found. Not Set In User Configuration');
        try {
          const gmailLabel = GmailApp.getUserLabelByName(userLabel);
          GmailApp.getThreadById(emailThreadId).removeLabel(gmailLabel);
        } catch (error) {
          console.error(
            error as any,
            'Could Not Find That Email By Thread Id or Could Not Find That Gmail Label or Remove The Label'
          );
        }
      }
      rowNumber--;
    }
    rowNumber++;
  });

  emailThreadIdsMap.forEach(({ rowNumber }, _emailThreadIdToDelete) => {
    automatedReceivedSheet.deleteRow(rowNumber);
  });
}

export function archiveDeleteAddOrRemoveGmailLabelsInFollowUpSheet({
  type,
}: {
  type: 'archive' | 'delete' | 'remove gmail label' | 'add gmail label';
}) {
  const followUpSheet = getSheetByName(FOLLOW_UP_EMAILS_SHEET_NAME);
  // const pendingEmailsSheet = type !== 'delete' && getSheetByName(PENDING_EMAILS_TO_SEND_SHEET_NAME);
  // const sentEmailsSheet = type !== 'delete' && getSheetByName(SENT_SHEET_NAME);

  const archivedFollowUpSheet = type === 'archive' && getSheetByName(ARCHIVED_FOLLOW_UP_SHEET_NAME);

  if (!followUpSheet)
    throw Error(
      `Could Not Find ${FOLLOW_UP_EMAILS_SHEET_NAME} Sheet ${archiveDeleteAddOrRemoveGmailLabelsInFollowUpSheet.name}`
    );
  // if (type !== 'delete' && !pendingEmailsSheet) throw Error(`Could Not Find Pending Emails To Send Sheet`);
  // if (type !== 'delete' && !sentEmailsSheet) throw Error(`Could Not Find Sent Automated Responses Sheet`);
  if (type === 'archive' && !archivedFollowUpSheet) throw Error(`Could Not Find Archived Emails Sheet`);

  const followUpSheetData = followUpSheet.getDataRange().getValues().slice(1);

  const followUpSheetHeaders = getAllHeaderColNumsAndLetters<typeof FOLLOW_UP_EMAILS__SHEET_HEADERS>({
    sheetName: FOLLOW_UP_EMAILS_SHEET_NAME,
  });
  const emailThreadIdCol = followUpSheetHeaders['Email Thread Id'];
  const emailMessageIdCol = followUpSheetHeaders['Email Message Id'];
  const archiveThreadIdCol = followUpSheetHeaders['Archive Thread Id'];
  const deleteThreadIdCol = followUpSheetHeaders['Warning: Delete Thread Id'];
  const removeGmailLabelIdCol = followUpSheetHeaders['Remove Gmail Label'];
  const addGmailLabelIdCol = followUpSheetHeaders['Add Follows Up Gmail Label'];

  let rowNumber: number = 2;
  followUpSheetData.forEach((row) => {
    const emailThreadId = row[emailThreadIdCol.colNumber - 1];
    const emailMessageId = row[emailMessageIdCol.colNumber - 1];
    const archiveCheckBox = row[archiveThreadIdCol.colNumber - 1];
    const deleteCheckBox = row[deleteThreadIdCol.colNumber - 1];
    const removeLabelCheckbox = row[removeGmailLabelIdCol.colNumber - 1];
    const addLabelCheckbox = row[addGmailLabelIdCol.colNumber - 1];

    if (
      (type === 'archive' && archiveCheckBox === true) ||
      (type === 'delete' && deleteCheckBox === true) ||
      (type === 'remove gmail label' && removeLabelCheckbox === true) ||
      (type === 'add gmail label' && addLabelCheckbox === true)
    ) {
      emailThreadIdsMap.set(emailThreadId, { rowNumber, emailMessageId });
      if (type === 'archive') {
        let archiveGmailLabel = GmailApp.getUserLabelByName(SENT_MESSAGES_ARCHIVE_LABEL_NAME);
        if (!archiveGmailLabel) {
          archiveGmailLabel = GmailApp.createLabel(SENT_MESSAGES_ARCHIVE_LABEL_NAME);
        }
        GmailApp.getThreadById(emailThreadId).addLabel(archiveGmailLabel);
        addRowsToTopOfSheet(1, archivedFollowUpSheet as GoogleAppsScript.Spreadsheet.Sheet);
        setValuesInRangeAndSortSheet(1, row.length, [row], archivedFollowUpSheet as GoogleAppsScript.Spreadsheet.Sheet);
      }
      if (type === 'delete') {
        try {
          GmailApp.getThreadById(emailThreadId).moveToTrash();
        } catch (error) {
          console.error(error as any, 'Could Not Find That Email By Thread Id or Move It To Trash');
        }
      }
      if (type === 'remove gmail label') {
        try {
          const gmailLabel = GmailApp.getUserLabelByName(FOLLOW_UP_MESSAGES_LABEL_NAME);
          GmailApp.getThreadById(emailThreadId).removeLabel(gmailLabel);
        } catch (error) {
          console.error(
            error as any,
            'Could Not Find That Email By Thread Id or Could Not Find That Gmail Label or Remove The Label'
          );
        }
      }
      if (type === 'add gmail label') {
        try {
          let followUpLabel = GmailApp.getUserLabelByName(FOLLOW_UP_MESSAGES_LABEL_NAME);
          if (!followUpLabel) {
            followUpLabel = GmailApp.createLabel(FOLLOW_UP_MESSAGES_LABEL_NAME);
          }
          GmailApp.getThreadById(emailThreadId).addLabel(followUpLabel);
        } catch (error) {
          console.error(
            error as any,
            'Could Not Find That Email By Thread Id or Could Not Find That Gmail Label or Remove The Label'
          );
        }
      }
      rowNumber--;
    }
    rowNumber++;
  });

  if (type !== 'add gmail label') {
    emailThreadIdsMap.forEach(({ rowNumber }, _emailThreadIdToDelete) => {
      followUpSheet.deleteRow(rowNumber);
    });
  }
}

type SheetsAndHeaders =
  | { sheetName: typeof AUTOMATED_RECEIVED_SHEET_NAME; headerName: typeof AUTOMATED_RECEIVED_SHEET_HEADERS[number][] }
  | { sheetName: typeof SENT_SHEET_NAME; headerName: typeof SENT_SHEET_HEADERS[number][] }
  | { sheetName: typeof FOLLOW_UP_EMAILS_SHEET_NAME; headerName: typeof FOLLOW_UP_EMAILS__SHEET_HEADERS[number][] }
  | { sheetName: typeof BOUNCED_SHEET_NAME; headerName: typeof BOUNCED_SHEET_HEADERS[number][] }
  | {
      sheetName: typeof ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME;
      headerName: typeof ALWAYS_RESPOND_DOMAIN_LIST_SHEET_HEADERS[number][];
    }
  | { sheetName: typeof DO_NOT_EMAIL_AUTO_SHEET_NAME; headerName: typeof DO_NOT_EMAIL_AUTO_SHEET_HEADERS[number][] }
  | {
      sheetName: typeof DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME;
      headerName: typeof DO_NOT_TRACK_DOMAIN_LIST_SHEET_HEADERS[number][];
    }
  | {
      sheetName: typeof PENDING_EMAILS_TO_SEND_SHEET_NAME;
      headerName: typeof PENDING_EMAILS_TO_SEND_SHEET_HEADERS[number][];
    }
  | { sheetName: typeof ARCHIVED_THREADS_SHEET_NAME; headerName: typeof ARCHIVED_THREADS_SHEET_HEADERS[number][] };

export function findColumnNumbersOrLettersByHeaderNames<T extends SheetsAndHeaders['headerName']>({
  sheetName,
  headerName,
}: {
  sheetName: SheetNames;
  headerName: T;
}): Partial<Record<T[number], { colNumber: number; colLetter: string }>> {
  const sheet = getSheetByName(sheetName);
  if (!sheet) throw Error(`Could Not find ${sheetName} in ${findColumnNumbersOrLettersByHeaderNames.name} function`);
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const [headerValues] = headerRow.getValues() as T[number][][];
  const headerColsMap = headerValues.reduce(
    (acc: Partial<Record<T[number], { colNumber: number; colLetter: string }>>, curVal, index) => {
      if (headerName.includes(curVal as unknown as never)) {
        let colLetter = sheet.getRange(1, index + 1).getA1Notation();
        if (colLetter) {
          [colLetter] = colLetter.split('1');
        }
        acc[curVal] = { colNumber: index + 1, colLetter: colLetter };
      }
      return acc;
    },
    {}
  );
  return headerColsMap;
}

export function getAllHeaderColNumsAndLetters<V extends SheetHeaders>({
  sheetName,
}: {
  sheetName: SheetNames;
}): { [Key in V[number]]: { colNumber: number; colLetter: string } } {
  const sheet = getSheetByName(sheetName);

  if (!sheet) throw Error(`Could Not find ${sheetName} in ${getAllHeaderColNumsAndLetters.name} function`);

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn());

  const [headerValues] = headerRow.getValues();

  const headerColsMap = headerValues.reduce((acc, curVal, index) => {
    let colLetter = sheet.getRange(1, index + 1).getA1Notation();
    if (colLetter) {
      [colLetter] = colLetter.split('1');
    }
    acc[curVal] = { colNumber: index + 1, colLetter: colLetter };
    return acc;
  }, {});
  return headerColsMap;
}

export function writeEmailDataToReceivedAutomationSheet(emailsForList: EmailReceivedSheetRowItem[]) {
  if (emailsForList.length === 0) return;
  const autoResultsListSheet = getSheetByName(`${AUTOMATED_RECEIVED_SHEET_NAME}`);
  if (!autoResultsListSheet) throw Error(`Cannot find ${AUTOMATED_RECEIVED_SHEET_NAME} Sheet`);

  const { 'Sent Thread Id': sentThreadIdCol } = getAllHeaderColNumsAndLetters({ sheetName: SENT_SHEET_NAME });

  const autoReceivedHeaders = getAllHeaderColNumsAndLetters<typeof AUTOMATED_RECEIVED_SHEET_HEADERS>({
    sheetName: AUTOMATED_RECEIVED_SHEET_NAME,
  });

  const dateCol = autoReceivedHeaders['Date of Email'];
  const archiveThreadIdCol = autoReceivedHeaders['Archive Thread Id'];
  const deleteThreadIdCol = autoReceivedHeaders['Warning: Delete Thread Id'];
  const removeGmalLabelCol = autoReceivedHeaders['Remove Gmail Label'];
  const manuallyMoveToFollowUpEmailCol = autoReceivedHeaders['Manually Move To Follow Up Emails'];
  const manuallyCreateEmailDraftCol = autoReceivedHeaders['Manually Create Pending Email'];
  const lastSentThreadIdCol = autoReceivedHeaders['Last Sent Email Thread Id To This Domain'];

  const { numCols, numRows } = getNumRowsAndColsFromArray(emailsForList);
  addRowsToTopOfSheet(numRows, autoResultsListSheet);
  setValuesInRangeAndSortSheet(numRows, numCols, emailsForList, autoResultsListSheet, {
    sortByCol: dateCol.colNumber,
    asc: false,
  });

  const colNumbersArray = [
    manuallyCreateEmailDraftCol.colNumber,
    manuallyMoveToFollowUpEmailCol.colNumber,
    archiveThreadIdCol.colNumber,
    deleteThreadIdCol.colNumber,
    removeGmalLabelCol.colNumber,
  ];
  colNumbersArray.forEach((colNumber) => {
    setCheckedValueForEachRow(
      emailsForList.map((_) => [false]),
      autoResultsListSheet,
      colNumber
    );
  });
  writeLinkInCellsFromSheetComparison(
    { sheetToWriteToName: AUTOMATED_RECEIVED_SHEET_NAME, colNumToWriteTo: lastSentThreadIdCol.colNumber },
    { sheetToLinkFromName: SENT_SHEET_NAME, colNumToLinkFrom: sentThreadIdCol.colNumber }
  );
  setSheetProtection(autoResultsListSheet, AUTOMATED_RECEIVED_SHEET_PROTECTION_DESCRIPTION, [
    archiveThreadIdCol.colLetter,
    deleteThreadIdCol.colLetter,
    removeGmalLabelCol.colLetter,
    manuallyMoveToFollowUpEmailCol.colLetter,
    manuallyCreateEmailDraftCol.colLetter,
  ]);
}

export function writeMessagesToAppliedLinkedinSheet(validAppliedLinkedInSheetRows: ValidAppliedLinkedInSheetRow[]) {
  if (validAppliedLinkedInSheetRows.length === 0) return;
  const appliedLinkedInSheet = getSheetByName(LINKEDIN_APPLIED_JOBS_SHEET_NAME);
  if (!appliedLinkedInSheet)
    throw Error(`Cannot find ${LINKEDIN_APPLIED_JOBS_SHEET_NAME} in ${writeMessagesToFollowUpEmailsSheet.name}`);

  const appliedLinkedInSheetHeaders = getAllHeaderColNumsAndLetters<typeof LINKEDIN_APPLIED_JOBS_SHEET_HEADERS>({
    sheetName: LINKEDIN_APPLIED_JOBS_SHEET_NAME,
  });

  const archiveCheckBox = appliedLinkedInSheetHeaders['Archive Thread Id'];
  const deleteCheckBox = appliedLinkedInSheetHeaders['Warning: Delete Thread Id'];
  const removeLabelCheckBox = appliedLinkedInSheetHeaders['Remove LinkedIn Jobs Gmail Label'];

  const colNumbersArray: number[] = [
    archiveCheckBox.colNumber,
    deleteCheckBox.colNumber,
    removeLabelCheckBox.colNumber,
  ];
  const checkboxesColLettersArray: string[] = [
    archiveCheckBox.colLetter,
    deleteCheckBox.colLetter,
    removeLabelCheckBox.colLetter,
  ];

  const { numCols, numRows } = getNumRowsAndColsFromArray(validAppliedLinkedInSheetRows);
  addRowsToTopOfSheet(numRows, appliedLinkedInSheet);
  setValuesInRangeAndSortSheet(numRows, numCols, validAppliedLinkedInSheetRows, appliedLinkedInSheet);

  colNumbersArray.forEach((colNumber) =>
    setCheckedValueForEachRow(
      validAppliedLinkedInSheetRows.map((_row) => [false]),
      appliedLinkedInSheet,
      colNumber
    )
  );
  setSheetProtection(
    appliedLinkedInSheet,
    LINKEDIN_APPLIED_JOBS_SHEET_PROTECTION_DESCRIPTION,
    checkboxesColLettersArray
  );
}
export function writeEventsToCalendarSheet(validRowInCalendarSheet: ValidRowToWriteInCalendarSheet[]) {
  if (validRowInCalendarSheet.length === 0) return;
  const calendarSheet = getSheetByName(CALENDAR_EVENTS_SHEET_NAME);
  if (!calendarSheet) throw Error(`Cannot find ${CALENDAR_EVENTS_SHEET_NAME} in ${writeEventsToCalendarSheet.name}`);

  const calendarSheetHeaders = getAllHeaderColNumsAndLetters<typeof CALENDAR_EVENTS_SHEET_HEADERS>({
    sheetName: CALENDAR_EVENTS_SHEET_NAME,
  });

  const startDateCol = calendarSheetHeaders['Event Start Time'];

  // const archiveCheckBox = calendarSheetHeaders['Archive Thread Id'];
  // const deleteCheckBox = calendarSheetHeaders['Warning: Delete Thread Id'];
  // const removeLabelCheckBox = calendarSheetHeaders['Remove LinkedIn Jobs Gmail Label'];

  // const colNumbersArray: number[] = [
  //   archiveCheckBox.colNumber,
  //   deleteCheckBox.colNumber,
  //   removeLabelCheckBox.colNumber,
  // ];
  // const checkboxesColLettersArray: string[] = [
  //   archiveCheckBox.colLetter,
  //   deleteCheckBox.colLetter,
  //   removeLabelCheckBox.colLetter,
  // ];

  const { numCols, numRows } = getNumRowsAndColsFromArray(validRowInCalendarSheet);
  addRowsToTopOfSheet(numRows, calendarSheet);
  setValuesInRangeAndSortSheet(numRows, numCols, validRowInCalendarSheet, calendarSheet, {
    asc: true,
    sortByCol: startDateCol.colNumber,
  });

  // colNumbersArray.forEach((colNumber) =>
  //   setCheckedValueForEachRow(
  //     validRowInCalendarSheet.map((_row) => [false]),
  //     calendarSheet,
  //     colNumber
  //   )
  // );
  // setSheetProtection(
  //   calendarSheet,
  //   LINKEDIN_APPLIED_JOBS_SHEET_PROTECTION_DESCRIPTION,
  //   checkboxesColLettersArray
  // );
}

export function writeMessagesToFollowUpEmailsSheet(validFollowUpList: ValidFollowUpSheetRowItem[]) {
  if (validFollowUpList.length === 0) return;

  const followUpSheet = getSheetByName(FOLLOW_UP_EMAILS_SHEET_NAME);
  if (!followUpSheet)
    throw Error(`Cannot find ${FOLLOW_UP_EMAILS_SHEET_NAME} Sheet in ${writeMessagesToFollowUpEmailsSheet.name}`);

  const followUpSheetColumnHeaders = getAllHeaderColNumsAndLetters<typeof FOLLOW_UP_EMAILS__SHEET_HEADERS>({
    sheetName: FOLLOW_UP_EMAILS_SHEET_NAME,
  });
  const sentSheetColumnHeaders = getAllHeaderColNumsAndLetters<typeof SENT_SHEET_HEADERS>({
    sheetName: SENT_SHEET_NAME,
  });
  const automatedReceivedSheetColumnHeaders = getAllHeaderColNumsAndLetters<typeof AUTOMATED_RECEIVED_SHEET_HEADERS>({
    sheetName: AUTOMATED_RECEIVED_SHEET_NAME,
  });

  const dateOfReceivedEmailCol = followUpSheetColumnHeaders['Date of Received Email'];

  // link sent / followup sheet columns for email message id
  const responseToSentEmailMessageIdCol = followUpSheetColumnHeaders['Response To Sent Email Message Id'];
  const sentEmailMessageIdCol = sentSheetColumnHeaders['Sent Email Message Id'];

  // link email thread ids columns for received & follow up sheets
  const receivedEmailThreadIdCol = automatedReceivedSheetColumnHeaders['Email Thread Id'];
  const followupEmailThreadIdCol = followUpSheetColumnHeaders['Email Thread Id'];

  const archiveThreadIdCol = followUpSheetColumnHeaders['Archive Thread Id'];
  const deleteThreadIdCol = followUpSheetColumnHeaders['Warning: Delete Thread Id'];
  const removeGmailLabelCol = followUpSheetColumnHeaders['Remove Gmail Label'];
  const addFollowUpLabelCol = followUpSheetColumnHeaders['Add Follows Up Gmail Label'];
  const manualRepliedToEmailCheckboxCol = followUpSheetColumnHeaders['Manual: Replied To Email'];

  const checkboxesColNumbersArray = [
    archiveThreadIdCol.colNumber,
    deleteThreadIdCol.colNumber,
    removeGmailLabelCol.colNumber,
    addFollowUpLabelCol.colNumber,
    manualRepliedToEmailCheckboxCol.colNumber,
  ];
  const checkboxesColLettersArray = [
    archiveThreadIdCol.colLetter,
    deleteThreadIdCol.colLetter,
    removeGmailLabelCol.colLetter,
    addFollowUpLabelCol.colLetter,
    manualRepliedToEmailCheckboxCol.colLetter,
  ];

  const { numCols, numRows } = getNumRowsAndColsFromArray(validFollowUpList);

  addRowsToTopOfSheet(numRows, followUpSheet);

  setValuesInRangeAndSortSheet(numRows, numCols, validFollowUpList, followUpSheet, {
    sortByCol: dateOfReceivedEmailCol.colNumber,
    asc: false,
  });

  checkboxesColNumbersArray.forEach((colNumber) => {
    setCheckedValueForEachRow(
      validFollowUpList.map((_) => [false]),
      followUpSheet,
      colNumber
    );
  });
  writeLinkInCellsFromSheetComparison(
    { sheetToWriteToName: FOLLOW_UP_EMAILS_SHEET_NAME, colNumToWriteTo: responseToSentEmailMessageIdCol.colNumber },
    { sheetToLinkFromName: SENT_SHEET_NAME, colNumToLinkFrom: sentEmailMessageIdCol.colNumber }
  );
  writeLinkInCellsFromSheetComparison(
    { sheetToWriteToName: FOLLOW_UP_EMAILS_SHEET_NAME, colNumToWriteTo: followupEmailThreadIdCol.colNumber },
    { sheetToLinkFromName: AUTOMATED_RECEIVED_SHEET_NAME, colNumToLinkFrom: receivedEmailThreadIdCol.colNumber }
  );
  setSheetProtection(followUpSheet, FOLLOW_UP_EMAILS_SHEET_PROTECTION_DESCRIPTION, checkboxesColLettersArray);
}

export function manuallyCreateEmailForSelectedRowsInReceivedSheet() {
  initialGlobalMap('pendingEmailsToSendMap');

  const automatedReceivedSheet = getSheetByName(AUTOMATED_RECEIVED_SHEET_NAME);
  const automatedReceivedSheetData = getAllDataFromSheet(AUTOMATED_RECEIVED_SHEET_NAME) as EmailReceivedSheetRowItem[];

  if (!automatedReceivedSheet)
    throw Error(
      `Cannot find ${AUTOMATED_RECEIVED_SHEET_NAME} SHEET in ${manuallyCreateEmailForSelectedRowsInReceivedSheet.name} execution...`
    );
  if (!automatedReceivedSheetData)
    throw Error(
      `Cannot find ${AUTOMATED_RECEIVED_SHEET_NAME} DATA in ${manuallyCreateEmailForSelectedRowsInReceivedSheet.name} execution...`
    );

  const automatedReceivedSheetColumnHeaders = getAllHeaderColNumsAndLetters<typeof AUTOMATED_RECEIVED_SHEET_HEADERS>({
    sheetName: AUTOMATED_RECEIVED_SHEET_NAME,
  });
  const manuallyCreateEmailDraftColNumber =
    automatedReceivedSheetColumnHeaders['Manually Create Pending Email'].colNumber;

  const threadId = automatedReceivedSheetColumnHeaders['Email Thread Id'].colNumber;
  const emailMessageIdCol = automatedReceivedSheetColumnHeaders['Email Message Id'].colNumber;
  const dateCol = automatedReceivedSheetColumnHeaders['Date of Email'].colNumber;
  const emailFromCol = automatedReceivedSheetColumnHeaders['From Email'].colNumber;
  const emailSubjectCol = automatedReceivedSheetColumnHeaders['Email Subject'].colNumber;
  const emailBodyCol = automatedReceivedSheetColumnHeaders['Email Body'].colNumber;
  const domainCol = automatedReceivedSheetColumnHeaders['Domain'].colNumber;
  const personFromCol = automatedReceivedSheetColumnHeaders['Person / Company Name'].colNumber;
  const phoneNumbersCol = automatedReceivedSheetColumnHeaders['US Phone Number'].colNumber;
  const salaryCol = automatedReceivedSheetColumnHeaders['Salary'].colNumber;
  const emailThreadPermaLinkCol = automatedReceivedSheetColumnHeaders['Thread Permalink'].colNumber;
  const emailsInBodyCol = automatedReceivedSheetColumnHeaders['Body Emails'].colNumber;
  const emailReplyToCol = automatedReceivedSheetColumnHeaders['ReplyTo Email'].colNumber;

  const emailBodyFromStringToArray = (emailBodyString: string) => {
    const hasCommaSeperatedValues = emailBodyString.match(/\,/gim);
    if (hasCommaSeperatedValues) {
      return emailBodyString.split(',').map((str) => str.trim());
    }
    return [emailBodyString];
  };

  automatedReceivedSheetData.forEach((row) => {
    const manuallyCreateEmailDraft = row[manuallyCreateEmailDraftColNumber - 1];
    const emailThreadId = row[threadId - 1] as string;
    if (manuallyCreateEmailDraft) {
      const emailsInBody = row[emailsInBodyCol - 1] as string;
      const emailReplyTo = row[emailReplyToCol - 1] as string;

      const date = row[dateCol - 1] as GoogleAppsScript.Base.Date;
      const domain = row[domainCol - 1] as string;
      const emailBody = row[emailBodyCol - 1] as string;
      const emailFrom = row[emailFromCol - 1] as string;
      const emailSubject = row[emailSubjectCol - 1] as string;
      const emailThreadPermaLink = row[emailThreadPermaLinkCol - 1] as string;
      const inResponseToEmailMessageId = row[emailMessageIdCol - 1] as string;
      const personFrom = row[personFromCol - 1] as string;
      const phoneNumbers = row[phoneNumbersCol - 1] as string;
      const salary = row[salaryCol - 1] as string;

      getEmailByThreadAndAddToMap(
        emailThreadId,
        {
          emailThreadId,
          date,
          domain,
          emailBody,
          emailFrom,
          emailSubject,
          emailThreadPermaLink,
          inResponseToEmailMessageId,
          personFrom,
          phoneNumbers,
          salary,
        },
        emailBodyFromStringToArray(emailsInBody),
        emailReplyTo
      );
    }
  });
  writeEmailsToPendingSheet();
}

export function initSpreadsheet() {
  checkExistsOrCreateSpreadsheet();

  setActiveSpreadSheet(SpreadsheetApp);

  getAndSetActiveSheetByName(AUTOMATED_RECEIVED_SHEET_NAME);
}
