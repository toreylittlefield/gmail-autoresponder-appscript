import { createOrSentTemplateEmail, DraftAttributeArray, EmailListItem, getToEmailArray } from '../email/email';
import { alwaysAllowMap, doNotSendMailAutoMap, emailsToAddToPendingSheet, emailThreadIdsMap } from '../global/maps';
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
  ARCHIVE_LABEL_NAME,
  AUTOMATED_RECEIVED_SHEET_HEADERS,
  AUTOMATED_RECEIVED_SHEET_NAME,
  BOUNCED_SHEET_NAME,
  BOUNCED_SHEET_HEADERS,
  DO_NOT_EMAIL_AUTO_INITIAL_DATA,
  DO_NOT_EMAIL_AUTO_SHEET_HEADERS,
  DO_NOT_EMAIL_AUTO_SHEET_NAME,
  DO_NOT_TRACK_DOMAIN_LIST_SHEET_HEADERS,
  DO_NOT_TRACK_DOMAIN_LIST_INITIAL_DATA,
  DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME,
  FOLLOW_UP_EMAILS__SHEET_HEADERS,
  FOLLOW_UP_EMAILS_SHEET_NAME,
  SENT_MESSAGES_LABEL_NAME,
  PENDING_EMAILS_TO_SEND_SHEET_HEADERS,
  PENDING_EMAILS_TO_SEND_SHEET_NAME,
  SENT_SHEET_NAME,
  SENT_SHEET_HEADERS,
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
  | typeof ARCHIVED_THREADS_SHEET_NAME;

export type SheetHeaders =
  | typeof AUTOMATED_RECEIVED_SHEET_HEADERS
  | typeof SENT_SHEET_HEADERS
  | typeof FOLLOW_UP_EMAILS__SHEET_HEADERS
  | typeof BOUNCED_SHEET_HEADERS
  | typeof ALWAYS_RESPOND_DOMAIN_LIST_SHEET_HEADERS
  | typeof DO_NOT_EMAIL_AUTO_SHEET_HEADERS
  | typeof DO_NOT_TRACK_DOMAIN_LIST_SHEET_HEADERS
  | typeof PENDING_EMAILS_TO_SEND_SHEET_HEADERS
  | typeof ARCHIVED_THREADS_SHEET_HEADERS;

const tabColors = ['blue', 'green', 'red', 'purple', 'orange', 'yellow', 'black', 'teal', 'gold', 'grey'] as const;

export let activeSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
export let activeSheet: GoogleAppsScript.Spreadsheet.Sheet;

type ReplyToUpdateType = [row: number, emailMessage: GoogleAppsScript.Gmail.GmailMessage[]][];
export const repliesToUpdateArray: ReplyToUpdateType = [];

export function WarningResetSheetsAndSpreadsheet() {
  try {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    activeSpreadsheet.getSheetByName(PENDING_EMAILS_TO_SEND_SHEET_NAME) &&
      sendOrMoveManuallyOrDeleteDraftsInPendingSheet({ type: 'delete' }, { deleteAll: true });
    deleteAllExistingProjectTriggers();
    const sheetsToDelete: GoogleAppsScript.Spreadsheet.Sheet[] = [];
    allSheets.forEach((sheetName) => {
      const sheet = activeSpreadsheet.getSheetByName(sheetName);
      if (sheet) sheetsToDelete.push(sheet);
    });

    // if (sheetsToDelete.length === 0)
    //   throw Error('Could Not Reset This Spreadsheet Because We Could Not Find The Sheets To Delete');
    const tempSheet = activeSpreadsheet.insertSheet();
    PropertiesService.getUserProperties().deleteAllProperties();
    sheetsToDelete.forEach((sheet) => activeSpreadsheet.deleteSheet(sheet));
    initSpreadsheet();
    activeSpreadsheet.deleteSheet(tempSheet);
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

export async function checkExistsOrCreateSpreadsheet(): Promise<'done'> {
  return new Promise((resolve, _reject) => {
    let { spreadsheetId } = getSpreadSheetId();
    const missingSheetsMap = filteredRecordOfMissingSheets();

    if (spreadsheetId) {
      try {
        SpreadsheetApp.openById(spreadsheetId);
        const file = DriveApp.getFileById(spreadsheetId);
        if (file.isTrashed()) throw Error('File in trash, creating a new sheet');
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
  setSheetProtection(sheet, `${sheetName} Protected Range`, unprotectColumnLetters);
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
  //@ts-expect-error
  sheet && sheet.setRowHeightsForced(2, sheet.getDataRange().getNumRows() - 1, 21);
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
  const emailObjects = Array.from(emailsToAddToPendingSheet.values());
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
              subject: emailSubject,
            }) as DraftAttributeArray)
          : (createOrSentTemplateEmail({
              type: 'replyDraftEmail',
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
  setSheetProtection(pendingEmailsSheet, 'Pending Emails To Send Protected Range', [
    sendCol.colLetter,
    deleteDiscardDraftCol.colLetter,
    manuallyMoveCol.colLetter,
  ]);
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

type ValidRowToWriteInSentSheet = [
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
      writeSentDraftsToSentEmailsSheet(sentEmailsSheet, rowsForSentSheet);
      writeDomainsListToDoNotRespondSheet();
      writeLinkInCellsFromSheetComparison(
        { sheetToWriteToName: SENT_SHEET_NAME, colNumToWriteTo: 1 },
        { sheetToLinkFromName: `${AUTOMATED_RECEIVED_SHEET_NAME}`, colNumToLinkFrom: 1 }
      );
    }
    const columnsObject = findColumnNumbersOrLettersByHeaderNames({
      sheetName: PENDING_EMAILS_TO_SEND_SHEET_NAME,
      headerName: ['Send', 'Delete / Discard Draft', 'Manually Move Draft To Sent Sheet'],
    });
    const sendCol = columnsObject.Send;
    const deleteDiscardDraftCol = columnsObject['Delete / Discard Draft'];
    const manuallyMoveCol = columnsObject['Manually Move Draft To Sent Sheet'];
    if (!sendCol || !deleteDiscardDraftCol || !manuallyMoveCol)
      throw Error(
        `Could Not Find Column Headers with ${findColumnNumbersOrLettersByHeaderNames.name} in ${sendOrMoveManuallyOrDeleteDraftsInPendingSheet.name} to set sheet protection`
      );

    setSheetProtection(pendingSheet, 'Pending Emails To Send Protected Range', [
      sendCol.colLetter,
      deleteDiscardDraftCol.colLetter,
      manuallyMoveCol.colLetter,
    ]);
  } catch (error) {
    console.error(error as any);
  }
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

function writeSentDraftsToSentEmailsSheet(
  sentEmailsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowsForSentSheet: ValidRowToWriteInSentSheet[]
) {
  if (sentEmailsSheet.getName() !== SENT_SHEET_NAME) throw Error('Sheet Must Be The Sent Automated Responses');
  const { numCols, numRows } = getNumRowsAndColsFromArray(rowsForSentSheet);
  addRowsToTopOfSheet(numRows, sentEmailsSheet);
  const columnObject = findColumnNumbersOrLettersByHeaderNames({
    sheetName: SENT_SHEET_NAME,
    headerName: ['Sent Email Message Date'],
  });
  const sentDateEmailMessageCol = columnObject['Sent Email Message Date'];
  if (!sentDateEmailMessageCol)
    throw Error(`Cannot Find Column ${sentDateEmailMessageCol}in ${writeSentDraftsToSentEmailsSheet.name}`);
  setValuesInRangeAndSortSheet(numRows, numCols, rowsForSentSheet, sentEmailsSheet, {
    asc: false,
    sortByCol: sentDateEmailMessageCol.colNumber,
  });
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
  const pendingEmailsSheet = type !== 'delete' && getSheetByName(PENDING_EMAILS_TO_SEND_SHEET_NAME);
  const sentEmailsSheet = type !== 'delete' && getSheetByName(SENT_SHEET_NAME);

  const archivedEmailsSheet = type === 'archive' && getSheetByName(ARCHIVED_THREADS_SHEET_NAME);

  if (!automatedReceivedSheet) throw Error(`Could Not Find ${AUTOMATED_RECEIVED_SHEET_NAME} Sheet`);
  if (type !== 'delete' && !pendingEmailsSheet) throw Error(`Could Not Find Pending Emails To Send Sheet`);
  if (type !== 'delete' && !sentEmailsSheet) throw Error(`Could Not Find Sent Automated Responses Sheet`);
  if (type === 'archive' && !archivedEmailsSheet) throw Error(`Could Not Find Archived Emails Sheet`);

  const automatedReceivedSheetData = automatedReceivedSheet.getDataRange().getValues().slice(1);

  const columnsObject = findColumnNumbersOrLettersByHeaderNames({
    sheetName: AUTOMATED_RECEIVED_SHEET_NAME,
    headerName: ['Email Thread Id', 'Warning: Delete Thread Id', 'Archive Thread Id', 'Remove Gmail Label'],
  });
  const emailThreadIdCol = columnsObject['Email Thread Id'];
  const archiveThreadIdCol = columnsObject['Archive Thread Id'];
  const deleteThreadIdCol = columnsObject['Warning: Delete Thread Id'];
  const removeGmailLabelIdCol = columnsObject['Remove Gmail Label'];

  if (!archiveThreadIdCol || !deleteThreadIdCol || !removeGmailLabelIdCol || !emailThreadIdCol)
    throw Error(
      `Cannot find col / header to delete, archive, or remove label in function ${
        archiveOrDeleteSelectEmailThreadIds.name
      } with arguments ${archiveOrDeleteSelectEmailThreadIds.arguments.toString()}`
    );

  let rowNumber: number = 2;
  automatedReceivedSheetData.forEach((row) => {
    const emailThreadId = row[emailThreadIdCol.colNumber - 1];
    const archiveCheckBox = row[archiveThreadIdCol.colNumber - 1];
    const deleteCheckBox = row[deleteThreadIdCol.colNumber - 1];
    const removeLabelCheckbox = row[removeGmailLabelIdCol.colNumber - 1];

    if (
      (type === 'archive' && archiveCheckBox === true) ||
      (type === 'delete' && deleteCheckBox === true) ||
      (type === 'remove gmail label' && removeLabelCheckbox === true)
    ) {
      emailThreadIdsMap.set(emailThreadId, rowNumber);
      if (type === 'archive') {
        let archiveGmailLabel = GmailApp.getUserLabelByName(ARCHIVE_LABEL_NAME);
        if (!archiveGmailLabel) {
          archiveGmailLabel = GmailApp.createLabel(ARCHIVE_LABEL_NAME);
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

  emailThreadIdsMap.forEach((rowNumber, _emailThreadIdToDelete) => {
    automatedReceivedSheet.deleteRow(rowNumber);
  });
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

export function writeEmailsListToAutomationSheet(emailsForList: EmailListItem[]) {
  const autoResultsListSheet = getSheetByName(`${AUTOMATED_RECEIVED_SHEET_NAME}`);
  if (!autoResultsListSheet) throw Error(`Cannot find ${AUTOMATED_RECEIVED_SHEET_NAME} Sheet`);

  if (emailsForList.length > 0) {
    const columnsObject = findColumnNumbersOrLettersByHeaderNames({
      sheetName: AUTOMATED_RECEIVED_SHEET_NAME,
      headerName: ['Date of Email', 'Archive Thread Id', 'Warning: Delete Thread Id', 'Remove Gmail Label'],
    });

    const dateCol = columnsObject['Date of Email'];
    const archiveThreadIdCol = columnsObject['Archive Thread Id'];
    const deleteThreadIdCol = columnsObject['Warning: Delete Thread Id'];
    const removeGmalLabelCol = columnsObject['Remove Gmail Label'];
    if (!dateCol || !archiveThreadIdCol || !deleteThreadIdCol || !removeGmalLabelCol)
      throw Error(
        `Missing header / column in ${findColumnNumbersOrLettersByHeaderNames.name}, function ${writeEmailsListToAutomationSheet.name}`
      );

    const { numCols, numRows } = getNumRowsAndColsFromArray(emailsForList);
    addRowsToTopOfSheet(numRows, autoResultsListSheet);
    setValuesInRangeAndSortSheet(numRows, numCols, emailsForList, autoResultsListSheet, {
      sortByCol: dateCol.colNumber,
      asc: false,
    });
    setCheckedValueForEachRow(
      emailsForList.map((_) => [false]),
      autoResultsListSheet,
      archiveThreadIdCol.colNumber
    );
    setCheckedValueForEachRow(
      emailsForList.map((_) => [false]),
      autoResultsListSheet,
      deleteThreadIdCol.colNumber
    );
    setCheckedValueForEachRow(
      emailsForList.map((_) => [false]),
      autoResultsListSheet,
      removeGmalLabelCol.colNumber
    );
    setSheetProtection(autoResultsListSheet, 'Automated Results List Protection', [
      archiveThreadIdCol.colLetter,
      deleteThreadIdCol.colLetter,
      removeGmalLabelCol.colLetter,
    ]);
  }
}

export function initSpreadsheet() {
  checkExistsOrCreateSpreadsheet();

  setActiveSpreadSheet(SpreadsheetApp);

  getAndSetActiveSheetByName(AUTOMATED_RECEIVED_SHEET_NAME);
}
