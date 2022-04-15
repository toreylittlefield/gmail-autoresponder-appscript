import { createOrSentTemplateEmail, DraftAttributeArray, EmailListItem, getToEmailArray } from '../email/email';
import { alwaysAllowMap, doNotSendMailAutoMap, emailsToAddToPendingSheet } from '../global/maps';
import { getUserProps, setUserProps } from '../properties-service/properties-service';
import { getDomainFromEmailAddress, initialGlobalMap } from '../utils/utils';
import {
  allSheets,
  ALWAYS_RESPOND_DOMAIN_LIST_HEADERS,
  ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME,
  ALWAYS_RESPOND_LIST_INITIAL_DATA,
  AUTOMATED_SHEET_HEADERS,
  AUTOMATED_SHEET_NAME,
  BOUNCED_SHEET_NAME,
  BOUNCED_SHEET_NAME_HEADERS,
  DO_NOT_EMAIL_AUTO_INITIAL_DATA,
  DO_NOT_EMAIL_AUTO_SHEET_HEADERS,
  DO_NOT_EMAIL_AUTO_SHEET_NAME,
  DO_NOT_TRACK_DOMAIN_LIST_HEADERS,
  DO_NOT_TRACK_DOMAIN_LIST_INITIAL_DATA,
  DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME,
  FOLLOW_UP_EMAILS_HEADERS,
  FOLLOW_UP_EMAILS_SHEET_NAME,
  PENDING_EMAILS_TO_SEND_HEADERS,
  PENDING_EMAILS_TO_SEND_SHEET_NAME,
  SENT_SHEET_NAME,
  SENT_SHEET_NAME_HEADERS,
} from '../variables/publicvariables';

export type SheetNames =
  | typeof AUTOMATED_SHEET_NAME
  | typeof SENT_SHEET_NAME
  | typeof FOLLOW_UP_EMAILS_SHEET_NAME
  | typeof BOUNCED_SHEET_NAME
  | typeof ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME
  | typeof DO_NOT_EMAIL_AUTO_SHEET_NAME
  | typeof DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME
  | typeof PENDING_EMAILS_TO_SEND_SHEET_NAME;

const tabColors = ['blue', 'green', 'red', 'purple', 'orange', 'yellow', 'black', 'teal', 'gold'] as const;

export let activeSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
export let activeSheet: GoogleAppsScript.Spreadsheet.Sheet;

type ReplyToUpdateType = [row: number, emailMessage: GoogleAppsScript.Gmail.GmailMessage[]][];
export const repliesToUpdateArray: ReplyToUpdateType = [];

export function WarningResetSheetsAndSpreadsheet() {
  try {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    sendOrMoveManuallyOrDeleteDraftsInPendingSheet({ type: 'delete' }, { deleteAll: true });
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

export async function checkExistsOrCreateSpreadsheet(): Promise<'done'> {
  return new Promise((resolve, _reject) => {
    let { spreadsheetId } = getSpreadSheetId();

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

    if (!spreadsheetId) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

      setUserProps({
        spreadsheetId: spreadsheet.getId(),
      });
      createSheet(spreadsheet, AUTOMATED_SHEET_NAME, AUTOMATED_SHEET_HEADERS, {
        tabColor: 'blue',
      });
      createSheet(spreadsheet, PENDING_EMAILS_TO_SEND_SHEET_NAME, PENDING_EMAILS_TO_SEND_HEADERS, {
        tabColor: 'gold',
        unprotectColumnLetters: ['A', 'L', 'U'],
      });
      createSheet(spreadsheet, SENT_SHEET_NAME, SENT_SHEET_NAME_HEADERS, { tabColor: 'green' });
      createSheet(spreadsheet, FOLLOW_UP_EMAILS_SHEET_NAME, FOLLOW_UP_EMAILS_HEADERS, { tabColor: 'black' });
      createSheet(spreadsheet, BOUNCED_SHEET_NAME, BOUNCED_SHEET_NAME_HEADERS, { tabColor: 'red' });
      createSheet(spreadsheet, ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME, ALWAYS_RESPOND_DOMAIN_LIST_HEADERS, {
        tabColor: 'teal',
        initData: ALWAYS_RESPOND_LIST_INITIAL_DATA,
      });
      createSheet(spreadsheet, DO_NOT_EMAIL_AUTO_SHEET_NAME, DO_NOT_EMAIL_AUTO_SHEET_HEADERS, {
        tabColor: 'purple',
        initData: DO_NOT_EMAIL_AUTO_INITIAL_DATA,
      });
      createSheet(spreadsheet, DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME, DO_NOT_TRACK_DOMAIN_LIST_HEADERS, {
        tabColor: 'orange',
        initData: DO_NOT_TRACK_DOMAIN_LIST_INITIAL_DATA,
      });
    }
    resolve('done');
  });
}

type Options = { tabColor: typeof tabColors[number]; initData: any[][]; unprotectColumnLetters?: string | string[] };

function createSheet(
  activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: SheetNames,
  headersValues: string[],
  options: Partial<Options>
) {
  const { initData = [], tabColor = 'green', unprotectColumnLetters = undefined } = options;
  const sheet = activeSS.insertSheet();
  sheet.setName(sheetName);
  sheet.setTabColor(tabColor);
  writeHeaders(sheet, headersValues);
  initData.length > 0 && setInitialSheetData(sheet, headersValues, initData);
  sheet.deleteColumns(headersValues.length + 1, sheet.getMaxColumns() - headersValues.length);

  const startRowIndex = initData.length > 50 ? initData.length : 50;
  sheet.deleteRows(startRowIndex, sheet.getMaxRows() - startRowIndex);

  sheet.autoResizeColumns(1, headersValues.length);
  setSheetProtection(sheet, `${sheetName} Protected Range`, unprotectColumnLetters);
  setSheetInProps(sheetName, sheet);
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

function writeHeaders(sheet: GoogleAppsScript.Spreadsheet.Sheet, headerValues: string[]) {
  const headers = sheet.getRange(1, 1, 1, headerValues.length);
  headers.setValues([headerValues]);
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

function setInitialSheetData(sheet: GoogleAppsScript.Spreadsheet.Sheet, headers: string[], initialData: any[][]) {
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
  sheet && sheet.setRowHeightsForced(2, sheet.getDataRange().getNumRows(), 21);
}

export function addToRepliesArray(index: number, emailMessages: GoogleAppsScript.Gmail.GmailMessage[]) {
  repliesToUpdateArray.push([index + 2, emailMessages]);
}

export function updateRepliesColumn(rowsToUpdate: ReplyToUpdateType) {
  try {
    const automatedSheet = getSheetByName('Automated Results List');
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
  personFrom: string,
  emailSubject: string,
  emailBody: string,
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
  const pendingEmailsSheet = getSheetByName('Pending Emails To Send');
  if (!pendingEmailsSheet) {
    throw Error(`Could Not Find Pending Emails To Send Sheet`);
  }
  const emailObjects = Array.from(emailsToAddToPendingSheet.values());
  const validSentRows: ValidRowToWriteInPendingSheet[] = emailObjects.map(
    ({
      send,
      date,
      emailBody,
      emailFrom,
      emailSubject,
      emailThreadId,
      emailThreadPermaLink,
      inResponseToEmailMessageId,
      personFrom,
      emailSendTo,
      isReplyorNewEmail,
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
        personFrom,
        emailSubject,
        emailBody,
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
  const { numCols, numRows } = getNumRowsAndColsFromArray(validSentRows);
  addRowsToTopOfSheet(numRows, pendingEmailsSheet);
  setValuesInRangeAndSortSheet(
    numRows,
    numCols,
    validSentRows,
    pendingEmailsSheet,
    { asc: false, sortByCol: 5 },
    { startCol: 1, startRow: 2 }
  );
  setCheckedValueForEachRow(
    emailObjects.map(({ send }) => [send]),
    pendingEmailsSheet,
    1
  );
  setCheckedValueForEachRow(
    emailObjects.map((_row) => [false]),
    pendingEmailsSheet,
    12
  );
  setCheckedValueForEachRow(
    emailObjects.map((_row) => [false]),
    pendingEmailsSheet,
    21
  );
  setSheetProtection(pendingEmailsSheet, 'Pending Emails To Send Protected Range', ['A', 'L', 'U']);
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
  personFrom: string,
  emailSubject: string,
  emailBody: string,
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

export function sendOrMoveManuallyOrDeleteDraftsInPendingSheet(
  { type }: SendDraftsOptions,
  { deleteAll = false }: DeleteDraftOptions
) {
  try {
    const pendingSheetEmailData = getAllDataFromSheet('Pending Emails To Send') as ValidRowToWriteInPendingSheet[];
    const pendingSheet = getSheetByName('Pending Emails To Send');
    const sentEmailsSheet = getSheetByName('Sent Automated Responses');

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
          _personFrom,
          _emailSubject,
          _emailBody,
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
                _personFrom,
                _emailSubject,
                _emailBody,
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
    setSheetProtection(pendingSheet, 'Pending Emails To Send Protected Range', ['A', 'L', 'U']);
    if (rowsForSentSheet.length > 0) {
      writeSentDraftsToSentEmailsSheet(sentEmailsSheet, rowsForSentSheet);
      writeDomainsListToDoNotRespondSheet();
    }
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
  const domain = getDomainFromEmailAddress(emailAddress);

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
  const { getPermalink, getId } = getThread();
  return { getDate, getId, getEmailMessageId, getPermalink };
}

function writeSentDraftsToSentEmailsSheet(
  sentEmailsSheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowsForSentSheet: ValidRowToWriteInSentSheet[]
) {
  if (sentEmailsSheet.getName() !== 'Sent Automated Responses')
    throw Error('Sheet Must Be The Sent Automated Responses');
  const { numCols, numRows } = getNumRowsAndColsFromArray(rowsForSentSheet);
  addRowsToTopOfSheet(numRows, sentEmailsSheet);
  setValuesInRangeAndSortSheet(numRows, numCols, rowsForSentSheet, sentEmailsSheet, { asc: false, sortByCol: 21 });
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
    const doNotRespondSheet = getSheetByName('Do Not Autorespond List');
    if (!doNotRespondSheet) throw Error(`Could find the Do Not Response List to write the array list to`);
    const existingData = getAllDataFromSheet('Do Not Autorespond List');
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

export function writeEmailsListToAutomationSheet(emailsForList: EmailListItem[]) {
  const autoResultsListSheet = getSheetByName('Automated Results List');
  if (!autoResultsListSheet) throw Error('Cannot find Automated Results List Sheet');
  if (emailsForList.length > 0) {
    const { numCols, numRows } = getNumRowsAndColsFromArray(emailsForList);
    addRowsToTopOfSheet(numRows, autoResultsListSheet);
    setValuesInRangeAndSortSheet(numRows, numCols, emailsForList, autoResultsListSheet, { sortByCol: 3, asc: false });
  }
}

export function initSpreadsheet() {
  checkExistsOrCreateSpreadsheet();

  setActiveSpreadSheet(SpreadsheetApp);

  getAndSetActiveSheetByName(AUTOMATED_SHEET_NAME);
}
