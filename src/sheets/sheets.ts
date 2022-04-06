import { getToEmailArray } from '../email/email';
import { getProps, setUserProps } from '../properties-service/properties-service';
import {
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
  SENT_SHEET_NAME,
  SENT_SHEET_NAME_HEADERS,
  SPREADSHEET_NAME,
} from '../variables/publicvariables';

export type SheetNames =
  | typeof AUTOMATED_SHEET_NAME
  | typeof SENT_SHEET_NAME
  | typeof BOUNCED_SHEET_NAME
  | typeof DO_NOT_EMAIL_AUTO_SHEET_NAME
  | typeof DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME;

type ReplyToUpdateType = [row: number, emailMessage: GoogleAppsScript.Gmail.GmailMessage[]][];

const tabColors = ['blue', 'green', 'red', 'purple', 'orange', 'yellow'] as const;

export let activeSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
export let activeSheet: GoogleAppsScript.Spreadsheet.Sheet;
export const doNotSendMailAutoMap = new Map<string, number>();
export const doNotTrackMap = new Map<string, boolean>();
export const repliesToUpdateArray: ReplyToUpdateType = [];

function checkExistsOrCreateSpreadsheet() {
  let { spreadsheetId } = getSpreadSheetId();

  if (spreadsheetId) {
    try {
      SpreadsheetApp.openById(spreadsheetId);
      const file = DriveApp.getFileById(spreadsheetId);
      if (file.isTrashed()) throw Error('File in trash, creating a new sheet');
    } catch (error) {
      PropertiesService.getUserProperties().deleteProperty('spreadsheetId');
      PropertiesService.getUserProperties().deleteProperty('sheetId');
      PropertiesService.getUserProperties().deleteProperty('sheetName');
      PropertiesService.getUserProperties().deleteAllProperties();
      spreadsheetId = null;
    }
  }

  if (!spreadsheetId) {
    const spreadsheet = SpreadsheetApp.create(SPREADSHEET_NAME, 2, AUTOMATED_SHEET_HEADERS.length);
    const [firstSheet] = spreadsheet.getSheets();
    setUserProps({
      spreadsheetId: spreadsheet.getId(),
    });
    createSheet(spreadsheet, AUTOMATED_SHEET_NAME, AUTOMATED_SHEET_HEADERS, { tabColor: 'blue' });
    spreadsheet.deleteSheet(firstSheet);
    createSheet(spreadsheet, SENT_SHEET_NAME, SENT_SHEET_NAME_HEADERS, { tabColor: 'green' });
    createSheet(spreadsheet, DO_NOT_EMAIL_AUTO_SHEET_NAME, DO_NOT_EMAIL_AUTO_SHEET_HEADERS, {
      tabColor: 'purple',
      initData: DO_NOT_EMAIL_AUTO_INITIAL_DATA,
    });
    createSheet(spreadsheet, BOUNCED_SHEET_NAME, BOUNCED_SHEET_NAME_HEADERS, { tabColor: 'red' });
    createSheet(spreadsheet, DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME, DO_NOT_TRACK_DOMAIN_LIST_HEADERS, {
      tabColor: 'orange',
      initData: DO_NOT_TRACK_DOMAIN_LIST_INITIAL_DATA,
    });
  }
}

type Options = { tabColor: typeof tabColors[number]; initData: any[][] };

function createSheet(
  activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: SheetNames,
  headersValues: string[],
  options: Partial<Options> = { initData: [], tabColor: 'green' }
) {
  const { initData = [], tabColor = '' } = options;
  const sheet = activeSS.insertSheet();
  sheet.setName(sheetName);
  sheet.setTabColor(tabColor || 'green');
  writeHeaders(sheet, headersValues);
  initData.length > 0 && setInitialSheetData(sheet, headersValues, initData);
  setSheetProtection(sheet, `${sheetName} Protected Range`);
  setSheetInProps(sheetName, sheet);
}

function getSpreadSheetId() {
  const { spreadsheetId } = getProps(['spreadsheetId']);
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

function setSheetInProps(sheetName: string, activeSheet: GoogleAppsScript.Spreadsheet.Sheet) {
  setUserProps({ [sheetName]: activeSheet.getSheetId().toString() });
}

function getSheet(sheetName: SheetNames) {
  try {
    if (activeSpreadsheet == null) throw Error('No active spreadsheet in getsheet method');
    const [sheetNameProp, sheetIdProp] = Object.entries(getProps([sheetName]))[0];

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
    if (activeSpreadsheet == null) throw Error('No active spreadsheet in method getSheetByName');
    const [_, sheetValueId] = Object.entries(getProps([sheetName]))[0];
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

function setSheetProtection(sheet: GoogleAppsScript.Spreadsheet.Sheet, description: string) {
  const [protection] = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (!protection) {
    sheet.protect().setWarningOnly(true).setDescription(description);
  }
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

export function setGlobalDoNotSendEmailAutoArrayList() {
  try {
    const doNotReplyList = getAllDataFromSheet('Do Not Autorespond List');
    if (!doNotReplyList) throw Error('Could Not Set The Global Do Not Reply Array');
    doNotReplyList.forEach(([domain, _, count]) => doNotSendMailAutoMap.set(domain, count));
  } catch (error) {
    console.error(error as any);
  }
}

export function formatRowHeight() {
  const automatedSheet = getSheetByName('Automated Results List');
  //@ts-expect-error
  automatedSheet && automatedSheet.setRowHeightsForced(2, automatedSheet.getDataRange().getNumRows(), 21);
  const sentResponsesSheet = getSheetByName('Sent Automated Responses');
  //@ts-expect-error
  sentResponsesSheet && sentResponsesSheet.setRowHeightsForced(2, sentResponsesSheet.getDataRange().getNumRows(), 21);
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

export function initSpreadsheet() {
  checkExistsOrCreateSpreadsheet();

  setActiveSpreadSheet(SpreadsheetApp);

  setGlobalDoNotSendEmailAutoArrayList();
  getAndSetActiveSheetByName(AUTOMATED_SHEET_NAME);

  // if (
  //   !getHeaders(sheet, AUTOMATED_SHEET_HEADERS)[0].every(
  //     (colVal, index) => colVal === AUTOMATED_SHEET_HEADERS[index]
  //   )
  // ) {
  //   writeHeaders(sheet, AUTOMATED_SHEET_HEADERS);
  // }
}
