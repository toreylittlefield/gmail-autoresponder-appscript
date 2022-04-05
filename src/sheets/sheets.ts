import { getProps, setUserProps } from '../properties-service/properties-service';
import {
  AUTOMATED_SHEET_HEADERS,
  AUTOMATED_SHEET_NAME,
  BOUNCED_SHEET_NAME,
  BOUNCED_SHEET_NAME_HEADERS,
  DO_NOT_EMAIL_AUTO_INITIAL_DATA,
  DO_NOT_EMAIL_AUTO_SHEET_HEADERS,
  DO_NOT_EMAIL_AUTO_SHEET_NAME,
  SENT_SHEET_NAME,
  SENT_SHEET_NAME_HEADERS,
  SPREADSHEET_NAME,
} from '../variables/publicvariables';

const tabColors = ['blue', 'green', 'red', 'purple'];

export let activeSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
export let activeSheet: GoogleAppsScript.Spreadsheet.Sheet;

export type SheetNames =
  | typeof AUTOMATED_SHEET_NAME
  | typeof SENT_SHEET_NAME
  | typeof BOUNCED_SHEET_NAME
  | typeof DO_NOT_EMAIL_AUTO_SHEET_NAME;

function checkExistsOrCreateSpreadsheet() {
  let { spreadsheetId } = getSpreadSheetId();

  if (spreadsheetId) {
    try {
      SpreadsheetApp.openById(spreadsheetId);
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
    createSheet(spreadsheet, AUTOMATED_SHEET_NAME, AUTOMATED_SHEET_HEADERS, 'blue');
    spreadsheet.deleteSheet(firstSheet);
    createSheet(spreadsheet, SENT_SHEET_NAME, SENT_SHEET_NAME_HEADERS, 'green');
    createSheet(spreadsheet, DO_NOT_EMAIL_AUTO_SHEET_NAME, DO_NOT_EMAIL_AUTO_SHEET_HEADERS, 'purple');
    createSheet(spreadsheet, BOUNCED_SHEET_NAME, BOUNCED_SHEET_NAME_HEADERS, 'red');
  }
}

function createSheet(
  activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: SheetNames,
  headersValues: string[],
  tabColor?: string
) {
  const sheet = activeSS.insertSheet();
  sheet.setName(sheetName);
  sheet.setTabColor(tabColor || tabColors.shift() || 'green');
  writeHeaders(sheet, headersValues);
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

function setInitialDoNotReplyData() {
  try {
    const doNotEmailSheet = getSheetByName(DO_NOT_EMAIL_AUTO_SHEET_NAME);
    if (!doNotEmailSheet) throw Error('Cannot Find Do Not Email Sheet');
    const initData = doNotEmailSheet
      .getRange(2, 2, DO_NOT_EMAIL_AUTO_INITIAL_DATA.length, DO_NOT_EMAIL_AUTO_INITIAL_DATA.length)
      .getValues();
    const hasInitialData = initData.every((row, index) => row[index] === DO_NOT_EMAIL_AUTO_INITIAL_DATA[index]);
    if (!hasInitialData) {
      DO_NOT_EMAIL_AUTO_INITIAL_DATA.forEach((row) => {
        doNotEmailSheet.appendRow(row);
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

export function formatRowHeight() {
  const automatedSheet = getSheetByName('Automated Results List');
  //@ts-expect-error
  automatedSheet && automatedSheet.setRowHeightsForced(2, automatedSheet.getDataRange().getNumRows(), 21);
  const sentResponsesSheet = getSheetByName('Sent Automated Responses');
  //@ts-expect-error
  sentResponsesSheet && sentResponsesSheet.setRowHeightsForced(2, sentResponsesSheet.getDataRange().getNumRows(), 21);
}

export function initSpreadsheet() {
  checkExistsOrCreateSpreadsheet();

  setActiveSpreadSheet(SpreadsheetApp);

  setInitialDoNotReplyData();
  getAndSetActiveSheetByName(AUTOMATED_SHEET_NAME);

  // if (
  //   !getHeaders(sheet, AUTOMATED_SHEET_HEADERS)[0].every(
  //     (colVal, index) => colVal === AUTOMATED_SHEET_HEADERS[index]
  //   )
  // ) {
  //   writeHeaders(sheet, AUTOMATED_SHEET_HEADERS);
  // }
}
