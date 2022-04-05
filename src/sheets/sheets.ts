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

type SheetNames =
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

function setActiveSpreadSheet(ssApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp) {
  try {
    const { spreadsheetId } = getSpreadSheetId();
    if (!spreadsheetId) throw Error('No Spreadsheet Id');
    const spreadsheet = ssApp.openById(spreadsheetId);
    ssApp.setActiveSpreadsheet(spreadsheet);
  } catch (error) {
    console.error(error as any);
  }
}

function findSheetById(sheetId: number | string, activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  const sheet = activeSS.getSheets().find((sheet) => sheet.getSheetId().toString() === sheetId);
  return sheet;
}

function setSheetInProps(sheetName: string, activeSheet: GoogleAppsScript.Spreadsheet.Sheet) {
  setUserProps({ [sheetName]: activeSheet.getSheetId().toString() });
}

function getSheet(activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet, sheetName: SheetNames) {
  try {
    const [sheetNameProp, sheetIdProp] = Object.entries(getProps([sheetName]))[0];

    const sheet = activeSS.getSheetByName(sheetName);
    if (!sheet || !sheetNameProp) {
      const activeSheet = activeSS.insertSheet(sheetName);
      setSheetInProps(sheetName, activeSheet);
      return activeSheet;
    } else {
      const sheet = findSheetById(sheetIdProp, activeSS);
      return sheet;
    }
  } catch (error) {
    console.error(error as any);
    return;
  }
}

function getActiveSheet(activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet, sheetName: SheetNames) {
  try {
    const sheet = getSheet(activeSS, sheetName);
    if (sheet) {
      activeSS.setActiveSheet(sheet);
      return sheet;
    }
    throw Error(`Cannot Find The Sheet`);
  } catch (error) {
    if (error instanceof Error) console.error(error.message);
    else console.log(error as string | object);
    return;
  }
}

function getHeaders(activeSheet: GoogleAppsScript.Spreadsheet.Sheet, headersValues: string[]) {
  const headers = activeSheet.getRange(1, 1, 1, headersValues.length);
  return headers.getValues();
}

function writeHeaders(activeSheet: GoogleAppsScript.Spreadsheet.Sheet, headerValues: string[]) {
  const headers = activeSheet.getRange(1, 1, 1, headerValues.length);
  headers.setValues([headerValues]);
  activeSheet.setFrozenRows(1);
  headers.setFontWeight('bold');
  activeSheet.getRange(1, headerValues.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

function setSheetProtection(activeSheet: GoogleAppsScript.Spreadsheet.Sheet, description: string) {
  const [protection] = activeSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (!protection) {
    activeSheet.protect().setWarningOnly(true).setDescription(description);
  }
}

function setInitialDoNotReplyData(activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  try {
    const doNotEmailSheet = getActiveSheet(activeSS, DO_NOT_EMAIL_AUTO_SHEET_NAME);
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

export function formatRowHeight(activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  const automatedSheet = getActiveSheet(activeSS, 'Automated Results List');
  automatedSheet && automatedSheet.setRowHeights(2, automatedSheet.getDataRange().getNumRows(), 21);
  const sentResponsesSheet = getActiveSheet(activeSS, 'Sent Automated Responses');
  sentResponsesSheet && sentResponsesSheet.setRowHeights(2, sentResponsesSheet.getDataRange().getNumRows(), 21);
}

export function initSpreadsheet() {
  checkExistsOrCreateSpreadsheet();

  setActiveSpreadSheet(SpreadsheetApp);

  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  setInitialDoNotReplyData(activeSpreadsheet);
  const activeSheet = getActiveSheet(activeSpreadsheet, AUTOMATED_SHEET_NAME);

  if (!activeSheet) return;

  if (
    !getHeaders(activeSheet, AUTOMATED_SHEET_HEADERS)[0].every(
      (colVal, index) => colVal === AUTOMATED_SHEET_HEADERS[index]
    )
  ) {
    writeHeaders(activeSheet, AUTOMATED_SHEET_HEADERS);
  }

  return activeSheet;
}
