import { getProps, setUserProps } from '../properties-service/properties-service';
import { SHEET_NAME, SPREADSHEET_NAME } from '../variables';

const headerValues = [
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

function checkExistsOrCreateSpreadsheet() {
  let { spreadsheetId } = getSpreadSheetId();

  if (spreadsheetId) {
    try {
      SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      PropertiesService.getUserProperties().deleteProperty('spreadsheetId');
      PropertiesService.getUserProperties().deleteProperty('sheetId');
      PropertiesService.getUserProperties().deleteProperty('sheetName');
      spreadsheetId = null;
    }
  }

  if (!spreadsheetId) {
    const spreadsheet = SpreadsheetApp.create(SPREADSHEET_NAME, 2, headerValues.length);
    const [firstSheet] = spreadsheet.getSheets();
    firstSheet.setName(SHEET_NAME);
    firstSheet.setTabColor('blue');
    setUserProps({
      spreadsheetId: spreadsheet.getId(),
      sheetName: firstSheet.getName(),
      sheetId: firstSheet.getSheetId().toString(),
    });
  }
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

function checkIfSheetExists(activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  const sheet = activeSS.getSheetByName(SHEET_NAME);
  if (!sheet) return false;
  return true;
}

function findSheetById(sheetId: number | string, activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  const sheet = activeSS.getSheets().find((sheet) => sheet.getSheetId().toString() === sheetId);
  return sheet;
}

function getSheet(activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  try {
    const { sheetId } = getProps(['sheetId']);
    if (checkIfSheetExists(activeSS) === false && !sheetId) {
      const activeSheet = activeSS.insertSheet(SHEET_NAME);
      setUserProps({ sheetId: activeSheet.getSheetId().toString(), sheetName: activeSheet.getSheetName() });
      return activeSheet;
    } else {
      const sheet = findSheetById(sheetId, activeSS);
      return sheet;
    }
  } catch (error) {
    console.error(error as any);
    return;
  }
}

function getActiveSheet(activeSS: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  try {
    const activeSheet = getSheet(activeSS);
    if (activeSheet) {
      activeSS.setActiveSheet(activeSheet);
      return activeSheet;
    }
    throw Error(`Cannot Find The Sheet`);
  } catch (error) {
    if (error instanceof Error) console.error(error.message);
    else console.log(error as string | object);
    return;
  }
}

function getHeaders(activeSheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const headers = activeSheet.getRange(1, 1, 1, headerValues.length);
  return headers.getValues();
}

function writeHeaders(activeSheet: GoogleAppsScript.Spreadsheet.Sheet) {
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

export function initSpreadsheet() {
  checkExistsOrCreateSpreadsheet();

  const ssApp = SpreadsheetApp;

  setActiveSpreadSheet(ssApp);

  const activeSpreadsheet = ssApp.getActiveSpreadsheet();
  const activeSheet = getActiveSheet(activeSpreadsheet);

  if (!activeSheet) return;

  setSheetProtection(activeSheet, 'Protected Automated Results List');

  if (!getHeaders(activeSheet).every((val) => typeof val === 'string')) {
    writeHeaders(activeSheet);
  }

  return activeSheet;
}
