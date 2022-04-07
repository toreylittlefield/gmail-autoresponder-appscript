import { checkExistsOrCreateSpreadsheet, WarningResetSheetsAndSpreadsheet } from '../sheets/sheets';

const menuName = `Autoresponder Email Settings Menu`;

function createMenuAfterStart(ui: GoogleAppsScript.Base.Ui, menu: GoogleAppsScript.Base.Menu) {
  const optionsMenu = ui.createMenu('Options');
  optionsMenu.addItem('Reset Entire Sheet', 'menuItemResetEntireSheet');

  menu
    .addItem(`Toggle Automatic Email Sending`, 'toggleAutoResponseOnOff')
    .addSeparator()
    .addSubMenu(optionsMenu)
    .addToUi();
}

export function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const hasSpreadsheetId = PropertiesService.getUserProperties().getProperty('spreadsheetId');

  const menu = ui.createMenu(menuName);

  if (!hasSpreadsheetId) {
    createMenuAfterStart(ui, menu);
  } else {
    menu.addItem(`Setup and Create Sheets`, `initializeSpreadsheets`).addToUi();
  }
}

export function initializeSpreadsheets() {
  const ui = SpreadsheetApp.getUi(); // Or DocumentApp or FormApp.
  const response = ui.alert(
    `Create Sheets!`,
    `This will create the sheets you need to run automations.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    checkExistsOrCreateSpreadsheet;
    SpreadsheetApp.getActiveSpreadsheet().removeMenu(menuName);
    const menu = ui.createMenu(menuName);
    createMenuAfterStart(ui, menu);
  }
}

export function toggleAutoResponseOnOff() {
  const ui = SpreadsheetApp.getUi();
  const isAutoResOn = PropertiesService.getUserProperties().getProperty('isAutoResOn');
  const onOrOff = isAutoResOn === 'On' ? 'Off' : 'On';
  const response = ui.alert(
    `Confirm: Turn Automatic Emailing ${onOrOff}?`,
    `
If automatic emailing is "ON": 
    Responses will be sent automatically without any action from you.
  
  
If automatic emailing is "OFF": 
    You can send emails by checking them in the "Pending Emails To Send" sheet and then by clicking the "Send Selected Emails" button.
  `,
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) {
    const newValue = isAutoResOn === 'On' ? 'Off' : 'On';
    PropertiesService.getUserProperties().setProperty('isAutoResOn', newValue);
    ui.alert(`${newValue}`, `Automatic Emailing Is Now ${newValue}`, ui.ButtonSet.OK);
  }
}

export function menuItemResetEntireSheet() {
  const ui = SpreadsheetApp.getUi(); // Or DocumentApp or FormApp.
  const response = ui.alert(
    `Warning!`,
    `This will reset the entire spreadsheet and delete all the data. You cannot recover the data. You'll have to run the initialization again.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    WarningResetSheetsAndSpreadsheet();
  }
}
