import {
  createFilterAndLabel,
  getUserCannedMessage,
  getUserEmails,
  getUserLabels,
  getUserNameForEmail,
} from '../email/email';
import {
  getSingleUserPropValue,
  getUserProps,
  setUserProps,
  UserRecords,
} from '../properties-service/properties-service';
import {
  checkExistsOrCreateSpreadsheet,
  sendOrMoveManuallyOrDeleteDraftsInPendingSheet,
  WarningResetSheetsAndSpreadsheet,
} from '../sheets/sheets';

const menuName = `Autoresponder Email Settings Menu`;

function createMenuAfterStart(ui: GoogleAppsScript.Base.Ui, menu: GoogleAppsScript.Base.Menu) {
  const optionsMenu = ui.createMenu('Options');
  optionsMenu.addItem(`Toggle Automatic Email Sending`, 'toggleAutoResponseOnOff');
  optionsMenu.addSeparator();

  optionsMenu.addItem('Reset Entire Sheet', 'menuItemResetEntireSheet');

  menu.addItem(`Sync Emails`, 'uiGetEmailsFromGmail');
  menu.addItem(`Send Selected Pending Draft Emails`, 'sendSelectedEmailsInPendingEmailsSheet');
  menu.addItem(`Delete Selected Pending Draft Emails`, 'deleteSelectedEmailsInPendingEmailsSheet');
  menu.addItem(`Move Selected Pending Draft Emails To Sent Sheet`, 'moveManuallySelectedEmailsInPendingEmailsSheet');
  menu.addItem('User Configuration', 'userConfigurationModal');
  menu.addSeparator().addSubMenu(optionsMenu).addToUi();
}

export function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const hasSpreadsheetId = getSingleUserPropValue('spreadsheetId');
  const menu = ui.createMenu(menuName);

  if (hasSpreadsheetId) {
    createMenuAfterStart(ui, menu);
  } else {
    menu.addItem(`Setup and Create Sheets`, `initializeSpreadsheets`).addToUi();
  }
}

export async function initializeSpreadsheets() {
  const ui = SpreadsheetApp.getUi(); // Or DocumentApp or FormApp.
  const response = ui.alert(
    `Create Sheets!`,
    `This will create the sheets you need to run automations.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    await checkExistsOrCreateSpreadsheet().catch((err) => {
      console.error(err);
      WarningResetSheetsAndSpreadsheet();
    });
    ui.alert(
      `Next Steps`,
      `Please fill out the "User Configuration" settings with your email and other options before sync emails`,
      ui.ButtonSet.OK
    );
    SpreadsheetApp.getActiveSpreadsheet().removeMenu(menuName);
    const menu = ui.createMenu(menuName);
    createMenuAfterStart(ui, menu);
  }
}

export function sendSelectedEmailsInPendingEmailsSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Send Selected Drafts`,
    `You are about to SEND any selected / checked draft emails in the "Pending Emails To Send" sheet. The rows for the draft emails will be moved to the "Sent Automated Responses" Sheet`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    sendOrMoveManuallyOrDeleteDraftsInPendingSheet({ type: 'send' }, {});
  }
}

export function moveManuallySelectedEmailsInPendingEmailsSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Move Selected Drafts`,
    `You are about to MOVE any selected / checked draft emails in the "Pending Emails To Send" sheet. 
    This WILL NOT send the selected draft emails to the recipient. You can still manually send the draft email inside of Gmail.
    This action will just simply manually move the selected row(s) in the spreadsheet.
    The rows for the draft emails will be moved to the "Sent Automated Responses" Sheet`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    sendOrMoveManuallyOrDeleteDraftsInPendingSheet({ type: 'manuallyMove' }, {});
  }
}
export function deleteSelectedEmailsInPendingEmailsSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Delete Selected Drafts`,
    `You are about to delete any selected / checked draft emails in the "Pending Emails To Send" sheet. The rows for the draft emails will be delete and you will have to run an email sync again to recreate them`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    sendOrMoveManuallyOrDeleteDraftsInPendingSheet({ type: 'delete' }, {});
  }
}

export function toggleAutoResponseOnOff() {
  const ui = SpreadsheetApp.getUi();
  const isAutoResOn = getSingleUserPropValue('isAutoResOn');
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
    setUserProps({ isAutoResOn: newValue });
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

    ui.alert(
      `Next Steps`,
      `Please reconfigure your "User Configuration" settings before syncing your emails`,
      ui.ButtonSet.OK
    );
  }
}

function newFilterAndLabel(currentEmail: string, ui: GoogleAppsScript.Base.Ui) {
  const { resFilter, resLabel } = createFilterAndLabel(currentEmail);
  if (resFilter.id && resLabel.id) {
    ui.alert(
      `Created Filter ID: ${resFilter.id} in GMAIL with messages to email ${
        (resFilter.criteria as GoogleAppsScript.Gmail.Schema.FilterCriteria).to
      } with automatically have the label: ${resLabel.name} applied to them`
    );
    setUserProps({ labelToSearch: resLabel.name, labelId: resLabel.id, filterId: resFilter.id });
  }
}

export function userConfigurationModal() {
  var html = HtmlService.createHtmlOutputFromFile('dist/Page').setWidth(400).setHeight(500);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModalDialog(html, 'User Options');
}

export function getUserPropertiesForPageModal() {
  const { currentEmailUserStore, emailAliases, mainEmail } = getUserEmails();
  const { nameForEmail } = getUserNameForEmail();
  const { subject, draftsList } = getUserCannedMessage();
  const { currentLabel, userLabels } = getUserLabels();

  return {
    emailAliases,
    mainEmail,
    currentEmailUserStore,
    nameForEmail,
    subject,
    draftsList,
    currentLabel,
    userLabels,
  };
}

export function processFormEventsFromPage(formObject: Partial<Record<UserRecords, string>>) {
  if (formObject.email) {
    setUserProps({ email: formObject.email });
    return getUserProps(['email']);
  }
  if (formObject.labelToSearch) {
    if (formObject.labelToSearch === 'create-label') {
      newFilterAndLabel(getSingleUserPropValue('email') || Session.getActiveUser().getEmail(), SpreadsheetApp.getUi());
    } else {
      setUserProps({ labelToSearch: formObject.labelToSearch });
      return getUserProps(['labelToSearch']);
    }
  }
  if (formObject.nameForEmail) {
    setUserProps({ nameForEmail: formObject.nameForEmail });
    return getUserProps(['nameForEmail']);
  }
  if (formObject.subject && formObject.draftId) {
    setUserProps({ subject: formObject.subject, draftId: formObject.draftId });
    return getUserProps(['subject']);
  }
  return;
}
