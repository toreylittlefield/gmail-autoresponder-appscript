import { getUserCalendarsAndCurrentCalendar } from '../calendar/calendar';
import {
  createFilterAndLabel,
  getUserCannedMessage,
  getUserEmails,
  getUserLabels,
  getUserNameForEmail,
} from '../email/email';
import { uiGetEmailsFromGmail } from '../index';
import {
  getSingleUserPropValue,
  getUserProps,
  setUserProps,
  UserRecords,
} from '../properties-service/properties-service';
import {
  archiveDeleteAddOrRemoveGmailLabelsInFollowUpSheet,
  archiveOrDeleteSelectEmailThreadIds,
  checkExistsOrCreateSpreadsheet,
  manuallyCreateEmailForSelectedRowsInReceivedSheet,
  manuallyMoveToFollowUpSheet,
  sendDraftsIfAutoResponseUserOptionIsOn,
  sendOrMoveManuallyOrDeleteDraftsInPendingSheet,
  WarningResetSheetsAndSpreadsheet,
} from '../sheets/sheets';
import {
  ARCHIVED_FOLLOW_UP_SHEET_NAME,
  FOLLOW_UP_MESSAGES_LABEL_NAME,
  RECEIVED_MESSAGES_ARCHIVE_LABEL_NAME,
  SENT_MESSAGES_ARCHIVE_LABEL_NAME,
  SENT_MESSAGES_LABEL_NAME,
} from '../variables/publicvariables';

const menuName = `Autoresponder Email Settings Menu`;

function createMenuAfterStart(ui: GoogleAppsScript.Base.Ui, menu: GoogleAppsScript.Base.Menu) {
  const optionsMenu = ui.createMenu('Global Options');
  optionsMenu.addItem(`Toggle Automatic Email Sending`, toggleAutoResponseOnOff.name);
  optionsMenu.addSeparator();
  optionsMenu.addItem('Warning: Reset Entire Spreadsheet & Delete Pending Drafts', menuItemResetEntireSheet.name);

  const pendingSheetActions = ui.createMenu('Pending Sheet Actions');
  pendingSheetActions.addItem(`Send Selected Pending Draft Emails`, sendSelectedEmailsInPendingEmailsSheet.name);
  pendingSheetActions.addItem(`Delete Selected Pending Draft Emails`, deleteSelectedEmailsInPendingEmailsSheet.name);
  pendingSheetActions.addItem(
    `Move Selected Pending Draft Emails To Sent Sheet`,
    moveManuallySelectedEmailsInPendingEmailsSheet.name
  );

  const receivedEmailsSheetActions = ui.createMenu('Received Sheet Actions');
  receivedEmailsSheetActions.addItem(`Move Selected To Follow Up Sheet`, uiButtonMoveSelectedToFollowUpSheet.name);
  receivedEmailsSheetActions.addItem(
    `Force Create Draft Emails in Pending Sheet For Selected Rows`,
    uiButtonManuallyCreateDraftEmailsForSelectedRowsInAutoReceivedSheet.name
  );
  receivedEmailsSheetActions.addItem(`Archive Selected Rows`, archiveSelectRowsInAutoReceivedSheet.name);
  receivedEmailsSheetActions.addItem(`Warning: Delete Selected Rows`, deleteSelectRowsInAutoReceivedSheet.name);
  receivedEmailsSheetActions.addItem(
    `Warning: Remove Label Selected Rows`,
    removeLabelSelectRowsInAutoReceivedSheet.name
  );

  const followUpSheetActions = ui.createMenu('Follow Up Sheet Actions');
  followUpSheetActions.addItem(`Archived Follow Up Messages`, uiButtonArchiveFollowUp.name);
  followUpSheetActions.addItem(`Warning: Delete Selected Email Threads`, uiButtonDeleteFollowUp.name);
  followUpSheetActions.addItem(`Remove From GMAIL Sent Message Label`, uiButtonRemoveLabelFollowUp.name);
  followUpSheetActions.addItem(`Add GMAIL Follow Up Label`, uiButtonAddLabelFollowUp.name);

  menu.addItem(`Get Emails & Create Drafts - Sync Emails`, uiGetEmailsFromGmail.name).addSeparator();
  menu.addSubMenu(receivedEmailsSheetActions).addToUi();
  menu.addSubMenu(pendingSheetActions).addToUi();
  menu.addSubMenu(followUpSheetActions).addToUi();
  menu.addSeparator().addItem('User Configuration', userConfigurationModal.name);
  menu.addSeparator().addSubMenu(optionsMenu).addToUi();
}

export function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const hasSpreadsheetId = getSingleUserPropValue('spreadsheetId');
  const menu = ui.createMenu(menuName);

  if (hasSpreadsheetId) {
    createMenuAfterStart(ui, menu);
  } else {
    menu.addItem(`Setup and Create Sheets`, initializeSpreadsheets.name).addToUi();
  }
}

export async function initializeSpreadsheets() {
  const ui = SpreadsheetApp.getUi();
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

export function uiButtonArchiveFollowUp() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Archived Selected In Follow Up Sheet`,
    `All rows with the "Archive" checkbox will be moved to the "Archive Follow Up" sheet. Use this to clean up rows you don't want to see any more.
    
    Archiving applies a GMAIL label ${SENT_MESSAGES_ARCHIVE_LABEL_NAME} to the email thread in Gmail. This action means it will not appear again in the ${ARCHIVED_FOLLOW_UP_SHEET_NAME} emails sheet. 
    To undo this you'll have to manually remove the label in GMAIL and run "Get Emails" again`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    archiveDeleteAddOrRemoveGmailLabelsInFollowUpSheet({ type: 'archive' });
  }
}

export function uiButtonDeleteFollowUp() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Warning: Delete Selected Rows In Follow Up Sheet`,
    `All rows with the "Delete" checkbox will be deleted from the sheet. Use this to clean up rows you don't need.
    
    Deleting will ALSO DELETE that email / thread in GMAIL by moving it to the trash in GMAIL. So be careful with this option.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    archiveDeleteAddOrRemoveGmailLabelsInFollowUpSheet({ type: 'delete' });
  }
}

export function uiButtonRemoveLabelFollowUp() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Remove GMAIL ${SENT_MESSAGES_LABEL_NAME} Label Selected Rows`,
    `All rows with the "Remove Gmail Label" will have the ${SENT_MESSAGES_LABEL_NAME} removed from them in GMAIL and it will delete the row in the sheet. 
    
    Use this so that you keep the email from appearing in this spreadsheet when a "Get Emails" sync is run.
    
    This action means it will not appear again in the spreadsheet even if there is a follow up email or a reply to a email you've already sent. 
    To undo this you'll have to apply the label again in GMAIL and run "Get Emails" again`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    archiveDeleteAddOrRemoveGmailLabelsInFollowUpSheet({ type: 'remove gmail label' });
  }
}

export function uiButtonAddLabelFollowUp() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Add GMAIL ${FOLLOW_UP_MESSAGES_LABEL_NAME} Label Selected Rows`,
    `All rows with the "Add Gmail Label" will have the ${FOLLOW_UP_MESSAGES_LABEL_NAME} added from them in GMAIL.
     
    To undo this you'll have to rempve the label in GMAIL`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    archiveDeleteAddOrRemoveGmailLabelsInFollowUpSheet({ type: 'add gmail label' });
  }
}

export function uiButtonMoveSelectedToFollowUpSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Move Selected Rows To Follow Up Sheet`,
    `All rows with the "Manually Move To Follow Up Emails" checkbox will be moved to the "Follow Up Emails" Sheet and will be updated to also have the "auto-responder-sent-email-label" GMAIL label added to each sent thread.
    
    Note: You cannot undo this action. Use this action for emails that you believe should be marked as Follow Up Emails and not emails that need an autoresponse message`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    manuallyMoveToFollowUpSheet();
  }
}
export function uiButtonManuallyCreateDraftEmailsForSelectedRowsInAutoReceivedSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Create Draft Emails For Selected Rows`,
    `All rows with the "Manually Create Pending Email" checkbox will be have a draft email created in the to the "Pending Emails To Send" sheet.
    
    Note: if a draft email in the pending sheet already exists for this email thread Id then the draft emails will not be created`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    manuallyCreateEmailForSelectedRowsInReceivedSheet();
  }
}

export function archiveSelectRowsInAutoReceivedSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Archive Selected Rows`,
    `All rows with the "Archive" checkbox will be moved to the "Archive" sheet. Use this to clean up rows you don't want to see any more.
    
    Archiving applies a GMAIL label ${RECEIVED_MESSAGES_ARCHIVE_LABEL_NAME} to the email thread in Gmail. This action means it will not appear again in the recieved emails sheet. 
    To undo this you'll have to manually remove the label in GMAIL and run "Get Emails" again`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    archiveOrDeleteSelectEmailThreadIds({ type: 'archive' });
  }
}
export function deleteSelectRowsInAutoReceivedSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Warning: Delete Selected Rows`,
    `All rows with the "Delete" checkbox will be deleted from the sheet. Use this to clean up rows you don't need.
    
    Deleting will ALSO DELETE that email / thread in GMAIL by moving it to the trash in GMAIL. So be careful with this option.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    archiveOrDeleteSelectEmailThreadIds({ type: 'delete' });
  }
}
export function removeLabelSelectRowsInAutoReceivedSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Remove GMAIL Label Selected Rows`,
    `All rows with the "Remove Gmail Label" will have the ${getSingleUserPropValue(
      'labelToSearch'
    )} removed from them in GMAIL and it will delete the row in the sheet. 
    
    Use this so that you keep the email from appearing in this spreadsheet when a "Get Emails" sync is run.
    
    This action means it will not appear again in the spreadsheet even if there is a follow up email or a reply to a email you've already sent. 
    To undo this you'll have to apply the label again in GMAIL and run "Get Emails" again`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response === ui.Button.OK) {
    archiveOrDeleteSelectEmailThreadIds({ type: 'remove gmail label' });
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
    sendDraftsIfAutoResponseUserOptionIsOn();
    ui.alert(`${newValue}`, `Automatic Emailing Is Now ${newValue}`, ui.ButtonSet.OK);
    newValue === 'On' &&
      ui.alert(
        `Trigger Created`,
        `Any pending draft emails that are selected for "SEND" will be automatically sent every 1 hour`,
        ui.ButtonSet.OK
      );
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
    const { draftId, email, filterId, labelId, labelToSearch, messageId, nameForEmail, spreadsheetId, subject } =
      getUserProps([
        'draftId',
        'email',
        'filterId',
        'labelId',
        'labelToSearch',
        'messageId',
        'nameForEmail',
        'spreadsheetId',
        'subject',
      ]);
    WarningResetSheetsAndSpreadsheet();
    const deleteRes = ui.alert(
      `Delete Stored User Configuration`,
      `Would you like to DELETE / RESET all saved "User Configurations"?`,
      ui.ButtonSet.YES_NO
    );
    if (deleteRes === ui.Button.YES) {
      ui.alert(
        `Next Steps`,
        `Please reconfigure your "User Configuration" settings before syncing your emails`,
        ui.ButtonSet.OK
      );
    } else {
      setUserProps({
        draftId,
        email,
        filterId,
        labelId,
        labelToSearch,
        messageId,
        nameForEmail,
        spreadsheetId,
        subject,
      });
    }
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
  SpreadsheetApp.getUi().showModalDialog(html, 'User Options');
}

export function getUserPropertiesForPageModal() {
  const { currentEmailUserStore, emailAliases, mainEmail } = getUserEmails();
  const { nameForEmail } = getUserNameForEmail();
  const { subject, draftsList } = getUserCannedMessage();
  const { currentLabel, userLabels } = getUserLabels();
  const { currentCalendar, listOfOwnerCalendarNames } = getUserCalendarsAndCurrentCalendar();

  return {
    emailAliases,
    mainEmail,
    currentEmailUserStore,
    nameForEmail,
    subject,
    draftsList,
    currentLabel,
    userLabels,
    currentCalendar,
    listOfOwnerCalendarNames,
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
  if (formObject.currentCalendarName) {
    setUserProps({ currentCalendarName: formObject.currentCalendarName });
    return getUserProps(['currentCalendarName']);
  }
  return;
}
