import { setDraftTemplateAutoResponder } from '../email/email';
import { runScript } from '../index';
import { getSingleUserPropValue, setUserProps } from '../properties-service/properties-service';
import { checkExistsOrCreateSpreadsheet, WarningResetSheetsAndSpreadsheet } from '../sheets/sheets';

const menuName = `Autoresponder Email Settings Menu`;

function createMenuAfterStart(ui: GoogleAppsScript.Base.Ui, menu: GoogleAppsScript.Base.Menu) {
  const optionsMenu = ui.createMenu('Options');
  optionsMenu.addItem(`Toggle Automatic Email Sending`, 'toggleAutoResponseOnOff');
  optionsMenu.addItem(`Add / Edit Email`, 'setEmail');
  optionsMenu.addItem(`Add / Edit Name To Send In Email`, 'setNameToSendInEmail');
  optionsMenu.addItem(`Add / Edit Canned Message Name`, 'setCannedMessageName');
  optionsMenu.addItem('Reset Entire Sheet', 'menuItemResetEntireSheet');

  menu.addItem(`Sync Emails`, 'runScript');
  menu.addItem(`Send Selected Pending Emails`, 'sendSelectedEmailsInPendingEmailsSheet');
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
    await checkExistsOrCreateSpreadsheet();
    setEmail();
    setCannedMessageName();
    setNameToSendInEmail();
    ui.alert(`Email Sync`, `The script is going to sync your emails`, ui.ButtonSet.OK);
    runScript();
    SpreadsheetApp.getActiveSpreadsheet().removeMenu(menuName);
    const menu = ui.createMenu(menuName);
    createMenuAfterStart(ui, menu);
  }
}

export function setEmail() {
  const ui = SpreadsheetApp.getUi();
  const currentEmail = getSingleUserPropValue('email');
  const mainEmail = Session.getActiveUser().getEmail();
  const emailAliases = GmailApp.getAliases();
  const allEmails =
    emailAliases.length > 0
      ? `Main EMAIL: ${mainEmail}, 
         EMAIL ALIASES: ${emailAliases.map((alias) => `\n ${alias}`)}`
      : `Main email: ${mainEmail}`;
  const message = currentEmail
    ? `Current Email is set to: ${currentEmail} 
    You Can Change This Email.
    
    You can also change to one of the following in your gmail account:
    ${allEmails}
    `
    : `No Email Set. Add An Email. 
    Use one of the following in your gmail account:
    ${allEmails}`;
  const promptResponse = ui.prompt(`Add / Edit Email`, message, ui.ButtonSet.OK_CANCEL);
  const response = promptResponse.getSelectedButton();
  if (response === ui.Button.OK) {
    const newEmail = promptResponse.getResponseText();
    setUserProps({ email: newEmail });
    ui.alert(`Email Changed`, `Your email is now set to: ${getSingleUserPropValue('email')}`, ui.ButtonSet.OK);
  }
}

export function setCannedMessageName() {
  const ui = SpreadsheetApp.getUi();
  const email = getSingleUserPropValue('email');

  if (!email) {
    ui.alert(`Please Set Email`, `You Need To Set An Email Before Setting The Message`, ui.ButtonSet.OK);
    return;
  }

  const drafts = GmailApp.getDrafts();
  const draftsFilteredByEmail = drafts.filter((draft) => {
    const { getFrom, getSubject } = draft.getMessage();
    return getFrom().match(email) && getSubject();
  });

  if (draftsFilteredByEmail.length === 0) {
    ui.alert(
      `No Templates Found`,
      `No templates where found associated with the email ${email}. 
    Check in GMAIL that you have saved a template / canned message.
    See: https://support.google.com/a/users/answer/9308990?hl=en
    `,
      ui.ButtonSet.OK
    );
    return;
  }

  const subjectsToPickFromDrafts = draftsFilteredByEmail.map(({ getMessage }) => getMessage().getSubject());

  const subject = getSingleUserPropValue('subject');

  const subjectsString = `You can copy from one of these message templates from your account:
  ${subjectsToPickFromDrafts.map((subject) => `\n ${subject}`)}`;

  const message = subject
    ? `Current Template / Canned message is set to: ${subject}.
    You can change this canned message. 
    
    ${subjectsString}
    `
    : `No Template/Canned Message Set. Add An Message.
       
    ${subjectsString}
    `;
  const promptResponse = ui.prompt(`Add / Edit Canned/Template Message`, message, ui.ButtonSet.OK_CANCEL);
  const response = promptResponse.getSelectedButton();
  if (response === ui.Button.OK) {
    const input = promptResponse.getResponseText();
    setUserProps({ subject: input });
    setDraftTemplateAutoResponder({ email, subject: input });
    ui.alert(
      `Canned message Changed`,
      `Your template message is now set to ${getSingleUserPropValue('subject')}`,
      ui.ButtonSet.OK
    );
  }
}

export function setNameToSendInEmail() {
  const ui = SpreadsheetApp.getUi();
  const nameForEmail = getSingleUserPropValue('nameForEmail');
  const message = nameForEmail
    ? `Current name in email is set to ${nameForEmail}. You can change this name.`
    : `No name is sent for email. Add a name to appear in the email.`;
  const promptResponse = ui.prompt(`Add / Edit Name To Appear In Email`, message, ui.ButtonSet.OK_CANCEL);
  const response = promptResponse.getSelectedButton();
  if (response === ui.Button.OK) {
    const input = promptResponse.getResponseText();
    setUserProps({ nameForEmail: input });
    ui.alert(`Name Changed`, `Your name is now set to ${getSingleUserPropValue('nameForEmail')}`, ui.ButtonSet.OK);
  }
}

export function sendSelectedEmailsInPendingEmailsSheet() {}

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
    setEmail();
    setCannedMessageName();
    setNameToSendInEmail();
    ui.alert(`Email Sync`, `The script is going to sync your emails`, ui.ButtonSet.OK);
    runScript();
  }
}
