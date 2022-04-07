import { setDraftTemplateAutoResponder } from '../email/email';
import { runScript } from '../index';
import { getSingleUserPropValue, setUserProps } from '../properties-service/properties-service';
import { checkExistsOrCreateSpreadsheet, WarningResetSheetsAndSpreadsheet } from '../sheets/sheets';
import { LABEL_NAME } from '../variables/publicvariables';

const menuName = `Autoresponder Email Settings Menu`;

function createMenuAfterStart(ui: GoogleAppsScript.Base.Ui, menu: GoogleAppsScript.Base.Menu) {
  const optionsMenu = ui.createMenu('Options');
  optionsMenu.addItem(`Toggle Automatic Email Sending`, 'toggleAutoResponseOnOff');
  optionsMenu.addItem(`Add / Edit Email`, 'setEmail');
  optionsMenu.addItem(`Add / Edit Canned Message Name`, 'setCannedMessageName');
  optionsMenu.addItem(`Add / Edit Gmail Label To Search | Watch`, 'setLabelToSearchInGmail');
  optionsMenu.addItem(`Add / Edit Name To Send In Email`, 'setNameToSendInEmail');
  optionsMenu.addSeparator();

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
    await checkExistsOrCreateSpreadsheet().catch((err) => {
      console.error(err);
      WarningResetSheetsAndSpreadsheet();
    });
    setEmail();
    setCannedMessageName();
    setLabelToSearchInGmail();
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
      : `Main email: ${mainEmail}

        You can also add aliases to your gmail account see:
        https://support.google.com/mail/answer/22370?hl=en#zippy=%2Cfilter-using-your-gmail-alias%2Csend-from-a-work-or-school-group-alias
      `;
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

export function setLabelToSearchInGmail() {
  const ui = SpreadsheetApp.getUi();

  const currentEmail = getSingleUserPropValue('email');

  if (!currentEmail) {
    ui.alert(`Please Set Email`, `You Need To Set An Email Before Setting The Message`, ui.ButtonSet.OK);
    return;
  }

  const currentLabel = getSingleUserPropValue('labelToSearch');
  const userLabels = GmailApp.getUserLabels();
  const userLabelsMessage =
    userLabels.length > 0
      ? `Gmail Labels : ${userLabels.map((item) => `\n ${item.getName()}`)}`
      : `No Labels Found

      We'll create a Gmail LABEL and FILTER for you automatically or you can cancel and create a label and filter on your own.

      To Create Labels In Gmail See: https://support.google.com/mail/answer/118708?hl=en&co=GENIE.Platform%3DAndroid`;
  const message = currentLabel
    ? `Current Label is set to: ${currentLabel}
      You Can Change This Label.

      You can also change to one of the following in your gmail account:
      ${userLabelsMessage}
      `
    : `No Label Is Set. Add a Label.

      Use one of the following in your gmail account:
      ${userLabelsMessage}`;
  if (userLabels.length === 0) {
    const response = ui.alert(`No Gmail Labels Found`, userLabelsMessage, ui.ButtonSet.OK);
    if (response === ui.Button.OK) {
      createFilterAndLabel(currentEmail, ui);
      return;
    } else return;
  }
  const promptResponse = ui.prompt(`Add / Edit Gmail Label To Use`, message, ui.ButtonSet.OK_CANCEL);
  const response = promptResponse.getSelectedButton();
  if (response === ui.Button.OK) {
    const input = promptResponse.getResponseText();
    setUserProps({ labelToSearch: input });
    ui.alert(
      `Label Changed`,
      `Your label to search Gmail Messages by is now set to: ${getSingleUserPropValue('labelToSearch')}`,
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
    setLabelToSearchInGmail();
    setNameToSendInEmail();
    ui.alert(`Email Sync`, `The script is going to sync your emails`, ui.ButtonSet.OK);
    runScript();
  }
}

function createFilterAndLabel(currentEmail: string, ui: GoogleAppsScript.Base.Ui) {
  const me = Session.getActiveUser().getEmail();

  const gmailUser = Gmail.Users as GoogleAppsScript.Gmail.Collection.UsersCollection;

  const labelsCollection = gmailUser.Labels as GoogleAppsScript.Gmail.Collection.Users.LabelsCollection;
  const newLabel = labelsCollection.create(
    {
      color: {
        backgroundColor: '#42d692',
        textColor: '#ffffff',
      },
      name: LABEL_NAME,
      labelListVisibility: 'labelShow',
      messageListVisibility: 'show',
      type: 'user',
    },
    me
  ) as GoogleAppsScript.Gmail.Schema.Label;

  const userSettings = gmailUser.Settings as GoogleAppsScript.Gmail.Collection.Users.SettingsCollection;
  const filters = userSettings.Filters as GoogleAppsScript.Gmail.Collection.Users.Settings.FiltersCollection;
  const newFilter = filters.create(
    {
      action: {
        addLabelIds: [newLabel.id as string],
      },
      criteria: {
        to: currentEmail,
      },
    },
    me
  );

  const resFilter = filters.get(me, newFilter.id as string);
  const resLabel = labelsCollection.get(me, newLabel.id as string);
  if (resFilter.id && resLabel.id) {
    ui.alert(
      `Created Filter ${resFilter.id} in GMAIL with messages to email ${
        (resFilter.criteria as GoogleAppsScript.Gmail.Schema.FilterCriteria).to
      } with automatically have the label: ${resLabel.name} applied to them`
    );
    setUserProps({ labelToSearch: resLabel.name, labelId: resLabel.id, filterId: resFilter.id });
  }
}
