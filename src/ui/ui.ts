import { getEmailsFromGmail } from '../index';
import {
  getSingleUserPropValue,
  getUserProps,
  setUserProps,
  UserRecords,
} from '../properties-service/properties-service';
import { checkExistsOrCreateSpreadsheet, WarningResetSheetsAndSpreadsheet } from '../sheets/sheets';
import { LABEL_NAME } from '../variables/publicvariables';

const menuName = `Autoresponder Email Settings Menu`;

function createMenuAfterStart(ui: GoogleAppsScript.Base.Ui, menu: GoogleAppsScript.Base.Menu) {
  const optionsMenu = ui.createMenu('Options');
  optionsMenu.addItem(`Toggle Automatic Email Sending`, 'toggleAutoResponseOnOff');
  optionsMenu.addSeparator();

  optionsMenu.addItem('Reset Entire Sheet', 'menuItemResetEntireSheet');

  menu.addItem('User Configuration', 'userConfigurationModal');
  menu.addItem(`Sync Emails`, 'getEmailsFromGmail');
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
    ui.alert(`Email Sync`, `The script is going to sync your emails`, ui.ButtonSet.OK);
    getEmailsFromGmail();
    SpreadsheetApp.getActiveSpreadsheet().removeMenu(menuName);
    const menu = ui.createMenu(menuName);
    createMenuAfterStart(ui, menu);
  }
}

function getUserLabels() {
  const currentLabel = getSingleUserPropValue('labelToSearch');
  const labels = GmailApp.getUserLabels();

  return { currentLabel, userLabels: labels.length > 0 ? labels.map(({ getName }) => getName()) : [] };
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

    ui.alert(`Email Sync`, `The script is going to sync your emails`, ui.ButtonSet.OK);
    getEmailsFromGmail();
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

function getUserEmails() {
  const emailAliases = GmailApp.getAliases();
  const mainEmail = Session.getActiveUser().getEmail();
  const currentEmailUserStore = getSingleUserPropValue('email') || 'none set';
  return { emailAliases, mainEmail, currentEmailUserStore };
}

function getUserNameForEmail() {
  const nameForEmail = getSingleUserPropValue('nameForEmail');
  return { nameForEmail };
}

type DraftsToPick = { subject: string; draftId: string; subjectBody: string };

function getUserCannedMessage(): { draftsList: DraftsToPick[]; subject: string } {
  const email = getSingleUserPropValue('email');
  if (!email) {
    return { draftsList: [], subject: '' };
  }
  const drafts = GmailApp.getDrafts();
  const draftsFilteredByEmail = drafts.filter((draft) => {
    const { getTo, getSubject } = draft.getMessage();
    return getTo() === '' && getSubject();
  });
  let draftsList = draftsFilteredByEmail.map(({ getId, getMessage }) => ({
    draftId: getId(),
    subject: getMessage().getSubject().trim(),
    subjectBody: getMessage().getPlainBody(),
  }));

  const subject = getSingleUserPropValue('subject');

  return { draftsList, subject: subject || '' };
}

export function processFormEventsFromPage(formObject: Partial<Record<UserRecords, string>>) {
  if (formObject.email) {
    setUserProps({ email: formObject.email });
    return getUserProps(['email']);
  }
  if (formObject.labelToSearch) {
    if (formObject.labelToSearch === 'create-label') {
      createFilterAndLabel(
        getSingleUserPropValue('email') || Session.getActiveUser().getEmail(),
        SpreadsheetApp.getUi()
      );
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
