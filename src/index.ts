import { extractDataFromEmailSearch, sendTemplateEmail } from './email/email';
import { getUserProps } from './properties-service/properties-service';
import {
  activeSheet,
  activeSpreadsheet,
  formatRowHeight,
  initSpreadsheet,
  writeDomainsListToDoNotRespondSheet,
  writeEmailsToPendingSheet,
} from './sheets/sheets';
import {
  getUserPropertiesForPageModal,
  initializeSpreadsheets,
  menuItemResetEntireSheet,
  onOpen,
  processFormEventsFromPage,
  sendSelectedEmailsInPendingEmailsSheet,
  toggleAutoResponseOnOff,
  userConfigurationModal,
} from './ui/ui';
import { hasAllRequiredUserProps, initialGlobalMap } from './utils/utils';

// (?:for\W)(.*)(?= at)(?: at\W)(.*) match linkedin email "you applied at..."

/**
 * 1. search for emails with the label that have been received in the last 90 days
 *
 */
export function uiGetEmailsFromGmail(e?: GoogleAppsScript.Events.TimeDriven) {
  const hasReqProps = hasAllRequiredUserProps();
  if (!hasReqProps) return;
  getEmailsFromGmail(e);
}

export function getEmailsFromGmail(e?: GoogleAppsScript.Events.TimeDriven) {
  try {
    const userConfiguration = getUserProps(['email', 'nameForEmail', 'labelToSearch', 'subject', 'draftId']);

    const { email, labelToSearch } = userConfiguration;
    if (!email) throw Error('No Email Set In User Configuration');
    if (!labelToSearch) throw Error('No Label Set In User Configuration');

    // PropertiesService.getUserProperties().deleteAllProperties();
    initSpreadsheet();
    if (!activeSpreadsheet) throw Error('No Active Spreadsheet');
    if (!activeSheet) throw Error('No Active Sheet');

    initialGlobalMap('doNotTrackMap');
    initialGlobalMap('emailmessagesIdMap');
    initialGlobalMap('doNotSendMailAutoMap');
    initialGlobalMap('pendingEmailsToSendMap');

    extractDataFromEmailSearch(email, labelToSearch, e);

    writeDomainsListToDoNotRespondSheet();
    writeEmailsToPendingSheet();

    /** send emails and replies */
    //addSentEmailsToDoNotReplyMap
    formatRowHeight('Always Autorespond List');
    formatRowHeight('Pending Emails To Send');
    if (false) {
      sendTemplateEmail('toreylittlefield@gmail.com', 'Responding To Your Message For: Software Engineer');
    }
  } catch (error) {
    console.error(error as any);
  }
}

/**
 * Runs The Main Script
 * @customFunction
 */
(global as any).getEmailsFromGmail = getEmailsFromGmail;

/**
 * Runs The UI Script
 * @customFunction
 */
(global as any).uiGetEmailsFromGmail = uiGetEmailsFromGmail;

/**
 * Renders the ui menu in spreadsheet on open event
 * @customFunction
 */
(global as any).onOpen = onOpen;

/**
 * Menu Options
 */
(global as any).toggleAutoResponseOnOff = toggleAutoResponseOnOff;
(global as any).menuItemResetEntireSheet = menuItemResetEntireSheet;
(global as any).initializeSpreadsheets = initializeSpreadsheets;
(global as any).sendSelectedEmailsInPendingEmailsSheet = sendSelectedEmailsInPendingEmailsSheet;
(global as any).userConfigurationModal = userConfigurationModal;
(global as any).getUserPropertiesForPageModal = getUserPropertiesForPageModal;
(global as any).processFormEventsFromPage = processFormEventsFromPage;
