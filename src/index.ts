import { extractDataFromEmailSearch, sendTemplateEmail } from './email/email';
import {
  activeSheet,
  activeSpreadsheet,
  formatRowHeight,
  initSpreadsheet,
  writeDomainsListToDoNotRespondSheet,
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

export function getEmailsFromGmail(e?: GoogleAppsScript.Events.TimeDriven) {
  try {
    const hasReqProps = hasAllRequiredUserProps();
    if (!hasReqProps) return;
    // PropertiesService.getUserProperties().deleteAllProperties();
    initSpreadsheet();
    if (!activeSpreadsheet) throw Error('No Active Spreadsheet');
    if (!activeSheet) throw Error('No Active Sheet');

    initialGlobalMap('doNotTrackMap');
    initialGlobalMap('emailmessagesIdMap');

    extractDataFromEmailSearch(e);

    formatRowHeight();

    writeDomainsListToDoNotRespondSheet();

    /** send emails and replies */
    //addSentEmailsToDoNotReplyMap
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
