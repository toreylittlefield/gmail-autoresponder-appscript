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
  setCannedMessageName,
  setEmail,
  setLabelToSearchInGmail,
  setNameToSendInEmail,
  userOptionsModal,
  toggleAutoResponseOnOff,
} from './ui/ui';
import { initialGlobalMap } from './utils/utils';

// (?:for\W)(.*)(?= at)(?: at\W)(.*) match linkedin email "you applied at..."

/**
 * 1. search for emails with the label that have been received in the last 90 days
 *
 */

export function runScript(e?: GoogleAppsScript.Events.TimeDriven) {
  try {
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
(global as any).runScript = runScript;

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
(global as any).setEmail = setEmail;
(global as any).setCannedMessageName = setCannedMessageName;
(global as any).setNameToSendInEmail = setNameToSendInEmail;
(global as any).setLabelToSearchInGmail = setLabelToSearchInGmail;
(global as any).userOptionsModal = userOptionsModal;
(global as any).getUserPropertiesForPageModal = getUserPropertiesForPageModal;
(global as any).processFormEventsFromPage = processFormEventsFromPage;
