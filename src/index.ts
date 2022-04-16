import { extractDataFromEmailSearch } from './email/email';
import { getUserProps } from './properties-service/properties-service';
import {
  activeSheet,
  activeSpreadsheet,
  formatRowHeight,
  initSpreadsheet,
  sendOrMoveManuallyOrDeleteDraftsInPendingSheet,
  writeDomainsListToDoNotRespondSheet,
  writeEmailsToPendingSheet,
  writeLinkInCellsFromSheetComparison,
} from './sheets/sheets';
import {
  archiveSelectRowsInAutoReceivedSheet,
  deleteSelectedEmailsInPendingEmailsSheet,
  deleteSelectRowsInAutoReceivedSheet,
  getUserPropertiesForPageModal,
  initializeSpreadsheets,
  menuItemResetEntireSheet,
  moveManuallySelectedEmailsInPendingEmailsSheet,
  onOpen,
  processFormEventsFromPage,
  removeLabelSelectRowsInAutoReceivedSheet,
  sendSelectedEmailsInPendingEmailsSheet,
  toggleAutoResponseOnOff,
  userConfigurationModal,
} from './ui/ui';
import { hasAllRequiredUserProps, initialGlobalMap } from './utils/utils';
import { AUTOMATED_RECEIVED_SHEET_NAME } from './variables/publicvariables';

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
    initialGlobalMap('emailThreadIdsMap');
    initialGlobalMap('alwaysAllowMap');
    initialGlobalMap('doNotSendMailAutoMap');
    initialGlobalMap('pendingEmailsToSendMap');

    extractDataFromEmailSearch(email, labelToSearch, e);

    writeDomainsListToDoNotRespondSheet;
    writeEmailsToPendingSheet();
    writeLinkInCellsFromSheetComparison(
      { colNumToWriteTo: 2, sheetToWriteToName: 'Pending Emails To Send' },
      { colNumToLinkFrom: 1, sheetToLinkFromName: `${AUTOMATED_RECEIVED_SHEET_NAME}` }
    );

    formatRowHeight('Always Autorespond List');
    formatRowHeight('Pending Emails To Send');
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
(global as any).moveManuallySelectedEmailsInPendingEmailsSheet = moveManuallySelectedEmailsInPendingEmailsSheet;
(global as any).deleteSelectedEmailsInPendingEmailsSheet = deleteSelectedEmailsInPendingEmailsSheet;
(global as any).userConfigurationModal = userConfigurationModal;
(global as any).getUserPropertiesForPageModal = getUserPropertiesForPageModal;
(global as any).processFormEventsFromPage = processFormEventsFromPage;
(global as any).sendOrMoveManuallyOrDeleteDraftsInPendingSheet = sendOrMoveManuallyOrDeleteDraftsInPendingSheet;
(global as any).archiveSelectRowsInAutoReceivedSheet = archiveSelectRowsInAutoReceivedSheet;
(global as any).deleteSelectRowsInAutoReceivedSheet = deleteSelectRowsInAutoReceivedSheet;
(global as any).removeLabelSelectRowsInAutoReceivedSheet = removeLabelSelectRowsInAutoReceivedSheet;
