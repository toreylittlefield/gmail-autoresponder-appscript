import { extractGMAILDataForFollowUpSearch, extractGMAILDataForNewMessagesReceivedSearch } from './email/email';
import { getUserProps } from './properties-service/properties-service';
import {
  activeSheet,
  activeSpreadsheet,
  findColumnNumbersOrLettersByHeaderNames,
  formatRowHeight,
  initSpreadsheet,
  sendDraftsIfAutoResponseUserOptionIsOn,
  sendOrMoveManuallyOrDeleteDraftsInPendingSheet,
  writeDomainsListToDoNotRespondSheet,
  writeEmailsToPendingSheet,
  writeLinkInCellsFromSheetComparison,
} from './sheets/sheets';
import { createTriggerForEmailsSync } from './trigger/trigger';
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
  uiButtonManuallyCreateDraftEmailsForSelectedRowsInAutoReceivedSheet,
  userConfigurationModal,
} from './ui/ui';
import { hasAllRequiredUserProps, initialGlobalMap } from './utils/utils';
import {
  allSheets,
  RECEIVED_MESSAGES_ARCHIVE_LABEL_NAME,
  AUTOMATED_RECEIVED_SHEET_NAME,
  PENDING_EMAILS_TO_SEND_SHEET_NAME,
  SENT_MESSAGES_ARCHIVE_LABEL_NAME,
  SENT_MESSAGES_LABEL_NAME,
} from './variables/publicvariables';

// (?:for\W)(.*)(?= at)(?: at\W)(.*) match linkedin email "you applied at..."

export function uiGetEmailsFromGmail(e?: GoogleAppsScript.Events.TimeDriven) {
  const hasReqProps = hasAllRequiredUserProps();
  if (!hasReqProps) return;
  if (createTriggerForEmailsSync() === true) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      `"Email Sync" will now run every 1 Hour in the background as a time based event trigger. There's nothing you need to do.`
    );
  }
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
    initialGlobalMap('autoReceivedSheetEmailThreadIdsMap');
    initialGlobalMap('alwaysAllowMap');
    initialGlobalMap('doNotSendMailAutoMap');
    initialGlobalMap('pendingEmailsToSendMap');
    initialGlobalMap('sentEmailsByDomainMap');
    initialGlobalMap('sentEmailsBySentMessageIdMap');

    extractGMAILDataForNewMessagesReceivedSearch(email, labelToSearch, RECEIVED_MESSAGES_ARCHIVE_LABEL_NAME, e);

    // order initializing this map matters as writing to the received automation sheet should occur first
    initialGlobalMap('followUpSheetMessageIdMap');
    extractGMAILDataForFollowUpSearch(email, SENT_MESSAGES_LABEL_NAME, SENT_MESSAGES_ARCHIVE_LABEL_NAME);

    writeDomainsListToDoNotRespondSheet;
    writeEmailsToPendingSheet();

    const autoColumns = findColumnNumbersOrLettersByHeaderNames({
      sheetName: AUTOMATED_RECEIVED_SHEET_NAME,
      headerName: ['Email Thread Id'],
    });
    const pendingColumns = findColumnNumbersOrLettersByHeaderNames({
      sheetName: PENDING_EMAILS_TO_SEND_SHEET_NAME,
      headerName: ['Email Thread Id'],
    });

    if (autoColumns['Email Thread Id'] && pendingColumns['Email Thread Id']) {
      writeLinkInCellsFromSheetComparison(
        {
          colNumToWriteTo: pendingColumns['Email Thread Id'].colNumber,
          sheetToWriteToName: PENDING_EMAILS_TO_SEND_SHEET_NAME,
        },
        {
          colNumToLinkFrom: autoColumns['Email Thread Id'].colNumber,
          sheetToLinkFromName: AUTOMATED_RECEIVED_SHEET_NAME,
        }
      );
    }

    allSheets.forEach((sheet) => {
      formatRowHeight(sheet);
    });
  } catch (error) {
    console.error(error as any);
  }
}

/**
 * Runs The Main Script
 * @customFunction
 */
(global as any).getEmailsFromGmail = getEmailsFromGmail;
(global as any).extractGMAILDataForFollowUpSearch = extractGMAILDataForFollowUpSearch;

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
(global as any).uiButtonManuallyCreateDraftEmailsForSelectedRowsInAutoReceivedSheet =
  uiButtonManuallyCreateDraftEmailsForSelectedRowsInAutoReceivedSheet;

(global as any).archiveSelectRowsInAutoReceivedSheet = archiveSelectRowsInAutoReceivedSheet;
(global as any).deleteSelectRowsInAutoReceivedSheet = deleteSelectRowsInAutoReceivedSheet;
(global as any).removeLabelSelectRowsInAutoReceivedSheet = removeLabelSelectRowsInAutoReceivedSheet;
(global as any).sendDraftsIfAutoResponseUserOptionIsOn = sendDraftsIfAutoResponseUserOptionIsOn;

// delete
