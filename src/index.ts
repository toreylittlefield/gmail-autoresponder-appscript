import { extractDataFromEmailSearch, sendTemplateEmail, setDraftTemplateAutoResponder } from './email/email';
import {
  activeSheet,
  activeSpreadsheet,
  formatRowHeight,
  initSpreadsheet,
  writeDomainsListToDoNotRespondSheet,
} from './sheets/sheets';
import { initialGlobalMap } from './utils/utils';

// (?:for\W)(.*)(?= at)(?: at\W)(.*) match linkedin email "you applied at..."

function runScript(e: GoogleAppsScript.Events.TimeDriven) {
  try {
    // PropertiesService.getUserProperties().deleteAllProperties();
    initSpreadsheet();
    if (!activeSpreadsheet) throw Error('No Active Spreadsheet');
    if (!activeSheet) throw Error('No Active Sheet');

    setDraftTemplateAutoResponder();

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
 * Runs The Autoresponder script
 *
 *
 * @customFunction
 */
(global as any).runScript = runScript;
