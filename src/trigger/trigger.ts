import { getEmailsFromGmail } from '../index';

/**
 *
 * Check if trigger already exists for a function
 */
function hasTriggerByName(functionName: Function['name']) {
  return ScriptApp.getProjectTriggers().some((trigger) => trigger.getHandlerFunction() === functionName);
}

/**
 * Creates time-driven trigger for {@link getEmailsFromGmail} that runs every 1 hour
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
export function createTriggerForEmailsSync() {
  if (!hasTriggerByName(getEmailsFromGmail.name))
    ScriptApp.newTrigger(getEmailsFromGmail.name).timeBased().everyHours(1);
}
