import { getEmailsFromGmail } from '../index';

/**
 *
 * Check if trigger already exists for a function
 */
function hasTriggerByName(nameOfFunction: Function['name']) {
  return ScriptApp.getProjectTriggers().some((trigger) => trigger.getHandlerFunction() === nameOfFunction);
}

/**
 * Creates time-driven trigger for {@link getEmailsFromGmail} that runs every 1 hour if it does not already exist
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
export function createTriggerForEmailsSync() {
  if (hasTriggerByName(getEmailsFromGmail.name)) return false;
  ScriptApp.newTrigger(getEmailsFromGmail.name).timeBased().everyHours(1).create();
  return true;
}

export function deleteAllExistingProjectTriggers() {
  ScriptApp.getProjectTriggers().forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}
