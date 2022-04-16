import { getEmailsFromGmail } from '../index';
import { sendDraftsIfAutoResponseUserOptionIsOn } from '../sheets/sheets';

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

/**
 * Creates time-driven trigger for {@link getEmailsFromGmail} that runs every 1 hour if it does not already exist
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 * @returns false if trigger already exists for {@link getEmailsFromGmail}
 * @returns true if trigger has just been newly created
 */
export function createTriggerForAutoResponsingToEmails() {
  if (hasTriggerByName(sendDraftsIfAutoResponseUserOptionIsOn.name)) return false;
  ScriptApp.newTrigger(sendDraftsIfAutoResponseUserOptionIsOn.name).timeBased().everyHours(1).create();
  return true;
}

/**
 * Deletes all triggers matching the function name
 */
export function deleteAllTriggersWithMatchingFunctionName(functionName: Function['name']) {
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (trigger.getHandlerFunction() === functionName) ScriptApp.deleteTrigger(trigger);
  });
}

/**
 * Deletes all triggers
 */
export function deleteAllExistingProjectTriggers() {
  ScriptApp.getProjectTriggers().forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}
