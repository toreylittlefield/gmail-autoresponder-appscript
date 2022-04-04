import { FUNCTION_NAME } from '../variables';

/**
 * Creates two time-driven triggers.
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
export function createTimeDrivenTriggers() {
  // Trigger every 6 hours.
  ScriptApp.newTrigger(FUNCTION_NAME).timeBased().everyMinutes(1).create();
  // Trigger every Monday at 09:00.
  // ScriptApp.newTrigger('AutoResponder').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
}
