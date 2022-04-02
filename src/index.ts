import { CANNED_MSG_NAME, EMAIL_ACCOUNT, FUNCTION_NAME, LABEL_NAME } from './variables';

// (?:for\W)(.*)(?= at)(?: at\W)(.*) match linkedin email "you applied at..."

/**
 * Creates two time-driven triggers.
 * @see https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers
 */
//@ts-ignore
function createTimeDrivenTriggers() {
  // Trigger every 6 hours.
  ScriptApp.newTrigger(FUNCTION_NAME).timeBased().everyMinutes(1).create();
  // Trigger every Monday at 09:00.
  // ScriptApp.newTrigger('AutoResponder').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
}

function setUserProps(props: Record<string, any>) {
  const userProps = PropertiesService.getUserProperties();

  userProps.setProperties(props);
}

function getProps(keys: string[]) {
  const userProps = PropertiesService.getUserProperties();
  const props: Record<string, any> = {};
  keys.forEach((key) => {
    const value = userProps.getProperty(key);
    props[key] = value;
  });
  return props;
}

function setInitalProps() {
  const userProps = PropertiesService.getUserProperties();

  if (!userProps.getProperty('subject') || !userProps.getProperty('email')) {
    setUserProps({ subject: CANNED_MSG_NAME, email: EMAIL_ACCOUNT });
  }
}

function setTemplateMsg({ subject, email }: { subject: string; email: string }) {
  const drafts = GmailApp.getDrafts();

  console.log({ subject, email });
  drafts.forEach((draft) => {
    const { getFrom, getSubject, getId } = draft.getMessage();
    if (subject === getSubject() && getFrom().match(email)) {
      console.log({ messageId: getId(), draftId: draft.getId() });
      setUserProps({ messageId: getId(), draftId: draft.getId() });
    }
  });
}

function getDraft() {
  setInitalProps();
  const props = getProps(['subject', 'email', 'draftId', 'messageId']);
  console.log({ props });
  const { subject, email, draftId, messageId } = props;
  if (!draftId || !messageId) {
    setTemplateMsg({ subject, email });
  }
  const draft = GmailApp.getDraft(draftId);
  return draft;
}

function runScript() {
  getDraft();
  AutoResponder();
}

//@ts-expect-error
function AutoResponder(event?: GoogleAppsScript.Events.TimeDriven) {
  /* The idea is to search for all emails that don't have this label
  and respond to them with a pre-recorded message like any other email client.
  TODO: constrain it to unread emails sent since a set date */

  // Search for subject:
  // const subject       = "this is a test";

  // Exclude this label:
  // (And creates it if it doesn't exist)

  let label = GmailApp.getUserLabelByName(LABEL_NAME);
  // Create label if it doesn't exist
  if (label == null) {
    label = GmailApp.createLabel(LABEL_NAME);
  }

  /**
   * 1. get unread mail sent to email_account
   * 2. search for unread mail that does not have the label for our email_account and email was sent less time < 30 minutes
   * 3. for each found message from 2., search for that "@domain.xyz" to check if we've already been messaged by that domain
   * 4. if the search results from 3. is of length 1, it is the first message and therefore we send the autoresponse
   * 5. otherwise we don't send an autoresponse to avoid sending the autoresponse to the same domain more than once
   */

  // Send our response email and label it responded to
  // const threads = GmailApp.search(
  //   "-subject:'re:' -is:chats -is:draft has:nouserlabels -label:" + LABEL_NAME + ' to:(' + EMAIL_ACCOUNT + ')'
  // );
  const threads = GmailApp.search('label:' + 'recruiters-autoresponder' + ' to:(' + EMAIL_ACCOUNT + ')');

  threads.forEach((thread) => {
    // Response
    // const response_body =
    //   "Hey there,<br><br>\
    //   Unfortunately I'm out of the office until the 32nd of Nebuary.<br><br>\
    // Thanks,<br>\
    // Jason";

    // Respond to email
    thread.getMessages().forEach((msg) => {
      console.log(msg.getFrom());
    });
    // thread.reply('', { htmlBody: response_body, from: EMAIL_ACCOUNT });

    // Add label to email for exclusion
    // thread.addLabel(label);
  });
}

/**
 * Performs a useless calculation
 *
 *
 * @customFunction
 */
(global as any).runScript = runScript;
