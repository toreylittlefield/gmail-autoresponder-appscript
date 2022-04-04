import { sendTemplateEmail, setDraftTemplateAutoResponder } from './email/email';
import { initSpreadsheet } from './sheets/sheets';
import { LABEL_NAME } from './variables/publicvariables';
import { EMAIL_ACCOUNT } from './variables/privatevariables';

// (?:for\W)(.*)(?= at)(?: at\W)(.*) match linkedin email "you applied at..."

function runScript() {
  // PropertiesService.getUserProperties().deleteAllProperties();
  setDraftTemplateAutoResponder();
  const activeSheet = initSpreadsheet();
  if (!activeSheet) return;
  AutoResponder(activeSheet);
  if (false) {
    sendTemplateEmail('toreylittlefield@gmail.com', 'Responding To Your Message For: Software Engineer');
  }
}

//@ts-expect-error
function AutoResponder(activeSheet: GoogleAppsScript.Spreadsheet.Sheet, event?: GoogleAppsScript.Events.TimeDriven) {
  /* The idea is to search for all emails that don't have this label
  and respond to them with a pre-recorded message like any other email client.
  TODO: constrain it to unread emails sent since a set date */

  // Search for subject:
  // const subject       = "this is a test";

  // Exclude this label:
  // (And creates it if it doesn't exist)
  // return;
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
  const threads = GmailApp.search('label:' + 'recruiters-linkedin-recruiters');
  // + ' to:(' + EMAIL_ACCOUNT + ')');

  const ourEmailDomain = '@' + EMAIL_ACCOUNT.split('@')[1].toString();
  const hasResponseFromRegex = new RegExp(`${ourEmailDomain}|canned\.response@${ourEmailDomain}`);

  const dataRange = activeSheet.getDataRange();
  const dataValues = dataRange.getValues().slice(1);

  let salaries: string[] = [];
  threads.forEach((thread, threadIndex) => {
    if (threadIndex > 2) return;
    // Respond to email
    const [firstMsg, ...restMsgs] = thread.getMessages();

    const firstMsgId = firstMsg.getId();
    const indexOfRowFirstMsgId = dataValues.findIndex((row) => {
      const [msgId] = row;
      return msgId === firstMsgId;
    });

    restMsgs.forEach((msg) => console.log(msg.getFrom().match(hasResponseFromRegex) || 'false'));
    const autoResponseMsg = restMsgs.filter((msg) => msg.getFrom().match(hasResponseFromRegex));

    if (indexOfRowFirstMsgId !== -1 && autoResponseMsg.length > 0) {
      const [
        emailId,
        emailDate,
        fromEmail,
        replyToEmail,
        bodyEmails,
        bodyMsg,
        salary,
        emailPermalink,
        hasEmailResponse,
      ] = dataValues[indexOfRowFirstMsgId];
      console.log(
        emailId,
        emailDate,
        fromEmail,
        replyToEmail,
        bodyEmails,
        bodyMsg,
        salary,
        emailPermalink,
        hasEmailResponse
      );
      activeSheet
        .getRange(indexOfRowFirstMsgId + 2, dataValues[indexOfRowFirstMsgId].length)
        .setValue(autoResponseMsg.map((row) => row.getTo()).toString());
      return;
    }

    const from = firstMsg.getFrom();

    const isNoReplyLinkedIn = from.match(/noreply@linkedin\.com/gi);
    if (isNoReplyLinkedIn) return;

    const body = firstMsg.getPlainBody();
    const replyTo = firstMsg.getReplyTo();
    const regexEmail = /([a-zA-Z0-9+._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi;
    //@ts-ignore
    const emailFrom = from.match(regexEmail);
    //@ts-ignore

    const emailsBody = body.match(regexEmail);
    //@ts-ignore

    const emailReplyTo = replyTo.match(regexEmail);
    const salaryAmount = body.match(
      /\$[1-2][0-9][0-9][-\s][1-2][0-9][0-9]|[1-2][0-9][0-9][-\s]\[1-2][0-9][0-9]|[1-2][0-9][0-9]k/gi
    );

    activeSheet.appendRow([
      firstMsg.getId(),
      firstMsg.getDate(),
      emailFrom ? [...new Set(emailFrom)].toString() : undefined,
      emailReplyTo ? [...new Set(emailReplyTo)].toString() : undefined,
      emailsBody ? [...new Set(emailsBody)].toString() : undefined,
      body,
      salaryAmount ? salaryAmount.toString() : undefined,
      thread.getPermalink(),
      autoResponseMsg ? true : false,
    ]);
    const lastRow = activeSheet.getLastRow();
    // messaging-digest-noreply@linkedin.com
    // inmail-hit-reply@linkedin.com
    activeSheet.setRowHeight(lastRow - 1, 21);
    salaryAmount && salaryAmount.length > 0 ? salaries.push(...Array.from(salaryAmount as string[])) : null;

    // thread.reply('', { htmlBody: response_body, from: EMAIL_ACCOUNT });

    // Add label to email for exclusion
    // thread.addLabel(label);
  });
  console.log({ salaries });
}

/**
 * Runs The Autoresponder script
 *
 *
 * @customFunction
 */
(global as any).runScript = runScript;
