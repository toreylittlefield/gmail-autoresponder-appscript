import { getProps, setInitialEmailProps, setUserProps } from '../properties-service/properties-service';
import { doNotSendMailAutoMap, getSheetByName, updateRepliesColumn } from '../sheets/sheets';
import { calcAverage, getDomainFromEmailAddress } from '../utils/utils';
import { EMAIL_ACCOUNT, NAME_TO_SEND_IN_EMAIL } from '../variables/privatevariables';
import { AUTOMATED_SHEET_NAME, LABEL_NAME } from '../variables/publicvariables';

export function setTemplateMsg({ subject, email }: { subject: string; email: string }) {
  const drafts = GmailApp.getDrafts();
  drafts.forEach((draft) => {
    const { getFrom, getSubject, getId } = draft.getMessage();
    if (subject === getSubject() && getFrom().match(email)) {
      setUserProps({ messageId: getId(), draftId: draft.getId() });
    }
  });
}

export function setDraftTemplateAutoResponder() {
  setInitialEmailProps();
  const props = getProps(['subject', 'email', 'draftId', 'messageId']);
  let { subject, email, draftId, messageId } = props;
  if (!draftId || !messageId) {
    setTemplateMsg({ subject, email });
  }
}

function getDraftTemplateAutoResponder() {
  const { draftId } = getProps(['draftId']);
  const draft = GmailApp.getDraft(draftId);
  return draft;
}

export function sendTemplateEmail(recipient: string, subject: string, htmlBodyMessage?: string) {
  try {
    const body = htmlBodyMessage || getDraftTemplateAutoResponder().getMessage().getBody();
    if (!body) throw Error('Could not find draft and send Email');
    GmailApp.sendEmail(recipient, subject, '', {
      from: EMAIL_ACCOUNT,
      htmlBody: body,
      name: NAME_TO_SEND_IN_EMAIL,
    });
  } catch (error) {
    console.error(error as any);
  }
}

export function getToEmailArray(emailMessages: GoogleAppsScript.Gmail.GmailMessage[]) {
  return emailMessages.map((row) => row.getTo()).toString();
}

export function AutoResponder(event?: GoogleAppsScript.Events.TimeDriven) {
  try {
    /* The idea is to search for all emails that don't have this label
  and respond to them with a pre-recorded message like any other email client.
  TODO: constrain it to unread emails sent since a set date */
    console.log({ event });
    const autoResultsListSheet = getSheetByName('Automated Results List');
    if (!autoResultsListSheet) throw Error(`Could Not Find the ${AUTOMATED_SHEET_NAME} to process the results`);
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

    const dataRange = autoResultsListSheet.getDataRange();
    const dataValues = dataRange.getValues().slice(1);

    let salaries: number[] = [];
    threads.forEach((thread, threadIndex) => {
      if (threadIndex > 10) return;
      // Respond to email
      const [firstMsg, ...restMsgs] = thread.getMessages();

      const firstMsgId = firstMsg.getId();
      const indexOfRowFirstMsgId = dataValues.findIndex((row) => {
        const [msgId] = row;
        return msgId === firstMsgId;
      });

      restMsgs.forEach((msg) => console.log(msg.getFrom().match(hasResponseFromRegex) || 'false'));
      const autoResponseMsg = restMsgs.filter((msg) => msg.getFrom().match(hasResponseFromRegex));
      console.log({ autoResponseMsg: autoResponseMsg.map((msg) => msg.getFrom()) });
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
        updateRepliesColumn(autoResultsListSheet, indexOfRowFirstMsgId, dataValues, autoResponseMsg);
        // activeSheet
        //   .getRange(indexOfRowFirstMsgId + 2, dataValues[indexOfRowFirstMsgId].length)
        //   .setValue(autoResponseMsg.map((row) => row.getTo()).toString());
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

      const emailsBody = [...new Set(body.match(regexEmail))];
      //@ts-ignore

      const emailReplyTo = replyTo.match(regexEmail);
      const salaryAmount = body.match(
        /\$[1-2][0-9][0-9][-\s][1-2][0-9][0-9]|[1-2][0-9][0-9][-\s]\[1-2][0-9][0-9]|[1-2][0-9][0-9]k/gi
      );

      autoResultsListSheet.appendRow([
        firstMsg.getId(),
        firstMsg.getDate(),
        emailFrom ? [...new Set(emailFrom)].toString() : undefined,
        emailReplyTo ? [...new Set(emailReplyTo)].toString() : undefined,
        emailsBody ? emailsBody.toString() : undefined,
        body,
        salaryAmount ? salaryAmount.toString() : undefined,
        thread.getPermalink(),
        autoResponseMsg.length > 0 ? getToEmailArray(autoResponseMsg) : false,
      ]);

      emailsBody.forEach((email) => {
        const domain = getDomainFromEmailAddress(email);
        const count = doNotSendMailAutoMap.get(domain);
        if (typeof count === 'number') {
          return doNotSendMailAutoMap.set(domain, count + 1);
        }
        if (count == null) {
          doNotSendMailAutoMap.set(domain, 0);
        }
        return;
      });
      // messaging-digest-noreply@linkedin.com
      // inmail-hit-reply@linkedin.com
      salaryAmount && salaryAmount.length > 0 && salaries.push(calcAverage(salaryAmount));

      // thread.reply('', { htmlBody: response_body, from: EMAIL_ACCOUNT });

      // Add label to email for exclusion
      // thread.addLabel(label);
    });
    console.log({ salaries: calcAverage(salaries) });
  } catch (error) {
    console.error(error as any);
  }
}
