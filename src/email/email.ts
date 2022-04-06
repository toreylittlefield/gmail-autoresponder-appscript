import { getProps, setInitialEmailProps, setUserProps } from '../properties-service/properties-service';
import { addToRepliesArray, doNotSendMailAutoMap, doNotTrackMap, getSheetByName } from '../sheets/sheets';
import { calcAverage, getDomainFromEmailAddress, getEmailFromString, regexEmail, regexSalary } from '../utils/utils';
import { EMAIL_ACCOUNT, NAME_TO_SEND_IN_EMAIL } from '../variables/privatevariables';
import { LABEL_NAME } from '../variables/publicvariables';

type EmailDataToSend = 'replyToEmail' | { emailSubject: string; personName: string };
type EmailReplySendArray = [emailAddress: string, replyOrNew: EmailDataToSend][];

const sendToEmailsMap = new Map<string, EmailDataToSend>();
export const emailmessagesIdMap = new Map<string, number>();

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

function checkAndAddToEmailMap(emails: EmailReplySendArray) {
  emails.forEach(([email, data]) => {
    const domain = getDomainFromEmailAddress(email);

    if (!doNotSendMailAutoMap.has(email) && !doNotSendMailAutoMap.has(domain) && !sendToEmailsMap.has(email)) {
      sendToEmailsMap.set(email, data);
    }
  });
}

function buildEmailsObject(
  replyToEmail: string,
  bodyEmails: string[],
  emailSubject: string,
  personFrom: string
): EmailReplySendArray {
  const emailObject: { [key: string]: EmailDataToSend } = {};
  bodyEmails.forEach((email) => {
    if (email !== replyToEmail) {
      emailObject.email = { personName: personFrom, emailSubject };
    }
  });
  emailObject[replyToEmail] = 'replyToEmail';
  return Object.entries(emailObject);
}

function addSentEmailsToDoNotReplyMap(sentEmails: string[]) {
  sentEmails.forEach((email) => {
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
}

export function getToEmailArray(emailMessages: GoogleAppsScript.Gmail.GmailMessage[]) {
  return emailMessages.map((row) => row.getTo()).toString();
}

function getAutoResponseMsgsFromThread(restMsgs: GoogleAppsScript.Gmail.GmailMessage[]) {
  const ourEmailDomain = '@' + EMAIL_ACCOUNT.split('@')[1].toString();

  const hasResponseFromRegex = new RegExp(`${ourEmailDomain}|canned\.response@${ourEmailDomain}`);

  return restMsgs.filter((msg) => msg.getFrom().match(hasResponseFromRegex));
}

function updateRepliesColumnIfMessageHasReplies(firstMsgId: string, restMsgs: GoogleAppsScript.Gmail.GmailMessage[]) {
  const messageAlreadyExists = emailmessagesIdMap.has(firstMsgId);

  const autoResponseMsg = getAutoResponseMsgsFromThread(restMsgs);

  if (autoResponseMsg.length > 0 && messageAlreadyExists) {
    const rowNumber = emailmessagesIdMap.get(firstMsgId) as number;
    addToRepliesArray(rowNumber, autoResponseMsg);
  }

  return autoResponseMsg;
}

/**
 * 1. get unread mail sent to email_account
 * 2. search for unread mail that does not have the label for our email_account and email was sent less time < 30 minutes
 * 3. for each found message from 2., search for that "@domain.xyz" to check if we've already been messaged by that domain
 * 4. if the search results from 3. is of length 1, it is the first message and therefore we send the autoresponse
 * 5. otherwise we don't send an autoresponse to avoid sending the autoresponse to the same domain more than once
 */

function isDomainEmailInDoNotTrackSheet(fromEmail: string) {
  const domain = getDomainFromEmailAddress(fromEmail);

  if (doNotTrackMap.has(domain) || doNotSendMailAutoMap.has(fromEmail)) return true;
  /**TODO: Can be optimized in future if slow perf */
  if (Array.from(doNotTrackMap.keys()).filter((domainOrEmailKey) => fromEmail.match(domainOrEmailKey)).length > 0)
    return true;
  return false;
}

export function extractDataFromEmailSearch(event?: GoogleAppsScript.Events.TimeDriven) {
  try {
    const autoResultsListSheet = getSheetByName('Automated Results List');
    if (!autoResultsListSheet) throw Error('Cannot find Automated Results List Sheet');
    console.log({ event });
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

    // Send our response email and label it responded to
    // const threads = GmailApp.search(
    //   "-subject:'re:' -is:chats -is:draft has:nouserlabels -label:" + LABEL_NAME + ' to:(' + EMAIL_ACCOUNT + ')'
    // );
    const threads = GmailApp.search('label:' + 'recruiters-linkedin-recruiters');
    // + ' to:(' + EMAIL_ACCOUNT + ')');

    let salaries: number[] = [];
    threads.forEach((thread, threadIndex) => {
      if (threadIndex > 10) return;
      // Respond to email
      const [firstMsg, ...restMsgs] = thread.getMessages();

      const firstMsgId = firstMsg.getId();

      const autoResponseMsg = updateRepliesColumnIfMessageHasReplies(firstMsgId, restMsgs);

      const from = firstMsg.getFrom();
      const emailThreadId = thread.getId();
      const emailSubject = thread.getFirstMessageSubject();

      const body = firstMsg.getPlainBody();
      const replyTo = firstMsg.getReplyTo();

      /** Use as a backup in case other split methods fail */
      // const emailFrom = [...new Set(from.match(regexEmail))];
      // const emailReplyTo = [...new Set(replyTo.match(regexEmail))];

      const emailFrom = getEmailFromString(from);
      const personFrom = from.split('<', 1)[0].trim();
      const emailReplyTo = getEmailFromString(replyTo);

      const bodyEmails = [...new Set(body.match(regexEmail))];
      const salaryAmount = body.match(regexSalary);

      if (isDomainEmailInDoNotTrackSheet(emailFrom)) return;

      // const isNoReplyLinkedIn = from.match(/noreply@linkedin\.com/gi);
      // if (isNoReplyLinkedIn) return;

      checkAndAddToEmailMap(buildEmailsObject(replyTo, bodyEmails, emailSubject, personFrom));

      /**
       * [
  'Email Thread Id',
  'Email Id',
  'Date',
  'From',
  'ReplyTo',
  'Person / Company Name',
  'Subject',
  'Body Emails',
  'Body',
  'Salary',
  'Email Permalink',
  'Has Email Response',
];
       */
      autoResultsListSheet.appendRow([
        emailThreadId,

        firstMsg.getId(),
        firstMsg.getDate(),
        emailFrom,
        emailReplyTo.length > 0 ? emailReplyTo.toString() : undefined,
        emailSubject,
        bodyEmails.length > 0 ? bodyEmails.toString() : undefined,
        body,
        salaryAmount ? salaryAmount.toString() : undefined,
        thread.getPermalink(),
        autoResponseMsg.length > 0 ? getToEmailArray(autoResponseMsg) : false,
      ]);

      // messaging-digest-noreply@linkedin.com
      // inmail-hit-reply@linkedin.com
      salaryAmount && salaryAmount.length > 0 && salaries.push(calcAverage(salaryAmount));

      // thread.reply('', { htmlBody: response_body, from: EMAIL_ACCOUNT });

      // Add label to email for exclusion
      // thread.addLabel(label);
    });
    addSentEmailsToDoNotReplyMap(Array.from(sendToEmailsMap.keys()));
    console.log({ salaries: calcAverage(salaries) });
  } catch (error) {
    console.error(error as any);
  }
}
