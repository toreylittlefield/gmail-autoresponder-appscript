import {
  doNotSendMailAutoMap,
  doNotTrackMap,
  EmailDataToSend,
  emailmessagesIdMap,
  emailsToSendMap,
  pendingEmailsToSendMap,
} from '../global/maps';
import { getSingleUserPropValue, getUserProps } from '../properties-service/properties-service';
import {
  addRowsToTopOfSheet,
  addToRepliesArray,
  getNumRowsAndColsFromArray,
  getSheetByName,
  setValuesInRangeAndSortSheet,
} from '../sheets/sheets';
import { calcAverage, getDomainFromEmailAddress, getEmailFromString, regexEmail, regexSalary } from '../utils/utils';

type EmailReplySendArray = [emailAddress: string, replyOrNew: EmailDataToSend][];

// export function setTemplateMsg({ subject, email }: { subject: string; email: string }) {
//   const drafts = GmailApp.getDrafts();
//   drafts.forEach((draft) => {
//     const { getFrom, getSubject, getId } = draft.getMessage();
//     if (subject === getSubject() && getFrom().match(email)) {
//       setUserProps({ messageId: getId(), draftId: draft.getId() });
//     }
//   });
// }

// type SetDraftTemplate = {
//   subject?: string;
//   email?: string;
// };

// function setInitialEmailProps({ subject = '', email = '' }: SetDraftTemplate) {
//   if (subject && email) {
//     setUserProps({ subject, email });
//     return;
//   }
//   // const userProps = PropertiesService.getUserProperties();
//   // if (!userProps.getProperty('subject') || !userProps.getProperty('email')) {
//   //   setUserProps({ subject: subject || CANNED_MSG_NAME, email: email || EMAIL_ACCOUNT });
//   // }
// }

// export function setDraftTemplateAutoResponder(obj: SetDraftTemplate = { email: '', subject: '' }) {
//   setInitialEmailProps(obj);
//   const props = getUserProps(['subject', 'email', 'draftId', 'messageId']);
//   let { subject, email, draftId, messageId } = props;
//   if (!draftId || !messageId) {
//     setTemplateMsg({ subject, email });
//   }
// }

function getDraftTemplateAutoResponder() {
  const { draftId } = getUserProps(['draftId']);
  if (!draftId) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(`Error!`, `Could not find the canned messaged need for the draft id`, ui.ButtonSet.OK);
    return;
  }
  const draft = GmailApp.getDraft(draftId);
  return draft;
}

export function sendTemplateEmail(recipient: string, subject: string, htmlBodyMessage?: string) {
  try {
    const name = getSingleUserPropValue('nameForEmail');
    const email = getSingleUserPropValue('email');
    if (!name) throw Error('You need to set a name to appear in the email');
    if (!email) throw Error('You need to set the email to send from');
    const draft = getDraftTemplateAutoResponder();
    const draftBody = draft && draft.getMessage().getBody();
    const body = htmlBodyMessage || draftBody ? draftBody : undefined;
    if (!body) throw Error('Could not find draft and send Email');
    GmailApp.sendEmail(recipient, subject, '', {
      from: email,
      htmlBody: body,
      name: name,
    });
  } catch (error) {
    console.error(error as any);
  }
}

function checkAndAddToEmailMap(emails: EmailReplySendArray) {
  emails.forEach(([email, data]) => {
    const domain = getDomainFromEmailAddress(email);

    if (!doNotSendMailAutoMap.has(email) && !doNotSendMailAutoMap.has(domain) && !emailsToSendMap.has(email)) {
      emailsToSendMap.set(email, data);
    }
  });
}

function buildEmailsObject(
  emailObj: Omit<EmailDataToSend, 'sendReplyEmail' | 'send' | 'isReplyorNewEmail' | 'emailSendTo'>,
  bodyEmails: string[],
  emailReplyTo: string
): EmailReplySendArray {
  const emailObject: Record<string, EmailDataToSend> = {};
  const isAutoResOn = getSingleUserPropValue('isAutoResOn');
  const onOrOff = isAutoResOn === 'On' ? true : false;
  bodyEmails.forEach((email) => {
    if (email !== emailObj.emailFrom && email !== emailReplyTo && !pendingEmailsToSendMap.has(email)) {
      const isReplyorNewEmail = 'new' as const;
      emailObject[email] = Object.assign({}, emailObj, {
        isReplyorNewEmail,
        sendReplyEmail: false,
        send: onOrOff,
        emailSendTo: email,
      });
    }
  });
  const isReplyorNewEmail = 'reply' as const;
  if (!pendingEmailsToSendMap.has(emailReplyTo)) {
    emailObject[emailReplyTo] = Object.assign({}, emailObj, {
      isReplyorNewEmail,
      sendReplyEmail: true,
      send: onOrOff,
      emailSendTo: emailReplyTo,
    });
  }
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
  const email = getSingleUserPropValue('email');
  if (!email) throw Error(`No email is set, you need to set an email, ${getAutoResponseMsgsFromThread.name}`);
  const ourEmailDomain = '@' + email.split('@')[1].toString();

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

type EmailListItem = [
  EmailThreadId: string,
  EmailMessageId: string,
  Date: GoogleAppsScript.Base.Date,
  From: string,
  ReplyTo: string,
  PersonCompanyName: string,
  Subject: string,
  BodyEmails: string | undefined,
  Body: string,
  Salary: string | undefined,
  ThreadPermalink: string,
  HasEmailResponse: string | false
];

export function extractDataFromEmailSearch(
  email: string,
  labelToSearch: string,
  _event?: GoogleAppsScript.Events.TimeDriven
) {
  try {
    const emailsForList: EmailListItem[] = [];

    // Exclude this label:
    // (And creates it if it doesn't exist)
    // return;

    // Send our response email and label it responded to
    // const threads = GmailApp.search(
    //   "-subject:'re:' -is:chats -is:draft has:nouserlabels -label:" + LABEL_NAME + ' to:(' + EMAIL_ACCOUNT + ')'
    // );
    const threads = GmailApp.search(`label:${labelToSearch} to:(${email})`);
    // 'recruiters-linkedin-recruiters');
    // + ' to:(' + EMAIL_ACCOUNT + ')');

    let salaries: number[] = [];
    threads.forEach((thread, _threadIndex) => {
      // if (threadIndex > 200) return;

      const emailMessageCount = thread.getMessageCount();
      const [firstMsg, ...restMsgs] = thread.getMessages();

      const firstMsgId = firstMsg.getId();

      const autoResponseMsg = emailMessageCount > 1 ? updateRepliesColumnIfMessageHasReplies(firstMsgId, restMsgs) : [];

      const from = firstMsg.getFrom();
      const emailThreadId = thread.getId();
      const emailThreadPermaLink = thread.getPermalink();
      const emailSubject = thread.getFirstMessageSubject();

      const emailBody = firstMsg.getPlainBody();
      const replyTo = firstMsg.getReplyTo();

      /** Use as a backup in case other split methods fail */
      // const emailFrom = [...new Set(from.match(regexEmail))];
      // const emailReplyTo = [...new Set(replyTo.match(regexEmail))];
      const emailFrom = getEmailFromString(from);
      const personFrom = from.split('<', 1)[0].trim();
      const emailReplyTo = replyTo ? getEmailFromString(replyTo) : emailFrom;

      const bodyEmails = [...new Set(emailBody.match(regexEmail))];
      const salaryAmount = emailBody.match(regexSalary);
      const emailMessageId = firstMsg.getId();
      const date = firstMsg.getDate();

      if (isDomainEmailInDoNotTrackSheet(emailFrom)) return;

      // const isNoReplyLinkedIn = from.match(/noreply@linkedin\.com/gi);
      // if (isNoReplyLinkedIn) return;

      checkAndAddToEmailMap(
        buildEmailsObject(
          {
            date,
            emailThreadId,
            emailBody,
            emailSubject,
            emailFrom,
            inResponseToEmailMessageId: emailThreadId,
            personFrom,
            emailThreadPermaLink,
          },
          bodyEmails,
          emailReplyTo
        )
      );

      salaryAmount && salaryAmount.length > 0 && salaries.push(calcAverage(salaryAmount));

      if (emailmessagesIdMap.has(emailThreadId)) return;

      emailsForList.push([
        emailThreadId,
        emailMessageId,
        date,
        emailFrom,
        emailReplyTo.toString(),
        personFrom,
        emailSubject,
        bodyEmails.length > 0 ? bodyEmails.toString() : undefined,
        emailBody,
        salaryAmount ? salaryAmount.toString() : undefined,
        emailThreadPermaLink,
        autoResponseMsg.length > 0 ? getToEmailArray(autoResponseMsg) : false,
      ]);

      // Add label to email for exclusion
      // thread.addLabel(label);
    });

    writeEmailsListToAutomationSheet(emailsForList);
    addSentEmailsToDoNotReplyMap(Array.from(emailsToSendMap.keys()));

    console.log({ salaries: calcAverage(salaries) });
  } catch (error) {
    console.error(error as any);
  }
}

function writeEmailsListToAutomationSheet(emailsForList: EmailListItem[]) {
  const autoResultsListSheet = getSheetByName('Automated Results List');
  if (!autoResultsListSheet) throw Error('Cannot find Automated Results List Sheet');
  if (emailsForList.length > 0) {
    const { numCols, numRows } = getNumRowsAndColsFromArray(emailsForList);
    addRowsToTopOfSheet(numRows, autoResultsListSheet);
    setValuesInRangeAndSortSheet(numRows, numCols, emailsForList, autoResultsListSheet, { sortByCol: 3, asc: false });
  }
}
