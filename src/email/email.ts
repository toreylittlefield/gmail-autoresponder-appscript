import {
  alwaysAllowMap,
  doNotSendMailAutoMap,
  doNotTrackMap,
  EmailDataToSend,
  emailmessagesIdMap,
  emailsToSendMap,
  pendingEmailsToSendMap,
} from '../global/maps';
import { getSingleUserPropValue, getUserProps } from '../properties-service/properties-service';
import { addToRepliesArray, writeEmailsListToAutomationSheet } from '../sheets/sheets';
import { calcAverage, getDomainFromEmailAddress, getEmailFromString, regexEmail, regexSalary } from '../utils/utils';
import { LABEL_NAME } from '../variables/publicvariables';

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

function getParamsForSendingEmails(htmlBodyMessage?: string) {
  const name = getSingleUserPropValue('nameForEmail');
  const email = getSingleUserPropValue('email');
  if (!name) throw Error('You need to set a name to appear in the email');
  if (!email) throw Error('You need to set the email to send from');
  const draft = getDraftTemplateAutoResponder();
  const draftBody = draft && draft.getMessage().getBody();
  const body = htmlBodyMessage || draftBody ? draftBody : undefined;
  if (!body) throw Error('Could not find draft and send Email');
  const gmailAdvancedOptions: GoogleAppsScript.Gmail.GmailAdvancedOptions = {
    from: email,
    htmlBody: body,
    name: name,
  };
  return gmailAdvancedOptions;
}

export type DraftAttributeArray = [
  draftId: string,
  draftSentMessageId: string,
  draftMessageDate: GoogleAppsScript.Base.Date,
  draftMessageSubject: string,
  draftMessageFrom: string,
  draftMessageTo: string,
  draftMessageBody: string
];
function getDraftAttrArrayToWriteToSheet(draft: GoogleAppsScript.Gmail.GmailDraft): DraftAttributeArray {
  const { getSubject, getDate, getFrom, getTo, getPlainBody, getId } = draft.getMessage();
  return [draft.getId().toString(), getId().toString(), getDate(), getSubject(), getFrom(), getTo(), getPlainBody()];
}

export function createNewDraftMessage(
  recipient: string,
  subject: string,
  gmailAdvancedOptions: GoogleAppsScript.Gmail.GmailAdvancedOptions
): DraftAttributeArray {
  const newDraft = GmailApp.createDraft(recipient, subject, '', gmailAdvancedOptions);
  return getDraftAttrArrayToWriteToSheet(newDraft);
}

function draftReplyToMessage(gmailMessageId: string, htmlBodyMessage?: string): DraftAttributeArray {
  const gmailAdvancedOptions = getParamsForSendingEmails(htmlBodyMessage);

  const gmailMessage = GmailApp.getMessageById(gmailMessageId);
  if (!gmailMessage) throw Error(`Cant find Gmail message: ${gmailMessageId} to create a draft reply`);

  const draftReply = gmailMessage.createDraftReply('', gmailAdvancedOptions);
  return getDraftAttrArrayToWriteToSheet(draftReply);
}

type SendTemplateOptions =
  | { type: 'replyDraft'; gmailMessageId: string; htmlBodyMessage?: string; recipient?: string; subject?: string }
  | { type: 'newDraftEmail'; gmailMessageId?: string; htmlBodyMessage?: string; recipient: string; subject: string }
  | { type: 'sendNewEmail'; gmailMessageId?: string; htmlBodyMessage?: string; recipient: string; subject: string };

export function createOrSentTemplateEmail({
  type,
  htmlBodyMessage,
  gmailMessageId = '',
  recipient = '',
  subject = '',
}: SendTemplateOptions): DraftAttributeArray | undefined {
  try {
    const gmailAdvancedOptions = getParamsForSendingEmails(htmlBodyMessage);
    switch (type) {
      case 'newDraftEmail':
        return createNewDraftMessage(recipient, subject, gmailAdvancedOptions);

      case 'replyDraft':
        return draftReplyToMessage(gmailMessageId, htmlBodyMessage);
      case 'sendNewEmail':
        GmailApp.sendEmail(recipient, subject, '', gmailAdvancedOptions);
        break;
    }
    return undefined;
  } catch (error) {
    console.error(error as any);
    return undefined;
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
    if (count == null && !alwaysAllowMap.has(domain)) {
      doNotSendMailAutoMap.set(domain, 0);
    }
    return;
  });
}

export function getToEmailArray(emailMessages: GoogleAppsScript.Gmail.GmailMessage[]) {
  return emailMessages.map((emailMsg) => emailMsg.getTo()).toString();
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

function isDomainEmailInDoNotTrackSheet(fromEmail: string) {
  const domain = getDomainFromEmailAddress(fromEmail);

  if (doNotTrackMap.has(domain) || doNotSendMailAutoMap.has(fromEmail)) return true;
  /**TODO: Can be optimized in future if slow perf */
  if (Array.from(doNotTrackMap.keys()).filter((domainOrEmailKey) => fromEmail.match(domainOrEmailKey)).length > 0)
    return true;
  return false;
}

export type EmailListItem = [
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
            inResponseToEmailMessageId: emailMessageId,
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

export function getUserLabels() {
  const currentLabel = getSingleUserPropValue('labelToSearch');
  const labels = GmailApp.getUserLabels();

  return { currentLabel, userLabels: labels.length > 0 ? labels.map(({ getName }) => getName()) : [] };
}

export function createFilterAndLabel(currentEmail: string) {
  const me = Session.getActiveUser().getEmail();

  const gmailUser = Gmail.Users as GoogleAppsScript.Gmail.Collection.UsersCollection;

  const labelsCollection = gmailUser.Labels as GoogleAppsScript.Gmail.Collection.Users.LabelsCollection;
  const newLabel = labelsCollection.create(
    {
      color: {
        backgroundColor: '#42d692',
        textColor: '#ffffff',
      },
      name: LABEL_NAME,
      labelListVisibility: 'labelShow',
      messageListVisibility: 'show',
      type: 'user',
    },
    me
  ) as GoogleAppsScript.Gmail.Schema.Label;

  const userSettings = gmailUser.Settings as GoogleAppsScript.Gmail.Collection.Users.SettingsCollection;
  const filters = userSettings.Filters as GoogleAppsScript.Gmail.Collection.Users.Settings.FiltersCollection;
  const newFilter = filters.create(
    {
      action: {
        addLabelIds: [newLabel.id as string],
      },
      criteria: {
        to: currentEmail,
      },
    },
    me
  );

  const resFilter = filters.get(me, newFilter.id as string);
  const resLabel = labelsCollection.get(me, newLabel.id as string);
  return { resFilter, resLabel };
}

export function getUserEmails() {
  const emailAliases = GmailApp.getAliases();
  const mainEmail = Session.getActiveUser().getEmail();
  const currentEmailUserStore = getSingleUserPropValue('email') || 'none set';
  return { emailAliases, mainEmail, currentEmailUserStore };
}

export function getUserNameForEmail() {
  const nameForEmail = getSingleUserPropValue('nameForEmail');
  return { nameForEmail };
}

type DraftsToPick = { subject: string; draftId: string; subjectBody: string };

export function getUserCannedMessage(): { draftsList: DraftsToPick[]; subject: string } {
  const drafts = GmailApp.getDrafts();
  const draftsFilteredByEmail = drafts.filter((draft) => {
    const { getTo, getSubject } = draft.getMessage();
    return getTo() === '' && getSubject();
  });
  const draftsList = draftsFilteredByEmail.map(({ getId, getMessage }) => ({
    draftId: getId(),
    subject: getMessage().getSubject().trim(),
    subjectBody: getMessage().getPlainBody(),
  }));

  const subject = getSingleUserPropValue('subject');

  return { draftsList, subject: subject || '' };
}
