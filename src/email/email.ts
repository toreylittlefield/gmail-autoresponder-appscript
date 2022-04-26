import {
  alwaysAllowMap,
  doNotSendMailAutoMap,
  doNotTrackMap,
  EmailDataToSend,
  emailsToAddToPendingSheetMap,
  emailThreadIdsMap,
  followUpSheetMessageIdMap,
  pendingEmailsToSendMap,
  sentEmailsByDomainMap,
  sentEmailsBySentMessageIdMap,
  ValidRowToWriteInSentSheet,
} from '../global/maps';
import { getSingleUserPropValue, getUserProps } from '../properties-service/properties-service';
import {
  addToRepliesArray,
  writeDomainsListToDoNotRespondSheet,
  writeEmailDataToReceivedAutomationSheet,
  writeMessagesToFollowUpEmailsSheet,
  writeToSentEmailsSheet,
} from '../sheets/sheets';
import {
  calcAverage,
  getAtDomainFromEmailAddress,
  getDomainFromEmailAddress,
  getEmailFromString,
  getPhoneNumbersFromString,
  initialGlobalMap,
  regexEmail,
  regexSalary,
} from '../utils/utils';
import { RECEIVED_MESSAGES_LABEL_NAME } from '../variables/publicvariables';

type EmailReplySendArray = [emailAddress: string, replyOrNew: EmailDataToSend][];

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

function getParamsForSendingEmails(personFrom: string, htmlBodyMessage?: string) {
  const name = getSingleUserPropValue('nameForEmail');
  const email = getSingleUserPropValue('email');
  if (!name) throw Error('You need to set a name to appear in the email');
  if (!email) throw Error('You need to set the email to send from');
  const draft = getDraftTemplateAutoResponder();
  const draftBody = draft && draft.getMessage().getBody();
  const body = htmlBodyMessage || draftBody ? draftBody : undefined;
  if (!body) throw Error('Could not find draft and send Email');
  const bodyWithName = body.replace(/Hello!/g, `Hello ${personFrom}!`);
  const gmailAdvancedOptions: GoogleAppsScript.Gmail.GmailAdvancedOptions = {
    from: email,
    htmlBody: bodyWithName,
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

function draftReplyToMessage(
  gmailMessageId: string,
  personFrom: string,
  htmlBodyMessage?: string
): DraftAttributeArray {
  const gmailAdvancedOptions = getParamsForSendingEmails(personFrom, htmlBodyMessage);

  const gmailMessage = GmailApp.getMessageById(gmailMessageId);
  if (!gmailMessage) throw Error(`Cant find Gmail message: ${gmailMessageId} to create a draft reply`);

  const draftReply = gmailMessage.createDraftReply('', gmailAdvancedOptions);
  return getDraftAttrArrayToWriteToSheet(draftReply);
}

type SendTemplateOptions =
  | {
      type: 'replyDraftEmail';
      gmailMessageId: string;
      htmlBodyMessage?: string;
      recipient?: string;
      subject?: string;
      personFrom: string;
    }
  | {
      type: 'newDraftEmail';
      gmailMessageId?: string;
      htmlBodyMessage?: string;
      recipient: string;
      subject: string;
      personFrom: string;
    }
  | {
      type: 'sendNewEmail';
      gmailMessageId?: string;
      htmlBodyMessage?: string;
      recipient: string;
      subject: string;
      personFrom: string;
    };

export function createOrSentTemplateEmail({
  type,
  htmlBodyMessage,
  personFrom,
  gmailMessageId = '',
  recipient = '',
  subject = '',
}: SendTemplateOptions): DraftAttributeArray | undefined {
  try {
    const gmailAdvancedOptions = getParamsForSendingEmails(personFrom, htmlBodyMessage);
    switch (type) {
      case 'newDraftEmail':
        return createNewDraftMessage(recipient, subject, gmailAdvancedOptions);

      case 'replyDraftEmail':
        return draftReplyToMessage(gmailMessageId, personFrom, htmlBodyMessage);
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

function addValidEmailDataToPendingSheetMap(emails: EmailReplySendArray) {
  emails.forEach(([email, data]) => {
    const domain = getAtDomainFromEmailAddress(email);

    if (
      !doNotSendMailAutoMap.has(email) &&
      !doNotSendMailAutoMap.has(domain) &&
      !emailsToAddToPendingSheetMap.has(email)
    ) {
      emailsToAddToPendingSheetMap.set(email, data);
    }
  });
}

export function getEmailByThreadAndAddToMap(
  emailThreadId: string,
  emailObjData: Omit<EmailDataToSend, 'send' | 'isReplyorNewEmail' | 'emailSendTo'>,
  bodyEmails: string[],
  emailReplyTo: string
) {
  const thread = GmailApp.getThreadById(emailThreadId);
  if (!thread) return;
  addValidEmailDataToPendingSheetMap(buildEmailsObjectForReplies(emailObjData, bodyEmails, emailReplyTo));
}

function buildEmailsObjectForReplies(
  emailObj: Omit<EmailDataToSend, 'send' | 'isReplyorNewEmail' | 'emailSendTo'>,
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
        send: onOrOff,
        emailSendTo: email,
      });
    }
  });
  const isReplyorNewEmail = 'reply' as const;
  if (!pendingEmailsToSendMap.has(emailReplyTo)) {
    emailObject[emailReplyTo] = Object.assign({}, emailObj, {
      isReplyorNewEmail,
      send: onOrOff,
      emailSendTo: emailReplyTo,
    });
  }
  return Object.entries(emailObject);
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
  const messageAlreadyExists = emailThreadIdsMap.has(firstMsgId);

  const autoResponseMsg = getAutoResponseMsgsFromThread(restMsgs);

  if (autoResponseMsg.length > 0 && messageAlreadyExists) {
    const data = emailThreadIdsMap.get(firstMsgId);
    addToRepliesArray(data ? data.rowNumber : 0, autoResponseMsg);
  }

  return autoResponseMsg;
}

function isDomainEmailInDoNotTrackSheet(fromEmail: string) {
  const domain = getAtDomainFromEmailAddress(fromEmail);

  if (alwaysAllowMap.has(domain) || alwaysAllowMap.has(fromEmail)) return false;

  if (doNotTrackMap.has(domain) || doNotSendMailAutoMap.has(fromEmail)) return true;
  /**TODO: Can be optimized in future if slow perf */
  if (Array.from(doNotTrackMap.keys()).filter((domainOrEmailKey) => fromEmail.match(domainOrEmailKey)).length > 0)
    return true;
  return false;
}

export type EmailReceivedSheetRowItem = [
  EmailThreadId: string,
  EmailMessageId: string,
  Date: GoogleAppsScript.Base.Date,
  From: string,
  ReplyTo: string,
  Subject: string,
  BodyEmails: string | undefined,
  Body: string,
  Domain: string,
  PersonCompanyName: string,
  PhoneNumbers: string,
  Salary: string | undefined,
  ThreadPermalink: string,
  HasEmailResponse: string | false,
  LastSentDate: false | GoogleAppsScript.Base.Date,
  LastSentThreadId: false | string,
  LastSentDraftSubject: false | string,
  LastSentPerson: false | string
];

export function extractGMAILDataForNewMessagesReceivedSearch(
  email: string,
  labelToSearch: string,
  labelToExclude?: string,
  _event?: GoogleAppsScript.Events.TimeDriven
) {
  try {
    const emailsForList: EmailReceivedSheetRowItem[] = [];

    // Exclude this label:
    // (And creates it if it doesn't exist)
    // return;

    // Send our response email and label it responded to
    // const threads = GmailApp.search(
    //   "-subject:'re:' -is:chats -is:draft has:nouserlabels -label:" + LABEL_NAME + ' to:(' + EMAIL_ACCOUNT + ')'
    // );
    const threads = GmailApp.search(`label:${labelToSearch} -label:${labelToExclude} to:(${email})`);

    let salaries: number[] = [];
    threads.forEach((thread, _threadIndex) => {
      const {
        autoResString,
        salaryRegexArray,
        bodyEmailsString,
        bodyEmails,
        date,
        domainFromEmail,
        emailBody,
        emailFrom,
        emailMessageId,
        emailReplyTo,
        emailReplyToString,
        emailSubject,
        emailThreadId,
        emailThreadPermaLink,
        personFrom,
        phoneNumbers,
        salaryAmount,
        lastSentDate,
        sentThreadId,
        sentDraftSubject,
        sentToPerson,
      } = makeEmailValidResponseObject(thread);

      if (isDomainEmailInDoNotTrackSheet(emailFrom)) return;

      addValidEmailDataToPendingSheetMap(
        buildEmailsObjectForReplies(
          {
            date,
            emailThreadId,
            emailBody,
            emailSubject,
            emailFrom,
            inResponseToEmailMessageId: emailMessageId,
            personFrom,
            emailThreadPermaLink,
            domain: domainFromEmail,
            phoneNumbers,
            salary: salaryAmount,
          },
          bodyEmails,
          emailReplyTo
        )
      );

      salaryRegexArray && salaryRegexArray.length > 0 && salaries.push(calcAverage(salaryRegexArray));

      // TODO: Check For Replies / Follow Up Messages?
      if (emailThreadIdsMap.has(emailThreadId)) return;

      emailsForList.push([
        emailThreadId,
        emailMessageId,
        date,
        emailFrom,
        emailReplyToString,
        emailSubject,
        bodyEmailsString,
        emailBody,
        domainFromEmail,
        personFrom,
        phoneNumbers,
        salaryAmount,
        emailThreadPermaLink,
        autoResString,
        lastSentDate,
        sentThreadId,
        sentDraftSubject,
        sentToPerson,
      ]);

      // Add label to email for exclusion
      // thread.addLabel(label);
    });

    writeEmailDataToReceivedAutomationSheet(emailsForList);

    console.log({ salaries: calcAverage(salaries) });
  } catch (error) {
    console.error(error as any);
  }
}

/**
 * @TODO needs logic to extract sent messages and only the data from sent messages before usage
 */
export function extractGMAILDataSentEmailsSearch(
  email: string,
  labelToSearch: string,
  labelToExclude?: string,
  _event?: GoogleAppsScript.Events.TimeDriven
) {
  try {
    initialGlobalMap('sentEmailsBySentMessageIdMap');
    const validRowInSentSheet: ValidRowToWriteInSentSheet[] = [];

    // Exclude this label:
    // (And creates it if it doesn't exist)
    // return;

    // Send our response email and label it responded to
    // const threads = GmailApp.search(
    //   "-subject:'re:' -is:chats -is:draft has:nouserlabels -label:" + LABEL_NAME + ' to:(' + EMAIL_ACCOUNT + ')'
    // );
    const sentThreads = GmailApp.search(`label:${labelToSearch} -label:${labelToExclude} to:(${email})`);

    sentThreads.forEach((sentThread, _threadIndex) => {
      const messagesInSentThread = sentThread.getMessages();

      messagesInSentThread.forEach((message) => {
        const from = message.getFrom();
        const fromEmail = getEmailFromString(from);
        if (fromEmail !== email) return;

        const messageIdToCompare = message.getId();

        if (sentEmailsBySentMessageIdMap.has(messageIdToCompare)) return;

        const isNewOrReply = sentThread.getMessageCount() === 1 ? 'new' : 'reply';
        const sentMessageTo = getEmailFromString(message.getTo());
        // const emailSubject = sentThread.getFirstMessageSubject();

        const {
          date: sentMessageDate,
          emailBody: sentMessageBody,
          emailFrom: sentMessageFrom,
          emailMessageId: sentMessageId,
          emailSubject: sentMessageSubject,
        } = getMessagePropertiesForResponseObject(message);

        const {
          emailThreadId,
          emailMessageId,
          date,
          emailFrom,
          emailReplyToString,
          emailSubject,
          emailBody,
          domainFromEmail,
          personFrom,
          phoneNumbers,
          salaryAmount,
          emailThreadPermaLink,
        } = makeEmailValidResponseObject(sentThread);

        validRowInSentSheet.push([
          emailThreadId,
          emailMessageId,
          isNewOrReply,
          date,
          emailFrom,
          emailReplyToString,
          emailSubject,
          emailBody,
          domainFromEmail,
          personFrom,
          phoneNumbers,
          salaryAmount,
          emailThreadPermaLink,
          false,
          '',
          sentMessageId,
          sentMessageDate,
          sentMessageSubject,
          sentMessageFrom,
          sentMessageTo,
          sentMessageBody,
          '',
          false,
          emailThreadId,
          sentMessageId,
          sentMessageDate,
          emailThreadPermaLink,
        ]);
      });
      /**
       *   emailThreadId: string,
  inResponseToEmailMessageId: string,
  isReplyorNewEmail: 'new' | 'reply',
  date: GoogleAppsScript.Base.Date,
  emailFrom: string,
  emailSendTo: string,
  emailSubject: string,
  emailBody: string,
  domain: string,
  personFrom: string,
  phoneNumbers: string,
  salary: string,
  emailThreadPermaLink: string,
  deleteDraft: boolean,
  draftId: string,
  draftSentMessageId: string,
  draftMessageDate: GoogleAppsScript.Base.Date,
  draftMessageSubject: string,
  draftMessageFrom: string,
  draftMessageTo: string,
  draftMessageBody: string,
  viewDraftInGmail: string,
  manuallyMoveDraftToSent: boolean,
  sentThreadId: string,
  sentEmailMessageId: string,
  sentEmailMessageDate: GoogleAppsScript.Base.Date,
  sentThreadPermaLink: string
       */

      // Add label to email for exclusion
      // thread.addLabel(label);
    });
    if (validRowInSentSheet.length > 0) {
      writeToSentEmailsSheet(validRowInSentSheet);
      writeDomainsListToDoNotRespondSheet();
    }
  } catch (error) {
    console.error(error as any);
  }
}

/**
 * get the latest sent data in the sent email sheet for this domain if it exists
 * @description ignores "linkedin.com" domain
 */
function getDataInSentMailByDomainMap(domain: string) {
  const dataToReturn = {
    lastSentDate: false,
    sentThreadId: false,
    sentDraftSubject: false,
    sentToPerson: false,
  } as const;
  if (domain === 'linkedin.com') return dataToReturn;
  const lastSentData = sentEmailsByDomainMap.get(domain);
  if (!lastSentData) return dataToReturn;
  const {
    'Sent Email Message Date': lastSentDate,
    'Sent Thread Id': sentThreadId,
    'Draft Subject': sentDraftSubject,
    'Person / Company Name': sentToPerson,
  } = lastSentData.rowObject;
  return { lastSentDate, sentThreadId, sentDraftSubject, sentToPerson } as {
    lastSentDate: GoogleAppsScript.Base.Date;
    sentDraftSubject: string;
    sentThreadId: string;
    sentToPerson: string;
  };
}

export function getMessagePropertiesForResponseObject(emailMessage: GoogleAppsScript.Gmail.GmailMessage) {
  const emailMessageId = emailMessage.getId();

  const from = emailMessage.getFrom();

  const emailSubject = emailMessage.getSubject();

  const emailBody = emailMessage.getPlainBody();
  const replyTo = emailMessage.getReplyTo();

  /** Use as a backup in case other split methods fail */
  // const emailFrom = [...new Set(from.match(regexEmail))];
  // const emailReplyTo = [...new Set(replyTo.match(regexEmail))];
  const emailFrom = getEmailFromString(from);
  const personFrom = from.split('<', 1)[0].trim();
  const phoneNumbers = getPhoneNumbersFromString(emailBody);
  const emailReplyTo = replyTo ? getEmailFromString(replyTo) : emailFrom;

  const bodyEmails = [...new Set(emailBody.match(regexEmail))];
  const domainFromEmail = (() => {
    const emailSet =
      bodyEmails.length > 0 && getDomainFromEmailAddress(emailFrom) === 'linkedin.com'
        ? new Set(bodyEmails)
        : new Set([...bodyEmails, emailFrom]);
    const emailsArray = Array.from(emailSet).flatMap((email) => getDomainFromEmailAddress(email));
    return Array.from(new Set(emailsArray)).toString();
  })();
  const salaryRegexArray = emailBody.match(regexSalary);
  const salaryAmount = salaryRegexArray ? salaryRegexArray.toString() : '';
  const date = emailMessage.getDate();
  const bodyEmailsString = bodyEmails.length > 0 ? bodyEmails.toString() : undefined;

  const { lastSentDate, sentThreadId, sentDraftSubject, sentToPerson } = getDataInSentMailByDomainMap(domainFromEmail);

  return {
    emailMessageId,
    emailBody,
    emailSubject,
    personFrom,
    phoneNumbers,
    salaryRegexArray,
    emailFrom,
    bodyEmails,
    domainFromEmail,
    emailReplyTo,
    salaryAmount,
    date,
    bodyEmailsString,
    lastSentDate,
    sentThreadId,
    sentDraftSubject,
    sentToPerson,
  };
}

export function makeEmailValidResponseObject(thread: GoogleAppsScript.Gmail.GmailThread) {
  const emailMessageCount = thread.getMessageCount();
  const [firstMsg, ...restMsgs] = thread.getMessages();
  const firstMsgId = firstMsg.getId();

  const autoResponseMsg = emailMessageCount > 1 ? updateRepliesColumnIfMessageHasReplies(firstMsgId, restMsgs) : [];

  const emailThreadId = thread.getId();
  const emailThreadPermaLink = thread.getPermalink();
  const {
    emailMessageId,
    date,
    emailReplyTo,
    emailSubject,
    bodyEmailsString,
    salaryRegexArray,
    emailFrom,
    bodyEmails,
    domainFromEmail,
    lastSentDate,
    emailBody,
    personFrom,
    phoneNumbers,
    salaryAmount,
    sentDraftSubject,
    sentThreadId,
    sentToPerson,
  } = getMessagePropertiesForResponseObject(firstMsg);

  const autoResString = autoResponseMsg.length > 0 ? getToEmailArray(autoResponseMsg) : (false as const);

  return {
    emailThreadId,
    emailMessageId,
    date,
    emailFrom,
    emailReplyTo,
    emailReplyToString: emailReplyTo.toString(),
    emailSubject,
    bodyEmails,
    bodyEmailsString,
    emailBody,
    domainFromEmail,
    personFrom,
    phoneNumbers,
    salaryAmount,
    emailThreadPermaLink,
    autoResString,
    salaryRegexArray,
    lastSentDate,
    sentThreadId,
    sentDraftSubject,
    sentToPerson,
  };
}

export type ValidFollowUpSheetRowItem = [
  EmailThreadId: string,
  EmailMessageId: string,
  DateOfEmail: GoogleAppsScript.Base.Date,
  EmailFrom: string,
  EmailSendReplyTo: string,
  EmailSubject: string,
  EmailsInBodyList: string | undefined,
  EmailBody: string,
  DomainOfSender: string,
  PersonCompanyName: string,
  PhoneNumbers: string,
  Salary: string | undefined,
  EmailThreadPermalink: string,
  ViewInGmailLink: string,
  FollowUpToSentThreadId: string,
  FollowUpToSentEmailMessageId: string,
  FollowUpToSentEmailDate: GoogleAppsScript.Base.Date | undefined,
  FollowUpToSendThreadPermaLink: string,
  DateOfLastMessageInEmailThread: GoogleAppsScript.Base.Date,
  ArchiveThreadIdCheckbox: false,
  DeleteThreadIdCheckbox: false,
  RemoveGmailLabelCheckbox: false,
  AddFollowUpGmailLabelCheckbox: false,
  ManualRepliedToEmailCheckbox: false,
  ManualRepliedToEmailDateCheckbox: GoogleAppsScript.Base.Date | undefined
];

export function extractGMAILDataForFollowUpSearch(
  email: string,
  labelToSearch: string,
  labelToExclude?: string,
  _event?: GoogleAppsScript.Events.TimeDriven
) {
  try {
    const validRowsInFollowUpSheet: ValidFollowUpSheetRowItem[] = [];

    // Exclude this label:
    // (And creates it if it doesn't exist)
    // return;

    // Send our response email and label it responded to
    // const threads = GmailApp.search(
    //   "-subject:'re:' -is:chats -is:draft has:nouserlabels -label:" + LABEL_NAME + ' to:(' + EMAIL_ACCOUNT + ')'
    // );
    const sentThreads = GmailApp.search(`label:${labelToSearch} -label:${labelToExclude} to:(${email})`);

    sentThreads.forEach((sentThread, _threadIndex) => {
      let sentThreadId = '';
      let sentMessageId = '';
      let sentMessageDate: undefined | GoogleAppsScript.Base.Date = undefined;
      let sentThreadPermaLink = '';

      const messagesInSentThread = sentThread.getMessages();

      messagesInSentThread.forEach((message) => {
        const messageIdToCompare = message.getId();

        // return if message already exist in received automation sheet
        const receivedSheetData = emailThreadIdsMap.get(sentThread.getId());
        if (receivedSheetData && receivedSheetData.emailMessageId === messageIdToCompare) return;

        // already exists in the follow up sheet
        if (followUpSheetMessageIdMap.has(messageIdToCompare)) return;

        const from = message.getFrom();
        const fromEmail = getEmailFromString(from);

        if (fromEmail === email) {
          const sentData = sentEmailsBySentMessageIdMap.get(message.getId());
          if (sentData) {
            const {
              'Sent Thread Id': sentThreadIdData,
              'Sent Email Message Id': sentMessageIdData,
              'Sent Email Message Date': sentMessageDateData,
              'Sent Thread PermaLink': sentThreadPermaLinkData,
            } = sentData.rowObject;
            sentThreadId = sentThreadIdData as string;
            sentMessageId = sentMessageIdData as string;
            sentMessageDate = sentMessageDateData as GoogleAppsScript.Base.Date;
            sentThreadPermaLink = sentThreadPermaLinkData as string;
          }
          return;
        }
        const emailThreadId = sentThread.getId();
        // const emailSubject = sentThread.getFirstMessageSubject();
        const emailThreadPermaLink = sentThread.getPermalink();
        const lastMessageDate = sentThread.getLastMessageDate();

        const {
          bodyEmailsString,
          date,
          domainFromEmail,
          emailBody,
          emailFrom,
          emailMessageId,
          emailReplyTo,
          emailSubject,
          personFrom,
          phoneNumbers,
          salaryAmount,
        } = getMessagePropertiesForResponseObject(message);

        validRowsInFollowUpSheet.push([
          emailThreadId,
          emailMessageId,
          date,
          emailFrom,
          emailReplyTo,
          emailSubject,
          bodyEmailsString,
          emailBody,
          domainFromEmail,
          personFrom,
          phoneNumbers,
          salaryAmount,
          emailThreadPermaLink,
          `https://mail.google.com/mail/u/0/#label/auto-responder-sent-email-label/${emailMessageId}`,
          sentThreadId,
          sentMessageId,
          sentMessageDate,
          sentThreadPermaLink,
          lastMessageDate,
          false,
          false,
          false,
          false,
          false,
          undefined,
        ]);
      });

      // Add label to email for exclusion
      // thread.addLabel(label);
    });
    writeMessagesToFollowUpEmailsSheet(validRowsInFollowUpSheet);
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
      name: RECEIVED_MESSAGES_LABEL_NAME,
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
