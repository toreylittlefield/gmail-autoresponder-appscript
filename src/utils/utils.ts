import {
  doNotTrackMap,
  doNotSendMailAutoMap,
  emailThreadIdsMap,
  pendingEmailsToSendMap,
  alwaysAllowMap,
  ValidRowToWriteInSentSheet,
  sentEmailsBySentMessageIdMap,
  sentEmailsByDomainMap,
  followUpSheetMessageIdMap,
  emailThreadsIdAppliedLinkedInMap,
} from '../global/maps';
import { getUserProps } from '../properties-service/properties-service';
import { getAllDataFromSheet, getAllHeaderColNumsAndLetters, SheetNames } from '../sheets/sheets';
import {
  ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME,
  AUTOMATED_RECEIVED_SHEET_NAME,
  DO_NOT_EMAIL_AUTO_SHEET_NAME,
  DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME,
  FOLLOW_UP_EMAILS_SHEET_NAME,
  LINKEDIN_APPLIED_JOBS_SHEET_NAME,
  PENDING_EMAILS_TO_SEND_SHEET_NAME,
  SENT_SHEET_HEADERS,
  SENT_SHEET_NAME,
} from '../variables/publicvariables';

export function calcAverage(numbersArray: any[]): number {
  return numbersArray.reduce((acc, curVal, index, array) => {
    acc += typeof curVal === 'number' ? curVal : parseInt(curVal);
    if (index === array.length - 1) {
      return (acc = Math.round(acc / array.length));
    }
    return acc;
  }, 0);
}

export function salariesToNumbers(salaryRegexMatch: RegExpMatchArray) {
  if (salaryRegexMatch.length === 0) return;
  const getDigits = salaryRegexMatch.map((val) => val.match(/[1-2][0-9][0-9]/gi));
  if (getDigits.length === 0) return;
  return calcAverage(getDigits);
}

export function getAtDomainFromEmailAddress(email: string) {
  const atDomainAddress = `@${email.split('@')[1]}`;
  return atDomainAddress;
}

export function getDomainFromEmailAddress(email: string) {
  const domain = email.split('@')[1];
  return domain;
}

export const regexEmail = /([a-zA-Z0-9+._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi;

export const regexSalary =
  /\$[1-2][0-9][0-9][-\s][1-2][0-9][0-9]|[1-2][0-9][0-9][-\s]\[1-2][0-9][0-9]|[1-2][0-9][0-9]k/gi;

export const regexValidUSPhoneNumber = /^(?:\+?1[-.●]?)?\(?([0-9]{3})\)?[-.●\s]?([0-9]{3})[-.●\s]?([0-9]{4})/gim;

export const getEmailFromString = (str: string) => str.split('<')[1].replace('>', '').trim();
export const getPhoneNumbersFromString = (str: string) => {
  const numbers = str.match(regexValidUSPhoneNumber);
  if (!numbers) return 'null';
  return numbers.reduce((acc, curVal, _index, arr) => {
    if (arr.length === 1) {
      return acc;
    }
    return (acc = `${acc}, ${curVal}`);
  });
};

type LinkedAppliedEmailData = {
  companyName: string;
  jobPosition: string;
  pointOfContact: string;
  viewJobOnLinkedInURL: string;
};
export function getDataFromLinkedInAppliedEmail({
  emailBody,
  emailSubject,
}: {
  emailBody: string;
  emailSubject: string;
}): LinkedAppliedEmailData {
  const { viewJobOnLinkedInURL = '' } = getViewJobURL(emailBody);
  const pointOfContact = '';
  const { companyName = '', jobPosition = '' } = getCompanyNameAndPositionFromSubject(emailSubject);
  return { viewJobOnLinkedInURL, pointOfContact, companyName, jobPosition };
}

function getViewJobURL(emailBody: string) {
  const [firstJob] = emailBody.match(/^(?:view job:\s)(.*)\b$/gim) as string[];
  const viewJobOnLinkedInURL = firstJob.replace(/View job:\s/g, '').trim();
  return { viewJobOnLinkedInURL };
}

function getCompanyName(emailSubject: string) {
  let companyName: string | RegExpMatchArray | null = emailSubject.match(/(at\s\b)(\w+\b)/g);
  if (!companyName) companyName = '';
  else {
    companyName = companyName
      .toString()
      .replace(/at\s\b/g, '')
      .trim();
  }
  return companyName;
}

function getJobPosition(emailSubject: string) {
  let jobPostion: string | RegExpMatchArray | null = emailSubject.match(/(for\s\b)(.*)(at\s\b)/g);
  if (!jobPostion) jobPostion = '';
  else {
    jobPostion = jobPostion
      .toString()
      .replace(/at|for/g, '')
      .trim();
  }
  return jobPostion;
}

function getCompanyNameAndPositionFromSubject(emailSubject: string) {
  const companyName = getCompanyName(emailSubject);
  const jobPosition = getJobPosition(emailSubject);

  return { companyName, jobPosition };
}

type MapNames =
  | 'autoReceivedSheetEmailThreadIdsMap'
  | 'followUpSheetEmailThreadIdsMap'
  | 'doNotTrackMap'
  | 'doNotSendMailAutoMap'
  | 'pendingEmailsToSendMap'
  | 'followUpSheetMessageIdMap'
  | 'alwaysAllowMap'
  | 'sentEmailsBySentMessageIdMap'
  | 'sentEmailsByDomainMap'
  | 'emailThreadsIdAppliedLinkedInMap';

export function initialGlobalMap(mapName: MapNames) {
  try {
    const getSheetData = (sheetName: SheetNames) => {
      const sheetData = getAllDataFromSheet(sheetName);
      if (!sheetData) throw Error(`Cannot initialize ${mapName}, could not get ${sheetName} sheet data`);
      return sheetData;
    };
    switch (mapName) {
      case 'autoReceivedSheetEmailThreadIdsMap':
        getSheetData(`${AUTOMATED_RECEIVED_SHEET_NAME}`).forEach(([emailThreadId, messageId], index) =>
          emailThreadIdsMap.set(emailThreadId, { rowNumber: index + 2, emailMessageId: messageId })
        );
        break;
      case 'emailThreadsIdAppliedLinkedInMap':
        getSheetData(LINKEDIN_APPLIED_JOBS_SHEET_NAME).forEach(([emailThreadId, emailMessageId], index) =>
          emailThreadsIdAppliedLinkedInMap.set(emailThreadId, { rowNumber: index + 2, emailMessageId })
        );
        break;
      case 'doNotTrackMap':
        getSheetData(DO_NOT_TRACK_DOMAIN_LIST_SHEET_NAME).forEach(([domainOrEmail]) =>
          doNotTrackMap.set(domainOrEmail, true)
        );
        break;
      case 'alwaysAllowMap':
        getSheetData(ALWAYS_RESPOND_DOMAIN_LIST_SHEET_NAME).forEach(([domainOrEmail]) =>
          alwaysAllowMap.set(domainOrEmail, true)
        );
        break;
      case 'doNotSendMailAutoMap':
        getSheetData(DO_NOT_EMAIL_AUTO_SHEET_NAME).forEach(([domain, _, count]) =>
          doNotSendMailAutoMap.set(domain, count)
        );
        break;
      case 'pendingEmailsToSendMap':
        getSheetData(PENDING_EMAILS_TO_SEND_SHEET_NAME).forEach(
          ([_send, _emailThreadId, _inResponseToEmailMessageId, _isReplyOrNewEmail, _date, _emailFrom, sendToEmail]) =>
            pendingEmailsToSendMap.set(sendToEmail, true)
        );
        break;
      case 'followUpSheetMessageIdMap':
        {
          const headersColumns = getAllHeaderColNumsAndLetters({ sheetName: FOLLOW_UP_EMAILS_SHEET_NAME });
          const emailMessageIdColNumber = headersColumns['Email Message Id'].colNumber;

          getSheetData(FOLLOW_UP_EMAILS_SHEET_NAME).forEach((row) => {
            followUpSheetMessageIdMap.set(row[emailMessageIdColNumber - 1], true);
          });
        }
        break;
      case 'sentEmailsBySentMessageIdMap':
        const data = getSheetData(SENT_SHEET_NAME) as ValidRowToWriteInSentSheet[];
        const headers = getAllHeaderColNumsAndLetters<typeof SENT_SHEET_HEADERS>({ sheetName: 'Sent Email Responses' });
        const colNumSentMessageId = headers['Sent Email Message Id'].colNumber;
        const headersAsKeys = Object.keys(headers) as typeof SENT_SHEET_HEADERS[number][];

        data.forEach((row) => {
          const sentMessageId = row[colNumSentMessageId - 1] as string;

          const keyHeaderRowValuePairs = row.reduce((acc, col, index) => {
            const key = headersAsKeys[index];
            acc[key] = col;
            return acc;
          }, {} as Record<typeof SENT_SHEET_HEADERS[number], ValidRowToWriteInSentSheet[number]>);

          sentEmailsBySentMessageIdMap.set(sentMessageId, {
            rowArray: row as ValidRowToWriteInSentSheet,
            rowObject: keyHeaderRowValuePairs,
          });
        });
        break;
      case 'sentEmailsByDomainMap': {
        const data = getSheetData(SENT_SHEET_NAME) as ValidRowToWriteInSentSheet[];
        const headers = getAllHeaderColNumsAndLetters<typeof SENT_SHEET_HEADERS>({
          sheetName: 'Sent Email Responses',
        });
        const colNumDomain = headers['Domain'].colNumber;
        const colNumSentDate = headers['Sent Email Message Date'].colNumber;
        const headersAsKeys = Object.keys(headers) as typeof SENT_SHEET_HEADERS[number][];

        data.forEach((row) => {
          const domain = row[colNumDomain - 1] as string;
          const sentDate = row[colNumSentDate - 1] as string;

          const keyHeaderRowValuePairs = row.reduce((acc, col, index) => {
            const key = headersAsKeys[index];
            acc[key] = col;
            return acc;
          }, {} as Record<typeof SENT_SHEET_HEADERS[number], ValidRowToWriteInSentSheet[number]>);

          const existingDomain = sentEmailsByDomainMap.get(domain);

          if (existingDomain) {
            if (existingDomain.rowObject['Sent Email Message Date'] < sentDate) {
              sentEmailsByDomainMap.set(domain, {
                rowObject: keyHeaderRowValuePairs,
                rowArray: row as ValidRowToWriteInSentSheet,
              });
            }
          } else {
            sentEmailsByDomainMap.set(domain, {
              rowObject: keyHeaderRowValuePairs,
              rowArray: row as ValidRowToWriteInSentSheet,
            });
          }
        });
        break;
      }
      default:
        break;
    }
  } catch (error) {
    console.error(error as any);
  }
}

export function hasAllRequiredUserProps() {
  const ui = SpreadsheetApp.getUi();
  const userConfiguration = getUserProps(['email', 'nameForEmail', 'labelToSearch', 'subject', 'draftId']);
  const userConfigKeysLength = Object.keys(userConfiguration).length;
  if (userConfigKeysLength === 0) {
    ui.alert(
      `Error`,
      `No User Configuration Set. Please set your email and other user configurations before syncing emails`,
      ui.ButtonSet.OK
    );
    return false;
  }
  if (userConfigKeysLength > 0) {
    const { email = '', draftId = '', labelToSearch = '', nameForEmail = '', subject = '' } = userConfiguration;
    const userConfigArray = Object.entries({ email, draftId, labelToSearch, nameForEmail, subject });
    const hasEveryUserProp = userConfigArray.every(([key, val]) => {
      if (!val) {
        ui.alert(
          `${key} is not set in your user configurations. Please set this value first before attempting to sync emails`
        );
        return false;
      }
      return true;
    });
    return hasEveryUserProp;
  }
  return false;
}

export function getFutureMonthAndYear(month: number) {
  if (month < 0 || month > 11 || typeof month !== 'number')
    throw Error(`${month} is invalid, must be a number between 0 & 11`);
  const nextYear = new Date().getFullYear();
  if (month < 11) {
    return { nextMonth: month + 1, nextYear: nextYear };
  }

  return { nextMonth: 0, nextYear: nextYear + 1 };
}

export function getCurrentMonthAndNextMonthDates() {
  const currentDate = new Date();
  const curMonthNum = currentDate.getMonth();

  const { nextMonth, nextYear } = getFutureMonthAndYear(curMonthNum);
  const day = currentDate.getDate();
  const nextDate = new Date(nextYear, nextMonth, day);
  return { currentDate, nextDate };
}
