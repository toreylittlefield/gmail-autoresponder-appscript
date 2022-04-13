import {
  doNotTrackMap,
  doNotSendMailAutoMap,
  emailmessagesIdMap,
  pendingEmailsToSendMap,
  alwaysAllowMap,
} from '../global/maps';
import { getUserProps } from '../properties-service/properties-service';
import { getAllDataFromSheet, SheetNames } from '../sheets/sheets';

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

export function getDomainFromEmailAddress(email: string) {
  const domain = `@${email.split('@')[1]}`;
  return domain;
}

export const regexEmail = /([a-zA-Z0-9+._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi;

export const regexSalary =
  /\$[1-2][0-9][0-9][-\s][1-2][0-9][0-9]|[1-2][0-9][0-9][-\s]\[1-2][0-9][0-9]|[1-2][0-9][0-9]k/gi;

export const getEmailFromString = (str: string) => str.split('<')[1].replace('>', '').trim();

type MapNames =
  | 'emailmessagesIdMap'
  | 'doNotTrackMap'
  | 'doNotSendMailAutoMap'
  | 'pendingEmailsToSendMap'
  | 'alwaysAllowMap';

export function initialGlobalMap(mapName: MapNames) {
  try {
    const getSheetData = (sheetName: SheetNames) => {
      const sheetData = getAllDataFromSheet(sheetName);
      if (!sheetData) throw Error(`Cannot initialize ${mapName}, could not get ${sheetName} sheet data`);
      return sheetData;
    };
    switch (mapName) {
      case 'emailmessagesIdMap':
        getSheetData('Automated Results List').forEach(([emailId], index) =>
          emailmessagesIdMap.set(emailId, index + 2)
        );
        break;
      case 'doNotTrackMap':
        getSheetData('Do Not Track List').forEach(([domainOrEmail]) => doNotTrackMap.set(domainOrEmail, true));
        break;
      case 'alwaysAllowMap':
        getSheetData('Always Autorespond List').forEach(([domainOrEmail]) => alwaysAllowMap.set(domainOrEmail, true));
        break;
      case 'doNotSendMailAutoMap':
        getSheetData('Do Not Autorespond List').forEach(([domain, _, count]) =>
          doNotSendMailAutoMap.set(domain, count)
        );
        break;
      case 'pendingEmailsToSendMap':
        getSheetData('Pending Emails To Send').forEach(
          ([_send, _emailThreadId, _inResponseToEmailMessageId, _isReplyOrNewEmail, _date, _emailFrom, sendToEmail]) =>
            pendingEmailsToSendMap.set(sendToEmail, true)
        );
        break;
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
