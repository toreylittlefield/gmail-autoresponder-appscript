import { emailmessagesIdMap } from '../email/email';
import { doNotTrackMap, getAllDataFromSheet, SheetNames } from '../sheets/sheets';

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

export const getEmailFromString = (str: string) => str.split('<')[1].replace('>', '');

type MapNames = 'emailmessagesIdMap' | 'doNotTrackMap';

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
      default:
        break;
    }
  } catch (error) {
    console.error(error as any);
  }
}
