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
