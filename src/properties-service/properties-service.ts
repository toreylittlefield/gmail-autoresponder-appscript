import { allSheets } from '../variables/publicvariables';

type Sheets = typeof allSheets[number];

export type UserRecords =
  | 'subject'
  | 'email'
  | 'draftId'
  | 'messageId'
  | 'spreadsheetId'
  | 'nameForEmail'
  | 'labelToSearch'
  | 'labelId'
  | 'filterId'
  | 'isAutoResOn'
  | 'currentCalendarName';

export type UserPropsKeys = UserRecords | Sheets;

export function setUserProps(props: Partial<Record<UserPropsKeys, string>>) {
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperties(props);
}
export function getUserProps<T extends UserPropsKeys>(keys: T[]) {
  const userProps = PropertiesService.getUserProperties();
  const props: Partial<Record<T, string>> = {};
  keys.forEach((key) => {
    const value = userProps.getProperty(key);
    if (value) props[key] = value;
  });
  return props;
}

export function getSingleUserPropValue(key: UserPropsKeys) {
  const userProps = PropertiesService.getUserProperties();
  const value = userProps.getProperty(key);
  return value;
}
