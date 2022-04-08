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
  | 'filterId';

export type UserPropsKeys = UserRecords | Sheets;

export function setUserProps(props: Partial<Record<UserPropsKeys, string>>) {
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperties(props);
}

export function getUserProps(keys: UserPropsKeys[]) {
  const userProps = PropertiesService.getUserProperties();
  const props: Record<string, any> = {};
  keys.forEach((key) => {
    const value = userProps.getProperty(key);
    props[key] = value;
  });
  return props;
}

export function getSingleUserPropValue(key: UserPropsKeys) {
  const userProps = PropertiesService.getUserProperties();
  const value = userProps.getProperty(key);
  return value;
}
