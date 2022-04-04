import { CANNED_MSG_NAME, EMAIL_ACCOUNT } from '../variables/privatevariables';

export function setUserProps(props: Record<string, any>) {
  const userProps = PropertiesService.getUserProperties();

  userProps.setProperties(props);
}

export function getProps(keys: string[]) {
  const userProps = PropertiesService.getUserProperties();
  const props: Record<string, any> = {};
  keys.forEach((key) => {
    const value = userProps.getProperty(key);
    props[key] = value;
  });
  return props;
}

export function setInitialEmailProps() {
  const userProps = PropertiesService.getUserProperties();

  if (!userProps.getProperty('subject') || !userProps.getProperty('email')) {
    setUserProps({ subject: CANNED_MSG_NAME, email: EMAIL_ACCOUNT });
  }
}
