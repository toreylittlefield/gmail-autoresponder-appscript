import { getProps, setInitialEmailProps, setUserProps } from '../properties-service/properties-service';
import { EMAIL_ACCOUNT, NAME_TO_SEND_IN_EMAIL } from '../variables/privatevariables';

export function setTemplateMsg({ subject, email }: { subject: string; email: string }) {
  const drafts = GmailApp.getDrafts();
  drafts.forEach((draft) => {
    const { getFrom, getSubject, getId } = draft.getMessage();
    if (subject === getSubject() && getFrom().match(email)) {
      setUserProps({ messageId: getId(), draftId: draft.getId() });
    }
  });
}

export function setDraftTemplateAutoResponder() {
  setInitialEmailProps();
  const props = getProps(['subject', 'email', 'draftId', 'messageId']);
  let { subject, email, draftId, messageId } = props;
  if (!draftId || !messageId) {
    setTemplateMsg({ subject, email });
  }
}

function getDraftTemplateAutoResponder() {
  const { draftId } = getProps(['draftId']);
  const draft = GmailApp.getDraft(draftId);
  return draft;
}

export function sendTemplateEmail(recipient: string, subject: string, htmlBodyMessage?: string) {
  try {
    const body = htmlBodyMessage || getDraftTemplateAutoResponder().getMessage().getBody();
    if (!body) throw Error('Could not find draft and send Email');
    GmailApp.sendEmail(recipient, subject, '', {
      from: EMAIL_ACCOUNT,
      htmlBody: body,
      name: NAME_TO_SEND_IN_EMAIL,
    });
  } catch (error) {
    console.error(error as any);
  }
}
