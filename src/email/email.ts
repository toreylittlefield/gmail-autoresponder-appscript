import { getProps, setInitialEmailProps, setUserProps } from '../properties-service/properties-service';

export function setTemplateMsg({ subject, email }: { subject: string; email: string }) {
  const drafts = GmailApp.getDrafts();
  drafts.forEach((draft) => {
    const { getFrom, getSubject, getId } = draft.getMessage();
    if (subject === getSubject() && getFrom().match(email)) {
      setUserProps({ messageId: getId(), draftId: draft.getId() });
    }
  });
}

export function getDraft() {
  setInitialEmailProps();
  const props = getProps(['subject', 'email', 'draftId', 'messageId']);
  console.log({ props });
  let { subject, email, draftId, messageId } = props;
  if (!draftId || !messageId) {
    setTemplateMsg({ subject, email });
    const draftObj = getProps(['draftId']);
    draftId = draftObj.draftId;
  }
  const draft = GmailApp.getDraft(draftId);
  return draft;
}
