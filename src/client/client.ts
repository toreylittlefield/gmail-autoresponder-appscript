type FormIds = 'email-form' | 'name-form' | 'canned-form' | 'label-form';

type UserPropsKeys =
  | 'subject'
  | 'email'
  | 'draftId'
  | 'messageId'
  | 'spreadsheetId'
  | 'nameForEmail'
  | 'labelToSearch'
  | 'labelId'
  | 'filterId';

type DraftsToPick = { subject: string; draftId: string };

type OnSucessPayload = {
  emailAliases: string[];
  mainEmail: string;
  currentEmailUserStore: string;
  nameForEmail: string;
  subject: string;
  subjectsToPickFromDrafts: DraftsToPick[];
};

let formIdFromEvent: FormIds;

function attachFormSubmitListeners() {
  const forms = Array.from(document.querySelectorAll('form'));
  forms.forEach((form) => {
    const id = form.id as FormIds;

    form.addEventListener('submit', function (event) {
      event.preventDefault();
      formIdFromEvent = id;
      handleSubmitForm(this);
    });
  });
}

function appendToDataList(elementId: string, elementToAppend: Node) {
  const datalist = document.getElementById(elementId);
  if (datalist) {
    datalist.appendChild(elementToAppend);
  }
}

function createChildElements(
  arrValues: string[],
  dataSet?: { key: string; dataSetValues: string[] },
  elementType: keyof HTMLElementTagNameMap = 'option'
) {
  const fragment = document.createDocumentFragment();

  arrValues.forEach((value, index) => {
    const option = document.createElement(elementType) as HTMLOptionElement;
    if (dataSet) {
      const value = dataSet.dataSetValues[index];
      const key = dataSet.key;
      option.dataset[key] = value;
    }
    option.value = value;

    fragment.appendChild(option);
  });

  return fragment;
}

function setSubject(subject: string) {
  const titleElement = document.querySelector(`#subject-form #current-value`);
  if (titleElement) titleElement.textContent = `${subject}`;
}

function setSubjectForm(subject: string, subjectsToPickFromDrafts: DraftsToPick[]) {
  const subjects = subjectsToPickFromDrafts.map(({ subject }) => subject);
  const draftIds = subjectsToPickFromDrafts.map(({ draftId }) => draftId);
  appendToDataList('available-subjects', createChildElements(subjects, { key: 'draftId', dataSetValues: draftIds }));

  const input = document.getElementById('subject-input');
  if (input instanceof HTMLInputElement) {
    input.addEventListener('change', function change(_event) {
      const list = this.list;
      if (list) {
        list.childNodes.forEach((child) => {
          if (child instanceof HTMLOptionElement) {
            if (child.dataset.draftId) {
              const draftId = child.dataset.draftId;
              if (child.value === this.value) {
                this.dataset.draftId = draftId;
                this.setCustomValidity('');
              } else {
                delete this.dataset.draftId;
              }
            }
          }
        });
      }
      if (this.dataset.draftId == null) {
        this.setCustomValidity('Please Select A Template From The List');
        this.checkValidity();
      }
    });
  }

  setSubject(subject);
}

function setEmailForm(emailAliases: string[], mainEmail: string, currentEmailUserStore: string) {
  const emailValues = Object.values([mainEmail, ...emailAliases]);

  appendToDataList('available-emails', createChildElements(emailValues));

  const input = document.getElementById('email-input');
  if (input instanceof HTMLInputElement) {
    input.focus();
  }
  setEmail(currentEmailUserStore);
}

function setEmail(currentEmailUserStore: string) {
  const titleElement = document.querySelector(`#email-form #current-value`);
  if (titleElement) titleElement.textContent = `${currentEmailUserStore}`;
}

function setNameForm(nameForEmail: string) {
  setName(nameForEmail);
}

function setName(nameForEmail: string) {
  const titleElement = document.querySelector(`#name-form #current-value`);
  if (titleElement) titleElement.textContent = `${nameForEmail}`;
}

function onSuccessGetUserProperties(userProperties: OnSucessPayload) {
  const { emailAliases, mainEmail, currentEmailUserStore, nameForEmail, subject, subjectsToPickFromDrafts } =
    userProperties;

  setEmailForm(emailAliases, mainEmail, currentEmailUserStore);
  setNameForm(nameForEmail);
  setSubjectForm(subject, subjectsToPickFromDrafts);
}

function toggleLoading(formId: string, disabled: boolean) {
  const loader = document.querySelector(`#${formId} #loader`);
  const submitButton = document.querySelector(`#${formId} button[type='submit']`) as HTMLButtonElement;
  submitButton.disabled = disabled;
  if (loader) {
    loader.classList.toggle('hide');
    loader.classList.toggle('show');
  }
}

function getDraftIAndSubjectFromForm(formObject: HTMLFormElement) {
  const [input] = Array.from(formObject).filter((element) => element.id === 'subject-input');
  if (input instanceof HTMLInputElement) {
    return { subject: input.value, draftId: input.dataset.draftId };
  }
  return null;
}

function handleSubmitForm(formObject: HTMLFormElement) {
  let payload: HTMLFormElement | Partial<Record<UserPropsKeys, string>> = formObject;
  if (formObject.id === 'subject-form') {
    const result = getDraftIAndSubjectFromForm(formObject);
    if (result) payload = result;
  }
  toggleLoading(formIdFromEvent, true);
  google.script.run.withSuccessHandler(processCallbackSuccess).processFormEventsFromPage(payload);
}

function processCallbackSuccess(formObject: Record<UserPropsKeys, string>) {
  const [key] = Object.keys(formObject);
  const [value] = Object.values(formObject);
  if (formObject.email) {
    setEmail(formObject.email);
  }
  if (formObject.subject) {
    setSubject(formObject.subject);
  }
  if (formObject.labelToSearch) {
  }
  if (formObject.nameForEmail) {
    setName(formObject.nameForEmail);
  }
  alert(`${key} Changed!
        Your ${key} is now set to:
            ${value}`);
  toggleLoading(formIdFromEvent, false);
}

function onLoadWrapper() {
  attachFormSubmitListeners();
  setTimeout(() => {
    const loader = document.getElementById('p2');
    if (loader) {
      loader.classList.toggle('hide');
    }
    const app = document.getElementById('main-wrapper');
    if (app) {
      app.classList.toggle('hide');
    }
  }, 500);
}
window.addEventListener('load', onLoadWrapper);

google.script.run.withSuccessHandler(onSuccessGetUserProperties).getUserPropertiesForPageModal();
