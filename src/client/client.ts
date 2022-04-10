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

type OnSucessPayload = {
  emailAliases: string[];
  mainEmail: string;
  currentEmailUserStore: string;
  nameForEmail: string;
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

function setEmailForm(emailAliases: string[], mainEmail: string, currentEmailUserStore: string) {
  const emailValues = Object.values([mainEmail, ...emailAliases]);

  const fragment = document.createDocumentFragment();

  emailValues.forEach((email) => {
    const option = document.createElement('OPTION') as HTMLOptionElement;
    option.value = email;

    const input = document.getElementById('email-input');
    if (input instanceof HTMLInputElement) {
      input.focus();
    }
    fragment.appendChild(option);
  });

  const datalist = document.getElementById('available-emails');
  if (datalist) {
    datalist.appendChild(fragment);
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
  const { emailAliases, mainEmail, currentEmailUserStore, nameForEmail } = userProperties;

  setEmailForm(emailAliases, mainEmail, currentEmailUserStore);
  setNameForm(nameForEmail);
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

function handleSubmitForm(formObject: any) {
  toggleLoading(formIdFromEvent, true);
  google.script.run.withSuccessHandler(processCallbackSuccess).processFormEventsFromPage(formObject);
}

function processCallbackSuccess(formObject: Record<UserPropsKeys, string>) {
  const [key] = Object.keys(formObject);
  const [value] = Object.values(formObject);
  if (formObject.email) {
    setEmail(formObject.email);
  }
  if (formObject.subject) {
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
