/**
 * @typedef {'email-form' | 'name-form' | 'canned-form' | 'label-form'} FormIds
 */

/**
    * @typedef {  'subject' | 'email'
  | 'draftId'
  | 'messageId'
  | 'spreadsheetId'
  | 'nameForEmail'
  | 'labelToSearch'
  | 'labelId'
  | 'filterId'} UserPropsKeys
    */

let forms = [];

/** @type FormIds */
let formIdFromEvent = '';

function attachFormSubmitListeners() {
  forms = Array.from(document.querySelectorAll('form'));
  forms.forEach((form) => {
    /** @type {FormIds} */
    const id = form.id;

    form.addEventListener('submit', function (event) {
      event.preventDefault();
      formIdFromEvent = id;
      handleSubmitForm(this);
    });
  });
}

function setEmailForm(emailAliases, mainEmail, currentEmailUserStore) {
  const emailValues = Object.values([mainEmail, ...emailAliases]);

  const fragment = document.createDocumentFragment();

  emailValues.forEach((email) => {
    const option = document.createElement('OPTION');
    option.value = email;

    const input = document.getElementById('email-input');
    if (input instanceof HTMLInputElement) {
      input.focus();
    }
    fragment.appendChild(option);
  });

  const datalist = document.getElementById('available-emails');
  datalist.appendChild(fragment);

  setEmail(currentEmailUserStore);
}

function setEmail(currentEmailUserStore) {
  document.querySelector(`#email-form #current-value`).textContent = `${currentEmailUserStore}`;
}

function setNameForm(nameForEmail) {
  setName(nameForEmail);
}

function setName(nameForEmail) {
  document.querySelector(`#name-form #current-value`).textContent = `${nameForEmail}`;
}

function onSuccessGetUserProperties(userProperties) {
  const { emailAliases, mainEmail, currentEmailUserStore, nameForEmail } = userProperties;

  setEmailForm(emailAliases, mainEmail, currentEmailUserStore);
  setNameForm(nameForEmail);
}

/**
 * @param {FormIds} formId
 * @param {boolean} disabled
 */
function toggleLoading(formId, disabled) {
  const loader = document.querySelector(`#${formId} #loader`);
  const submitButton = document.querySelector(`#${formId} button[type='submit']`);
  loader.classList.toggle('hide');
  loader.classList.toggle('show');
  submitButton.disabled = disabled;
}

function handleSubmitForm(formObject) {
  toggleLoading(formIdFromEvent, true);
  google.script.run.withSuccessHandler(processCallbackSuccess).processFormEventsFromPage(formObject);
}

/**
 * @param {Object} formObject
 * @param {string} [formObject.email]
 * @param {string} [formObject.subject]
 * @param {string} [formObject.labelToSearch]
 * @param {string} [formObject.nameForEmail]
 */
function processCallbackSuccess(formObject) {
  console.log({ formObject });
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

function processEmailForm({ email }) {
  setEmail(email);
}

function onLoadWrapper() {
  attachFormSubmitListeners();
  setTimeout(() => {
    document.getElementById('p2').classList.toggle('hide');
    document.getElementById('main-wrapper').classList.toggle('hide');
  }, 500);
}
window.addEventListener('load', onLoadWrapper);

google.script.run.withSuccessHandler(onSuccessGetUserProperties).getUserPropertiesForPageModal();
