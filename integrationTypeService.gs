/**
 * @fileoverview Code related to information that the Odo Add-on
 * may show in various contexts if the chosen integration type is
 * INTEGRATION_TYPE.GENERIC_SERVICE
 */

/**
 * Function used to return the record data as a formatted Card to be
 * displayed. Called from integrationTypeAll.gs as a context specific
 * handler for this integration.
 *
 * @param {string} Calling context (i.e. CALL_CONTEXT.GMAIL_VIEW)
 *
 * @return {CardService.Card}
 */
function buildServiceBasicCard(context) {
  let config = getConfig();

  let integrationData;

  // if there is previously stored data for this integration type,
  // display it. else, show the default values.
  if (
    config.saved &&
    config.integrationType === INTEGRATION_TYPE.GENERIC_SERVICE
  ) {
    integrationData = config.integrationData;
  } else {
    integrationData = serviceBasicGetDefaultConfig();
  }

  let buttonUrl = integrationData.buttonUrl;
  let message = integrationData.message;
  let buttonText = integrationData.buttonText;

  buttonUrl = findAndReplaceMergeKeys(buttonUrl, { uriEncodeValue: true });
  message = findAndReplaceMergeKeys(message);
  buttonText = findAndReplaceMergeKeys(buttonText);

  let card = CardService.newCardBuilder();

  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText(message));

  section.addWidget(
    CardService.newTextButton()
      .setText(buttonText)
      .setOpenLink(CardService.newOpenLink().setUrl(buttonUrl))
  );

  let button2Url = integrationData.button2Url;
  let message2 = integrationData.message2;
  let button2Text = integrationData.button2Text;

  if (button2Text && button2Url) {
    button2Url = findAndReplaceMergeKeys(button2Url, { uriEncodeValue: true });

    button2Text = findAndReplaceMergeKeys(button2Text);

    if (message2) {
      message2 = findAndReplaceMergeKeys(message2);
      section.addWidget(CardService.newTextParagraph().setText(message2));
    }

    section.addWidget(
      CardService.newTextButton()
        .setText(button2Text)
        .setOpenLink(CardService.newOpenLink().setUrl(button2Url))
    );
  }

  card.addSection(section);

  return card.build();
}

/**
 * Creates and returns the card that gives the user options to configure
 * the Service integration. Called from integrationTypeAll.gs based on the
 * value of the 'buildConfigureIntegrationCard' parameter.
 * 
 * @return {CardService.Card}
 */
function buildServiceBasicConfigureCard() {
  let config = getConfig();
  let integrationData;

  integrationData = getConfigIntegrationData(INTEGRATION_TYPE.GENERIC_SERVICE);
  /*
  if (
    config.saved &&
    config.integrationType === INTEGRATION_TYPE.GENERIC_SERVICE
  ) {
    integrationData = config.integrationData;
  } else {
    integrationData = serviceBasicGetDefaultConfig();
  }
  */

  let buttonUrl = integrationData.buttonUrl;
  let message = integrationData.message;
  let buttonText = integrationData.buttonText;

  let button2Url = integrationData.button2Url;
  let message2 = integrationData.message2;
  let button2Text = integrationData.button2Text;

  let card = CardService.newCardBuilder();
  let section = CardService.newCardSection();
  let input;

  section.addWidget(CardService.newTextParagraph().setText('Primary Button:'));

  input = CardService.newTextInput()
    .setFieldName('message')
    .setTitle('Message')
    .setValue(message);
  section.addWidget(input);

  input = CardService.newTextInput()
    .setFieldName('buttonText')
    .setTitle('Button Text')
    .setValue(buttonText);
  section.addWidget(input);

  input = CardService.newTextInput()
    .setFieldName('buttonUrl')
    .setTitle('Button URL')
    .setValue(buttonUrl);
  section.addWidget(input);

  section.addWidget(
    CardService.newTextParagraph().setText('<br><br>Optional Secondary Button:')
  );

  input = CardService.newTextInput()
    .setFieldName('message2')
    .setTitle('Message 2')
    .setValue(message2);
  section.addWidget(input);

  input = CardService.newTextInput()
    .setFieldName('button2Text')
    .setTitle('Button 2 Text')
    .setValue(button2Text);
  section.addWidget(input);

  input = CardService.newTextInput()
    .setFieldName('button2Url')
    .setTitle('Button 2 URL')
    .setValue(button2Url);
  section.addWidget(input);

  card.addSection(section);

  return card;
}


/**
 * Function that gets called for this particular integration when user 
 * clicks '←  Done' button in integration configuration card. Saves the 
 * selections and returns them as an object to be stored in the
 * 'integrationData' field of the config object if/when the user saves their
 * configurations.
 * 
 * This is the handler that's defined as 
 * 'saveConfigureIntegrationSelections' in integrationTypeAll.gs.
 * 
 * @param {object} formInputs - Contains user selections
 * 
 * @return {object}
 */
function saveServiceBasicConfigureSelections(formInputs) {
  let message = formInputs['message'].stringInputs.value[0];
  let buttonText = formInputs['buttonText'].stringInputs.value[0];
  let buttonUrl = formInputs['buttonUrl'].stringInputs.value[0];

  let message2 = '';
  let button2Text = '';
  let button2Url = '';

  if (formInputs['message2']) {
    message2 = formInputs['message2'].stringInputs.value[0];
  }
  if (formInputs['button2Text']) {
    button2Text = formInputs['button2Text'].stringInputs.value[0];
  }
  if (formInputs['button2Url']) {
    button2Url = formInputs['button2Url'].stringInputs.value[0];
  }

  let integrationData = {
    message: message,
    buttonText: buttonText,
    buttonUrl: buttonUrl,
    message2: message2,
    button2Text: button2Text,
    button2Url: button2Url,
  };

  return integrationData;
}

/**
 * Function that returns default configuration fields and values for
 * an integration type of INTEGRATION_TYPE.GENERIC_SERVICE, to be stored
 * as the 'integrationData' field in 'config'. Called when setting
 * up default configuration.
 *
 * @return {Object}
 */
function serviceBasicGetDefaultConfig() {
  return {
    message:
      'Click the button below to launch a ' +
      '{{toolName}} session ' +
      'with <b>{{token || participants}}</b>.',

    buttonText: 'Launch {{toolName}}',

    // note: any merge tags in the buttonUrl will get URI encoded
    buttonUrl:
      THIRD_PARTY_SERVICE_URL +
      '?loadingMessage=Loading%20' +
      '{{toolName}}%20with%20' +
      '{{token || participants}}',

    message2: '',
    button2Text: '',
    button2Url: '',
  };
}
