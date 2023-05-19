/**
 * @fileoverview Code related to configuration of the Odo Add-on.
 */

/**
 * Card builder that's called when user selects the Univeral action "Configure"
 *
 * @return {CardService.Card} Card to show for configuration interface.
 */
function onOdoConfig() {
  let wasConfigured = true;

  let config = getConfig();
  if (!config.saved) {
    wasConfigured = false;
  }

  // clear out any prior selections from the "Customize Integrations" card
  let up = PropertiesService.getUserProperties();
  up.deleteProperty(PROP_STRINGIFIED_INTEGRATION_CUSTOMIZATIONS);

  let card = CardService.newCardBuilder();
  let header = CardService.newCardHeader();
  header
    .setTitle('Configure Odo')
    .setSubtitle('Configure your demo Add-on')
    .setImageStyle(CardService.ImageStyle.SQUARE)
    .setImageUrl(ODO_ICON);

  card.setHeader(header);
  card.setName('mainConfigurationCard');

  //// General Config Section ////
  let appearanceSection = CardService.newCardSection();
  appearanceSection.setHeader('General');

  let customerNameWidget = CardService.newTextInput()
    .setFieldName('customerName')
    .setValue(config.customerName)
    .setTitle('Customer Name');
  appearanceSection.addWidget(customerNameWidget);

  let customerLogoWidget = CardService.newTextInput()
    .setFieldName('customerLogoUrl')
    .setValue(config.customerLogoUrl)
    .setTitle('Customer Logo URL');
  appearanceSection.addWidget(customerLogoWidget);

  let toolNameWidget = CardService.newTextInput()
    .setFieldName('toolName')
    .setValue(config.toolName)
    .setTitle('Tool Name');
  appearanceSection.addWidget(toolNameWidget);

  let welcomeMessageWidget = CardService.newTextInput()
    .setFieldName('welcomeSplashMessage')
    .setValue(config.welcomeSplashMessage)
    .setTitle('Welcome Splash Card Message');
  appearanceSection.addWidget(welcomeMessageWidget);

  let dateFormatWidget = CardService.newSelectionInput()
    .setFieldName('dateFormat')
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Date Format');

  let supportedDateFormats = [
    'MM/dd/yyyy',
    'dd MMM, yyyy',
    'dd-MM-yyyy',
    'MM-dd-yyyy',
    'dd/MM/yyyy',
    'MMM dd, yyyy',
  ];

  for (let i = 0; i < supportedDateFormats.length; i++) {
    let selected = supportedDateFormats[i] === config.dateFormat;

    dateFormatWidget.addItem(
      supportedDateFormats[i],
      supportedDateFormats[i],
      selected
    );
  }

  appearanceSection.addWidget(dateFormatWidget);
  appearanceSection.addWidget(
    CardService.newTextParagraph().setText('<br><br>')
  );
  card.addSection(appearanceSection);
  //////////////////////

  //// Integration Section ////
  let integrationSection = CardService.newCardSection();
  integrationSection.setHeader('Simulated Integration');

  let integrationTypeWidget = CardService.newSelectionInput()
    .setFieldName('integrationType')
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Integration Type');

  for (integrationType in INTEGRATION_HOOKS) {
    if (!INTEGRATION_HOOKS.hasOwnProperty(integrationType)) {
      continue;
    }

    let hooks = INTEGRATION_HOOKS[integrationType];

    let selected = integrationType === config.integrationType;

    integrationTypeWidget.addItem(
      integrationTypeToPrintableString(integrationType),
      integrationType,
      selected
    );
  }

  integrationSection.addWidget(integrationTypeWidget);

  let customizeIntegrationAction = CardService.newAction();

  customizeIntegrationAction.setFunctionName(
    'onUserSelectedCustomizeIntegration'
  );

  integrationSection.addWidget(
    CardService.newTextButton()
      .setText('Customize Integration →')
      .setOnClickAction(customizeIntegrationAction)
  );

  card.addSection(integrationSection);
  //////////////////////

  //// Footer with Button(s) ////
  let footer = CardService.newFixedFooter();

  let saveAction = CardService.newAction().setFunctionName(
    'onUserSelectedConfigSaveSelections'
  );

  footer.setPrimaryButton(
    CardService.newTextButton().setText('Save').setOnClickAction(saveAction)
  );

  if (wasConfigured) {
    let resetAction = CardService.newAction().setFunctionName(
      'onUserSelectedConfigResetSelections'
    );

    footer.setSecondaryButton(
      CardService.newTextButton()
        .setText('Reset All')
        .setOnClickAction(resetAction)
    );
  }

  card.setFixedFooter(footer);
  //////////////////////

  return card.build();
}


/**
 * Builds a card that shows the user options to customize their integration.
 * Contents will vary based on the chosen integration type. Called when
 * user clicks "Customize Integration" from "Configure Odo" card.
 *
 * @param {Object} event Event information passed in by Card framework.
 *
 * @return {CardService.Card} Card to show result.
 */
function onUserSelectedCustomizeIntegration(event) {
  // get the user selected integration type to customize
  let formInputs = event.commonEventObject.formInputs;
  let selectedIntegrationType =
    formInputs.integrationType.stringInputs.value[0];

  return buildCustomizeIntegrationCard(selectedIntegrationType);
}

/**
 * Saves the user selected configurations from the Configure Odo page
 *
 * @param {Object} event Event information passed in by Card framework.
 *
 * @return {CardService.Card} Card to show result.
 */
function onUserSelectedConfigSaveSelections(event) {
  let config = {};
  let formInputs = event.commonEventObject.formInputs;

  config.saved = true;

  config.customerName = formInputs.customerName.stringInputs.value[0];
  config.customerLogoUrl = formInputs.customerLogoUrl.stringInputs.value[0];
  config.toolName = formInputs.toolName.stringInputs.value[0];
  config.integrationType = formInputs.integrationType.stringInputs.value[0];
  config.dateFormat = formInputs.dateFormat.stringInputs.value[0];
  config.welcomeSplashMessage =
    formInputs.welcomeSplashMessage.stringInputs.value[0];

  let up = PropertiesService.getUserProperties();
  let integrationDataStr = up.getProperty(
    PROP_STRINGIFIED_INTEGRATION_CUSTOMIZATIONS
  );

  if (integrationDataStr) {
    let integrationData = JSON.parse(integrationDataStr);
    console.log(JSON.stringify(integrationData))
    config.integrationData = integrationData;
  } else {
    let hooks = INTEGRATION_HOOKS[config.integrationType];
    config.integrationData = hooks.defaultIntegrationConfig();
  }

  saveConfig(config);

  return buildRefreshNeededCard();
}

/**
 * Resets the Odo configurations
 *
 * @return {CardService.Card} Card to show result
 */
function onUserSelectedConfigResetSelections() {
  resetConfig();

  return buildRefreshNeededCard();
}

/**
 * Returns the current Odo configuration, which is an Object (struct) with
 * various members related to Odo's config. Can be modified and saved via a
 * call to saveConfig().
 *
 * @return {Object}
 */
function getConfig() {
  let up = PropertiesService.getUserProperties();

  let configStr = up.getProperty(USER_PROPERTY_CONFIG);

  let config = null;

  if (configStr) {
    config = JSON.parse(configStr);
  } else {
    return _getDefaultConfig();
  }

  return config;
}

/**
 * Takes a configuration Object (struct) and saves it. Will
 * be returned with the same data the next time getConfig() is called.
 * @param {Object} config - The configuration object to be saved.
 *
 */
function saveConfig(config) {
  config.saved = true;
  let configStr = JSON.stringify(config);

  let up = PropertiesService.getUserProperties();
  up.setProperty(USER_PROPERTY_CONFIG, configStr);
}

/**
 * Function that fully resets the configuration.
 *
 */
function resetConfig() {
  let up = PropertiesService.getUserProperties();

  // Delete all stored properties. This will include the USER_PROPERTY_CONFIG,
  // as well as any integration specific properties.
  up.deleteAllProperties();
}

/**
 * Dump contents of config to console. For debugging purposes only.
 */
function _dumpConfig() {
  let config = getConfig();

  console.log(JSON.stringify(config));
}

/**
 * Private function that constructs and returns a default config (i.e. user
 * has not configured Odo yet). Used internally by getConfig().
 *
 * @return {Object}
 */
function _getDefaultConfig() {
  let defaultIntegrationType = INTEGRATION_TYPE.RECORDS_BASED;

  let config = {
    saved: false, // set to true first time configi saved by user,
    welcomeSplashShhown: false,
    customerName: DEFAULT_CUSTOMER_NAME,
    customerLogoUrl: DEFAULT_CUSTOMER_LOGO_URL,
    toolName: DEFAULT_CUSTOMER_TOOL_NAME,
    integrationType: defaultIntegrationType,
    dateFormat: 'MMM dd, yyyy',
    welcomeSplashMessage:
      'Welcome to {{toolName}} for Workspace. Click the button below '
      + 'to get going!',
  };

  let hooks = INTEGRATION_HOOKS[defaultIntegrationType];
  config.integrationData = hooks.defaultIntegrationConfig();

  return config;
}
