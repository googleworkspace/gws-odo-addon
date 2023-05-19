/**
 * @fileoverview Code related to all integration types that the Odo Add-on
 * supports, mostly related to setting and saving the configurations for
 * those integration types.
 */

/**
 * Integration types that Odo can be configured to simulate.
 * @enum {string}
 */
const INTEGRATION_TYPE = {
  GENERIC_SERVICE: 'GENERIC_SERVICE', // see integrationTypeService.gs
  RECORDS_BASED: 'RECORDS_BASED', // see integrationTypeRecords.gs
  FILE_REPOSITORY: 'FILE_REPOSITORY', // see integrationFileRepository.gs
  IMAGE_LIBRARY: 'IMG_LIBRARY', // see integrationImageLibrary.gs
};

/**
 * Hooks that each integration type must implement of the following format:
 *
 *  [INTEGRATION_TYPE.XXX] {
 *      // Printable string for this integration type (shown in configuration
 *      // card)
 *      printableString: <string>,
 *
 *      // Function the returns a struct with a default configuration
 *      // for this integration type (to be used if user doesn't configure
 *      // it themselves, as well to define default values on the integration
 *      // configuration card. Stored in 'integrationData' field of
 *      // the 'config' object.
 *      //
 *      // @return {Object}
 *      defaultIntegrationConfig: <function>,
 *
 *      // Handler that builds a card with customization options. Shown
 *      // when user selects 'Customize Integration' from configuration card.
 *      //
 *      // @return {CardService.Card}
 *      buildConfigureIntegrationCard: <function>,
 *
 *      // Handler that saves results of integration configuration card (when
 *      // user clicks "← Done"). Passed a formInputs objects with user's
 *      // selections. Should return the struct/Object to be saved in the
 *      // 'integrationData' field of the 'config' object.
 *      // @param {Objects} formInputs
 *      //
 *      // @return {Object}
 *      saveConfigureIntegrationSelections: <function>,
 *
 *      // Define context specific handlers. All handlers get passed
 *      // the context (i.e. CALL_CONTEXT.DOCS) they were called from.
 *      //
 *      // Handler:
 *      //   @param {string} context
 *      //
 *      //   @return {CardService.Card}
 *      contextSpecificHandlers: {
 *        [CALL_CONTEXT.XXX] : <function>,
 *        [CALL_CONTEXT.DEFAULT] : <function>, // optional
 *      }
 *  }
 *
 */
const INTEGRATION_HOOKS = {
  [INTEGRATION_TYPE.RECORDS_BASED]: {
    printableString: 'Records Based',
    defaultIntegrationConfig: recordGetDefaultConfig,
    buildConfigureIntegrationCard: buildRecordsConfigureCard,
    saveConfigureIntegrationSelections: saveRecordsConfigureSelections,

    contextSpecificHandlers: {
      [CALL_CONTEXT.GMAIL_HOMEPAGE]: buildGmailHomepage,
      [CALL_CONTEXT.GMAIL_COMPOSE]: buildRecordCard,
      [CALL_CONTEXT.GMAIL_VIEW]: buildRecordCard,
      [CALL_CONTEXT.DOCS]: buildRecordCard,
      [CALL_CONTEXT.SHEETS]: buildRecordCard,
    },
  },

  [INTEGRATION_TYPE.GENERIC_SERVICE]: {
    printableString: 'Generic Service',
    defaultIntegrationConfig: serviceBasicGetDefaultConfig,
    buildConfigureIntegrationCard: buildServiceBasicConfigureCard,
    saveConfigureIntegrationSelections: saveServiceBasicConfigureSelections,

    // define context specific handlers. all handlers get passed
    // the context (i.e. CALL_CONTEXT.DOCS) they were called from.
    contextSpecificHandlers: {
      [CALL_CONTEXT.DEFAULT]: buildServiceBasicCard,
    },
  },

  [INTEGRATION_TYPE.FILE_REPOSITORY]: {
    printableString: 'File Repository',
    defaultIntegrationConfig: fileRepoGetDefaultConfig,
    buildConfigureIntegrationCard: buildFileRepoConfigureCard,
    saveConfigureIntegrationSelections: saveFileRepoConfigureSelections,

    // define context specific handlers. all handlers get passed
    // the context (i.e. CALL_CONTEXT.DOCS) they were called from.
    contextSpecificHandlers: {
      [CALL_CONTEXT.DRIVE] : buildFileRepoCard,
    },
  },

  [INTEGRATION_TYPE.IMAGE_LIBRARY]: {
    printableString: 'Image Library for Slides',
    defaultIntegrationConfig: imgLibraryGetDefaultConfig,
    buildConfigureIntegrationCard: buildImgLibraryConfigureCard,
    saveConfigureIntegrationSelections: saveImgLibraryConfigureSelections,

    // define context specific handlers. all handlers get passed
    // the context (i.e. CALL_CONTEXT.DOCS) they were called from.
    contextSpecificHandlers: {
      [CALL_CONTEXT.SLIDES] : buildImgLibraryCard,
    },
  },
};

/**
 * Function that converts an Integration Type enum into a printable
 * string that can be shown to the user.
 *
 * @param {string} Record type (i.e. INTEGRATION_TYPE.RECORDS_BASED)
 *
 * @return {string}
 */
function integrationTypeToPrintableString(integrationType) {
  let integrationConfig = INTEGRATION_HOOKS[integrationType];

  return integrationConfig.printableString;
}

/**
 * Function used to return the a formatted Card to be displayed.
 * The specific card returned will depend on the simulated integration type,
 * and the context it was called from. This function serves as a central
 * dispatch hub to generate most cards shown in the Add-on.
 *
 * @param {string} context Calling context (i.e. CALL_CONTEXT.GMAIL_MESSAGE)
 *
 * @return {string{CardService.Card}
 */
function buildIntegrationCard(context) {
  console.log('buildIntegrationCard: entering...');

  let config = getConfig();

  // if Odo yet to be configured, show splash screen
  if (!config || !config.saved) {
    return buildOdoWelcomeCard();
  } else if (!config.welcomeSplashShown) {
    return buildCustomerToolWelcomeSplash(context);
  }

  console.log(
    'buildIntegrationCard: context = ' +
      context +
      ', ' +
      'integrationType = ' +
      config.integrationType
  );

  let contextHandlers =
    INTEGRATION_HOOKS[config.integrationType].contextSpecificHandlers;

  if (contextHandlers.hasOwnProperty(context)) {
    return contextHandlers[context](context);
  } else if (contextHandlers.hasOwnProperty(CALL_CONTEXT.DEFAULT)) {
    return contextHandlers[CALL_CONTEXT.DEFAULT](context);
  } else {
    return buildContextNotSupportedCard(context);
  }
}

/**
 * Builds and returns a placeholder card to indicate that a particular
 * implementation has not yet been implemented.
 * 
 * @return {CardService.Card}
 */
function buildIntegrationNotYetImplementedCard() {
  let card = CardService.newCardBuilder();

  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let section = CardService.newCardSection();
  let message = 'This integration has not yet been implemented.';
  section.addWidget(CardService.newTextParagraph().setText(message));

  card.addSection(section);

  return card;
}

/**
 * Builds and returns a placeholder card to indicate that a particular
 * implementation does not support any actions in the given context
 * 
 * @param {string} context - The calling context (CALL_CONTEXT.XXX)
 * 
 * @return {CardService.Card}
 */
function buildContextNotSupportedCard(context) {
  let card = CardService.newCardBuilder();

  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let message = 'This integration does not support '
    + `this context (${context}).`;

  let section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText(message));

  card.addSection(section);

  return card.build();
}

/**
 * Builds and returns the card to customize the integration for the 
 * particular integration type specified by the paramater integrationType.
 * As part of this, calls the 'buildConfigureIntegrationCard' function defined
 * for the integration.
 *  
 * This function is called from config.gs, but can also be called by
 * specific integrations as a means to refresh their configuration card (see
 * integrationTypeRecords.gs for an example).
 * 
 * @param {string} integrationType - The selected integration type
 * 
 * @return {CardService.Card}
 */
function buildCustomizeIntegrationCard(integrationType) {
  // call the integration specific hook (function) to
  // generate the card to configure the specific integration
  let hooks = INTEGRATION_HOOKS[integrationType];
  let card = hooks.buildConfigureIntegrationCard();

  let subTitle = `${integrationTypeToPrintableString(integrationType)}`;
  let header = CardService.newCardHeader()
    .setTitle(`Customize Integration`)
    .setSubtitle(subTitle)
    .setImageStyle(CardService.ImageStyle.SQUARE)
    .setImageUrl(ODO_ICON);

  card.setHeader(header);

  // Set up params to pass to save handler when user clicks "Done"
  let params = {
    selectedIntegrationType: integrationType,
  };

  let doneAction = CardService.newAction()
    .setFunctionName('saveOdoIntegrationCustomizationSelections')
    .setParameters(params);

  let footer = CardService.newFixedFooter();
  footer.setPrimaryButton(
    CardService.newTextButton().setText('←  Done').setOnClickAction(doneAction)
  );
  card.setFixedFooter(footer);

  return card.build();
}

/**
 * Saves the integration specific configurations when user clicks the 
 * '←  Done' button in the configuration card. Does this by calling the
 * 'saveConfigureIntegrationSelections' function defined for the selected
 * integration type.
 * 
 * Note: The data returned from 'saveConfigureIntegrationsSelections' is
 * saved in the property PROP_STRINGIFIED_INTEGRATION_CUSTOMIZATIONS such that
 * it can be written to the 'integrationData' section of the config if/when the
 * user clicks 'Save' in the main configuration card.
 * 
 * Returns ActionResponse that causes card stack to pop back to main
 * configuration card.
 * 
 * @param {Object} event
 * 
 * @return {CardService.ActionResponse}
 */
function saveOdoIntegrationCustomizationSelections(event) {
  let selectedIntegrationType = event.parameters.selectedIntegrationType;
  let formInputs = event.commonEventObject.formInputs;

  // pop this integration customization card to reveal the
  // main Odo configuration card.
  let action = CardService.newActionResponseBuilder()
    .setNavigation(
      CardService.newNavigation().popToNamedCard('mainConfigurationCard')
    )
    .setStateChanged(false)
    .build();

  // call the integration specific handler to save the selections
  let hooks = INTEGRATION_HOOKS[selectedIntegrationType];

  if (!hooks.saveConfigureIntegrationSelections) {
    return action;
  }

  let integrationData = hooks.saveConfigureIntegrationSelections(formInputs);

  // save the configuration settings the user has selected when they click
  // 'Done'. These will be saved as the integrationData portion of the 'config'
  // structure if/when the user clicks 'Save' on the main configuraion card.
  let up = PropertiesService.getUserProperties();
  let integrationDataStr = JSON.stringify(integrationData);

  up.setProperty(
    PROP_STRINGIFIED_INTEGRATION_CUSTOMIZATIONS,
    integrationDataStr
  );

  return action;
}

/**
 * Builds and returns a card that shows a welcome splash message
 * on first use of the integration after it's been configured.
 * 
 * @param {String} context
 */
function buildCustomerToolWelcomeSplash(context) {
  let config = getConfig();
  let card = CardService.newCardBuilder();

  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let section = CardService.newCardSection();

  section.addWidget(CardService.newImage().setImageUrl(config.customerLogoUrl));

  let welcomeSplashMessage = findAndReplaceMergeKeys(
    config.welcomeSplashMessage
  );
  section.addWidget(
    CardService.newTextParagraph().setText(welcomeSplashMessage)
  );

  let params = {
    context: context,
  };
  let action = CardService.newAction()
    .setFunctionName('_refreshIntegrationCard')
    .setParameters(params);

  section.addWidget(
    CardService.newTextButton()
      .setText('Get Started')
      .setOnClickAction(action)
  );

  card.addSection(section);

  return card.build();
}

/**
 * Returns the integrationData stored in the configuration for the
 * given integration type passed. If the specific integration hasn't been
 * configured yet, then returns the defaults.
 * 
 * @param {String} integrationType
 * 
 * @return {Object}
 */
function getConfigIntegrationData(integrationType) {
  let integrationData;

  let config = getConfig();
  let hooks = INTEGRATION_HOOKS[integrationType];

  // if there is previously stored data for this integration type (i.e. user
  // clicked "Save" on "Configure Odo" card), then display it. 
  // else, show the default values.
  if (
    config.saved &&
    config.integrationType === integrationType
  ) {
    integrationData = config.integrationData;
  } else {
    integrationData = hooks.defaultIntegrationConfig();
  }

  return integrationData;
}

/**
 * Returns an ActionResponse that causes the current card to be updated
 * with the integration card for the specified context. The context is passed
 * in via an event parameter.
 * 
 * @param {Object} event
 * 
 * @return {CardService.ActionResponse}
 */
function _refreshIntegrationCard(event) {
  let context = event.parameters.context;
  let config = getConfig();

  config.welcomeSplashShown = true;
  saveConfig(config);

  let card = buildIntegrationCard(context);

  let action = CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card))
    .setStateChanged(false)
    .build();

  return action;
}
