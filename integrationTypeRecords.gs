/**
 * @fileoverview Code related to "records" that the Odo Add-on
 * may show in various contexts if the chosen integration type is
 * INTEGRATION_TYPE.RECORDS_BASED.
 */

/**
 * Types of Records that can be operated on (for RECORDS_BASED integration
 * types).
 * @enum {string}
 */
const RECORD_TYPE = {
  CUSTOMER: 'CUSTOMER',
  EMPLOYEE: 'EMPLOYEE',
  ASSET: 'ASSET',
  GENERIC: 'CUSTOM',
};

/**
 * Allowed data types for Record fields. These tell Odo how to format and
 * display the record fields.
 * @enum {string}
 */
const RECORD_FIELD_TYPE = {
  TEXT: 'TEXT', // just generic text
  EMAIL: 'EMAIL', // an email address
  PERSON_ID: 'PERSON_ID', // some form of ID for an individual (i.e. name)
  ORG_ID: 'ORD_ID', // some form of ID for an org (i.e. company name)
  DATE_EPOCH_MS: 'DATE_EPOCH_MS', // an absolute date, in ms since epoch
  DATE_OFFSET_DAYS: 'DATE_OFFSET_DAYS', // offset from present date in days
  FILE_URL: 'FILE_URL', // link to a file (i.e. pdf stored in Drive),
  FOLDER_URL: 'FOLDER_URL', // link to a folder (i.e. Drive folder)
};

const PROP_SELECTED_RECORD_TYPE = 'PROP_SELECTED_RECORD_TYPE';

/**
 * Function that returns default configuration fields and values for
 * an integration type of INTEGRATION_TYPE.RECORDS_BASED, to be stored
 * as the 'integrationData' field in 'config'. Called when setting
 * up default configuration.
 *
 * @return {Object}
 */
function recordGetDefaultConfig() {
  let defaultRecordType = RECORD_TYPE.CUSTOMER;

  return {
    // type of record
    type: defaultRecordType,

    // data associated with specific record (varies by record type).
    recordFields: _recordGetDefaultFields(defaultRecordType),
  };
}

/**
 * Function used to return the default record data based on recordType.
 *
 * @return {Object}
 */
function _recordGetDefaultFields(recordType) {
  let recordFields;

  switch (recordType) {
    case RECORD_TYPE.ASSET:
      recordFields = [
        ['Asset ID', '{{selectedText || CG51112}}', RECORD_FIELD_TYPE.TEXT],
        ['Asset Type', 'Laptop', RECORD_FIELD_TYPE.TEXT],
        ['Description', 'Chromebook', RECORD_FIELD_TYPE.TEXT],
        [
          'Allocated To',
          '{{senderEmail || road.runner@cymbal.dev}}',
          RECORD_FIELD_TYPE.EMAIL,
        ],
        ['Asset File', DEFAULT_RECORD_FILE, RECORD_FIELD_TYPE.FILE_URL],
      ];
      break;

    case RECORD_TYPE.EMPLOYEE:
      recordFields = [
        [
          'Name',
          '{{senderName || Wile E. Coyote}}',
          RECORD_FIELD_TYPE.PERSON_ID,
        ],
        [
          'Email',
          '{{senderEmail || wile.e.coyote@acme.com}}',
          RECORD_FIELD_TYPE.EMAIL,
        ],
        ['Manager Email', 'road.runner@cymbal.dev', RECORD_FIELD_TYPE.EMAIL],
        ['Employee ID', '{{selectedText || EID54521}}', RECORD_FIELD_TYPE.TEXT],
        ['Company Start Date', 1641816000000, RECORD_FIELD_TYPE.DATE_EPOCH_MS],
        ['Contract File', DEFAULT_RECORD_FILE, RECORD_FIELD_TYPE.FILE_URL],
      ];
      break;

    case RECORD_TYPE.CUSTOMER:
    default:
      recordFields = [
        ['Customer Name', 'ACME Corp', RECORD_FIELD_TYPE.ORG_ID],
        ['Customer ID', '{{selectedText || C121GW}}', RECORD_FIELD_TYPE.TEXT],
        [
          'Contact Name',
          '{{senderName || Wile E. Coyote}}',
          RECORD_FIELD_TYPE.PERSON_ID,
        ],
        [
          'Contact Email',
          '{{senderEmail || wile.e.coyote@acme.com}}',
          RECORD_FIELD_TYPE.EMAIL,
        ],
        ['Last Emailed Date', -1, RECORD_FIELD_TYPE.DATE_OFFSET_DAYS],
        ['Last Email Subject', 'Order for Bird Trap', RECORD_FIELD_TYPE.TEXT],
        ['Contract File', DEFAULT_RECORD_FILE, RECORD_FIELD_TYPE.FILE_URL],
      ];
      break;
  }

  return recordFields;
}

/**
 * Function used to return the record data as a formatted Card to be
 * displayed. Called from integrationTypeAll.gs as a context specific
 * handler for this integration.
 *
 * @param {string} Calling context (i.e. CALL_CONTEXT.GMAIL_VIEW)
 *
 * @return {CardService.Card}
 */
function buildRecordCard(context) {
  let config = getConfig();

  let senderEmail = undefined;
  let selectedText = '';

  let kvPairs = getAllMergeKeyPairs();

  if (kvPairs.hasOwnProperty('{{senderEmail}}')) {
    senderEmail = kvPairs['{{senderEmail}}']; // only def for Gmail contexts
  }
  if (kvPairs.hasOwnProperty('{{selectedText}}')) {
    selectedText = kvPairs['{{selectedText}}']; // only def for Editors
  }

  if (!selectedText) {
    if (context === CALL_CONTEXT.DOCS) {
      return buildNoSelectionRecordCard(context);
    } else if (context === CALL_CONTEXT.SHEETS) {
      return buildNoSelectionRecordCard(context);
    }
  } else {
    if (context === CALL_CONTEXT.SHEETS && isFreshSheetLoad()) {
      return buildNoSelectionRecordCard(context);
    }
  }

  let recordType = config.integrationData.type;
  let recordFields = config.integrationData.recordFields;

  let card = CardService.newCardBuilder();

  let recordTypeString = recordTypeToPrintableString(recordType);

  let titleText;
  if (context === CALL_CONTEXT.GMAIL_COMPOSE) {
    titleText = 'Attach File From ' + recordTypeString + ' Record';
  } else {
    titleText = 'View ' + recordTypeString + ' Record';
  }

  card.setHeader(
    CardService.newCardHeader().setImageUrl(RECORD_ICON).setTitle(titleText)
  );

  let peekHeader = CardService.newCardHeader()
    .setTitle(titleText)
    .setSubtitle('Click here to view last record')
    .setImageUrl(RECORD_ICON);

  card.setPeekCardHeader(peekHeader);

  let attachmentFieldName = '';
  let section = CardService.newCardSection();

  // output the different sections of the record
  for (let i = 0; i < recordFields.length; i++) {
    let recordEntry = recordFields[i];
    let fieldName = recordEntry[0];
    let fieldValue = recordEntry[1];
    let fieldType = recordEntry[2];
    let startIcon = _getStartIconForRecordFieldType(fieldType);

    // replace any merge keys
    fieldValue = findAndReplaceMergeKeys(fieldValue, { boldValue: true });

    let displayValue = fieldValue;

    if (fieldType === RECORD_FIELD_TYPE.DATE_EPOCH_MS) {
      let msSinceEpoch = fieldValue;
      let targetDate = new Date(msSinceEpoch);
      displayValue = Utilities.formatDate(targetDate, 'GMT', config.dateFormat);
    } else if (fieldType === RECORD_FIELD_TYPE.DATE_OFFSET_DAYS) {
      let today = new Date();
      let daysOffset = Number(fieldValue);

      let targetDateMS = today.getTime() + 24 * 3600000 * daysOffset;

      let targetDate = new Date(targetDateMS);
      displayValue = Utilities.formatDate(targetDate, 'GMT', config.dateFormat);
    }

    let widget = CardService.newDecoratedText()
      .setTopLabel(fieldName)
      .setStartIcon(startIcon)
      .setText(displayValue);

    if (
      fieldType === RECORD_FIELD_TYPE.FILE_URL ||
      fieldType === RECORD_FIELD_TYPE.FOLDER_URL
    ) {
      widget.setOpenLink(CardService.newOpenLink().setUrl(fieldValue));
      attachmentFieldName = fieldName;
      attachmentFileUrl = fieldValue;
    }

    section.addWidget(widget);
  }

  card.addSection(section);

  if (context === CALL_CONTEXT.GMAIL_COMPOSE) {
    let footer = CardService.newFixedFooter();

    let params = {
      fileUrl: attachmentFileUrl,
      fileName: `${attachmentFieldName} for ${senderEmail}`,
    };

    let buttonAction = CardService.newAction()
      .setFunctionName('_attachFileToGmailMessage')
      .setParameters(params);

    let buttonText = `ATTACH ${attachmentFieldName.toUpperCase()} TO EMAIL`;
    footer.setPrimaryButton(
      CardService.newTextButton()
        .setText(buttonText)
        .setOnClickAction(buttonAction)
    );

    card.setFixedFooter(footer);
  } else if (context === CALL_CONTEXT.DOCS || context === CALL_CONTEXT.SHEETS) {
    let sections = buildRecordActionSections(context);
    for (var i = 0; i < sections.length; i++) {
      card.addSection(sections[i]);
    }
  }

  return card.build();
}

/**
 * Returns card to show in an editors context when no document text
 * has yet been selected.
 *
 * @return {CardService.Card}
 */
function buildNoSelectionRecordCard(context) {
  let card = CardService.newCardBuilder();

  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let opts = {};

  if (context === CALL_CONTEXT.SHEETS) {
    opts.showGenerateReport = true;
  }

  let sections = buildRecordActionSections(context, opts);
  for (var i = 0; i < sections.length; i++) {
    card.addSection(sections[i]);
  }

  return card.build();
}

/**
 * Returns one or more sections to be displayed that contain instructions
 * and/or buttons for the user to take action with (i.e. look up a record)
 * @return {[CardService.CardSection]}
 */
function buildRecordActionSections(context, opts) {
  let section = CardService.newCardSection();
  let section2 = CardService.newCardSection();

  let selectInstructionText;

  if (context === CALL_CONTEXT.DOCS) {
    selectInstructionText = 'select text in this document';
  } else if (context === CALL_CONTEXT.SHEETS) {
    selectInstructionText = 'select a cell in this sheet';
  }

  let instructions =
    'To lookup a record, ' +
    `${selectInstructionText}` +
    ' related to a record and click the button below.';

  section.addWidget(CardService.newTextParagraph().setText(instructions));
  let buttonAction = CardService.newAction();

  if (context === CALL_CONTEXT.DOCS) {
    buttonAction.setFunctionName('buildDocsKeyOnSelectedTextCard');
  } else if (context === CALL_CONTEXT.SHEETS) {
    buttonAction.setFunctionName('buildSheetsKeyOnSelectedCellCard');
  }

  let buttonText = 'Lookup Record';
  let button = CardService.newTextButton()
    .setText(buttonText)
    .setOnClickAction(buttonAction);
  section.addWidget(button);

  returnSections = [section];

  if (context === CALL_CONTEXT.SHEETS) {
    if (opts && opts.showGenerateReport) {
      let instructions =
        'Or use the button below to ' + 'start a quick report:';
      section2 = CardService.newCardSection();
      section2.addWidget(CardService.newTextParagraph().setText(instructions));

      let buttonAction2 = CardService.newAction();
      buttonAction2.setFunctionName('buildSheetsReportGeneratorCard');

      let buttonText2 = 'Start Report';
      let button2 = CardService.newTextButton()
        .setText(buttonText2)
        .setOnClickAction(buttonAction2);
      section2.addWidget(button2);

      returnSections.push(section2);
    }
  }

  return returnSections;
}

/**
 * Generates a card from which a user can generate a report with data
 * written into a new sheet.
 *
 * @return {CardService.Card}
 */
function buildSheetsReportGeneratorCard() {
  let config = getConfig();
  let recordType = config.integrationData.type;
  let recordTypeString = recordTypeToPrintableString(recordType);

  let card = CardService.newCardBuilder();

  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let instructions =
    `Configure your ${recordTypeString} report using ` +
    'the options below. Once ready, ' +
    'click the button below.';

  let section = CardService.newCardSection();

  section.addWidget(CardService.newTextParagraph().setText(instructions));

  let dateWidget1 = CardService.newDatePicker()
    .setTitle('Report Start Date')
    .setFieldName('startDate');

  section.addWidget(dateWidget1);

  let dateWidget2 = CardService.newDatePicker()
    .setTitle('Report End Date')
    .setFieldName('endDate');

  section.addWidget(dateWidget2);

  let buttonAction = CardService.newAction();
  buttonAction.setFunctionName('createReportData');

  let buttonText = `Generate ${recordTypeString} Report`;
  let button = CardService.newTextButton()
    .setText(buttonText)
    .setOnClickAction(buttonAction);
  section.addWidget(button);

  card.addSection(section);

  return card.build();
}

/**
 * Writes report data into a new sheet in the active Spreadsheet. Returns an
 * ActionResponse to load the card that displays the record.
 * 
 * @return {CardService.ActionResponse}
 */
function createReportData() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let config = getConfig();
  let sheet = ss.insertSheet();

  recordFields = config.integrationData.recordFields;
  let headerRow = [];
  for (var i = 0; i < recordFields.length; i++) {
    headerRow.push(recordFields[i][0]); // field name
  }
  sheet.appendRow(headerRow);

  let numRows = 10;

  for (var i = 0; i < numRows; i++) {
    let row = [];
    for (var j = 0; j < recordFields.length; j++) {
      let type = recordFields[j][2];
      let value = recordFields[j][1];

      if (type === RECORD_FIELD_TYPE.DATE_EPOCH_MS) {
        let msSinceEpoch = fieldValue;
        let targetDate = new Date(msSinceEpoch);
        value = Utilities.formatDate(targetDate, 'GMT', config.dateFormat);
      } else if (type === RECORD_FIELD_TYPE.DATE_OFFSET_DAYS) {
        let today = new Date();
        let daysOffset = Number(value);
        let targetDateMS = today.getTime() + 24 * 3600000 * daysOffset;
        let targetDate = new Date(targetDateMS);
        value = Utilities.formatDate(targetDate, 'GMT', config.dateFormat);
      } else if (
        typeof value === 'string' &&
        value.includes('{{selectedText')
      ) {
        value = getDefaultMergeValue(value);
        value += 'R' + (i + 101);
      } else {
        value = findAndReplaceMergeKeys(value);
      }

      row.push(value);
    }

    sheet.appendRow(row);
  }

  let card = buildRecordCard(CALL_CONTEXT.SHEETS);

  let actionResponse = CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card))
    .setStateChanged(false)
    .build();

  return actionResponse;
}

/**
 * Creates and returns the card that gives the user options to configure
 * the Records integration. Called from integrationTypeAll.gs based on the
 * value of the 'buildConfigureIntegrationCard' parameter.
 *  
 * @return {CardService.Card}
 */
function buildRecordsConfigureCard() {
  let config = getConfig();
  let up = PropertiesService.getUserProperties();

  // Reads the currently selected record type from the 
  // PROP_SELECTED_RECORD_TYPE property. This is used to facilitate
  // refreshing the card when the user selectes a different record type
  // from the drop-down selector.
  let selectedRecordType = up.getProperty(PROP_SELECTED_RECORD_TYPE);

  if (!selectedRecordType) {
    if (
      config.saved &&
      config.integrationType === INTEGRATION_TYPE.RECORDS_BASED
    ) {
      selectedRecordType = config.integrationData.type;
    } else {
      selectedRecordType = RECORD_TYPE.CUSTOMER; // default
    }

    up.setProperty(PROP_SELECTED_RECORD_TYPE, selectedRecordType);
  }

  let card = CardService.newCardBuilder();

  let recordTypeSection = CardService.newCardSection();

  let params = {
    selectedIntegrationType: INTEGRATION_TYPE.RECORDS_BASED,
  };
  let onChangeAction = CardService.newAction()
    .setFunctionName('_refreshRecordsConfigureCard')
    .setParameters(params);

  let selectRecordTypeWidget = CardService.newSelectionInput()
    .setFieldName('selectedRecordType')
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Record Type')
    .setOnChangeAction(onChangeAction);

  _addRecordTypeSelectionToDropdown(
    selectedRecordType,
    selectRecordTypeWidget,
    RECORD_TYPE.CUSTOMER
  );
  _addRecordTypeSelectionToDropdown(
    selectedRecordType,
    selectRecordTypeWidget,
    RECORD_TYPE.ASSET
  );
  _addRecordTypeSelectionToDropdown(
    selectedRecordType,
    selectRecordTypeWidget,
    RECORD_TYPE.EMPLOYEE
  );

  recordTypeSection.addWidget(selectRecordTypeWidget);

  card.addSection(recordTypeSection);

  recordFieldsSection = _buildRecordFieldsSection(selectedRecordType);

  card.addSection(recordFieldsSection);

  return card;
}

/**
 * Function that causes the configuration card to be refreshed. Called when the
 * record type is selected/changes in the drop-down selector.
 * 
 * @param {Object} event
 * 
 * @return {CardService.ActionResponse}
 */
function _refreshRecordsConfigureCard(event) {
  // user selected a different record type from the drop-down.
  // store their new selection and then reload this card.
  let formInputs = event.commonEventObject.formInputs;
  let selectedIntegrationType = event.parameters.selectedIntegrationType;

  let up = PropertiesService.getUserProperties();

  let selectedRecordType = formInputs.selectedRecordType.stringInputs.value[0];

  up.setProperty(PROP_SELECTED_RECORD_TYPE, selectedRecordType);

  // DAA TODO: I don't like that here, inside the implementation of a specific
  // integration, we're reaching back out to 'buildCustomizeIntegrationCard',
  // which is part of the higher-level integration framework. It breaks the
  // clean separation that otherwise exists between specific integration
  // implentations and the main Odo framework. Perhaps think of a better way
  // to accomplish this, or at least clearly document how/why/when to use it.
  let refreshedCard = buildCustomizeIntegrationCard(selectedIntegrationType);
  let actionResponse = CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
    .setStateChanged(false)
    .build();

  return actionResponse;
}

/**
 * Function that gets called for this particular integration when user 
 * clicks '←  Done' button in integration configuration card. Saves the 
 * selections and returns them as an object to be stored in the
 * 'integrationData' field of the config object if/when the user saves their
 * configurations.
 * 
 * This is the handler that's defined as 'saveConfigureIntegrationSelections' 
 * in integrationTypeAll.gs.
 * 
 * @param {object} formInputs - Contains user selections
 * 
 * @return {object}
 */
function saveRecordsConfigureSelections(formInputs) {
  let up = PropertiesService.getUserProperties();

  let integrationData = {};

  integrationData.type = formInputs.selectedRecordType.stringInputs.value[0];

  integrationData.recordFields = [];

  for (key in formInputs) {
    if (!formInputs.hasOwnProperty(key)) {
      continue;
    }

    let fieldNameKey = key.split('fieldValueId---');
    if (fieldNameKey.length !== 2) {
      continue;
    }

    let fieldName = fieldNameKey[1];
    let typeKey = 'fieldTypeId---' + fieldName;

    let fieldValue = formInputs[key].stringInputs.value[0];

    let fieldType = formInputs[typeKey].stringInputs.value[0];

    integrationData.recordFields.push([fieldName, fieldValue, fieldType]);
  }

  // clear selected record type for next time
  up.deleteProperty(PROP_SELECTED_RECORD_TYPE);

  return integrationData;
}

/**
 * Adds the given record type to the drop-down widget passed.
 * 
 * @param {text} selectedRecordType - Current selection to show as default
 * @param {CardService.SelectionInput} selectedRecordTypeWidget
 * @param {text} recordType - Record type to add to drop-down
 */
function _addRecordTypeSelectionToDropdown(
  selectedRecordType,
  selectedRecordTypeWidget,
  recordType
) {
  let selected = false;

  if (selectedRecordType === recordType) {
    selected = true;
  }

  selectedRecordTypeWidget.addItem(
    recordTypeToPrintableString(recordType),
    recordType,
    selected
  );
}

/**
 * @param {text} selectedRecordType - Current selection for record type
 * 
 * Builds and returns a CardSection containing the configurable selections
 * for the selected record type.
 * 
 * @return {CardService.CardSection}
 */
function _buildRecordFieldsSection(selectedRecordType) {
  let config = getConfig();
  let section = CardService.newCardSection();
  let recordFields;

  // if there is previously stored data for this integration type,
  // display it. else, show the default values.
  if (
    config.saved &&
    config.integrationType === INTEGRATION_TYPE.RECORDS_BASED &&
    config.integrationData.type === selectedRecordType
  ) {
    recordFields = config.integrationData.recordFields;
  } else {
    recordFields = _recordGetDefaultFields(selectedRecordType);
  }

  for (let i = 0; i < recordFields.length; i++) {
    let fieldArray = recordFields[i];
    let fieldName = fieldArray[0];
    let fieldValue = fieldArray[1];
    let fieldType = fieldArray[2];

    let fieldID = 'fieldValueId---' + fieldName;
    let fieldValueInput = CardService.newTextInput()
      .setFieldName(fieldID)
      .setTitle(fieldName)
      .setValue(fieldValue);

    fieldID = 'fieldTypeId---' + fieldName;
    let fieldTypeSelect = CardService.newSelectionInput()
      .setFieldName(fieldID)
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setTitle('Field Type');

    _populateRecordFieldTypeSelect(fieldTypeSelect, fieldType);

    section.addWidget(fieldValueInput);
    section.addWidget(fieldTypeSelect);
  }

  return section;
}

/**
 * Populated a given drop-down widget with the possible record field types
 * (i.e. RECORD_FIELD_TYPE.XXX) that a user can select when configuring their
 * record type.
 * 
 * @param {CardService.SelectionInput} dropdownWidget
 * @param {text} selectedFieldType - current selection to show as default
 */
function _populateRecordFieldTypeSelect(dropdownWidget, selectedFieldType) {
  dropdownWidget.addItem(
    RECORD_FIELD_TYPE.TEXT,
    RECORD_FIELD_TYPE.TEXT,
    RECORD_FIELD_TYPE.TEXT === selectedFieldType
  );
  dropdownWidget.addItem(
    RECORD_FIELD_TYPE.EMAIL,
    RECORD_FIELD_TYPE.EMAIL,
    RECORD_FIELD_TYPE.EMAIL === selectedFieldType
  );
  dropdownWidget.addItem(
    RECORD_FIELD_TYPE.PERSON_ID,
    RECORD_FIELD_TYPE.PERSON_ID,
    RECORD_FIELD_TYPE.PERSON_ID === selectedFieldType
  );
  dropdownWidget.addItem(
    RECORD_FIELD_TYPE.ORG_ID,
    RECORD_FIELD_TYPE.ORG_ID,
    RECORD_FIELD_TYPE.ORG_ID === selectedFieldType
  );
  dropdownWidget.addItem(
    RECORD_FIELD_TYPE.DATE_EPOCH_MS,
    RECORD_FIELD_TYPE.DATE_EPOCH_MS,
    RECORD_FIELD_TYPE.DATE_EPOCH_MS === selectedFieldType
  );
  dropdownWidget.addItem(
    RECORD_FIELD_TYPE.DATE_OFFSET_DAYS,
    RECORD_FIELD_TYPE.DATE_OFFSET_DAYS,
    RECORD_FIELD_TYPE.DATE_OFFSET_DAYS === selectedFieldType
  );
  dropdownWidget.addItem(
    RECORD_FIELD_TYPE.FILE_URL,
    RECORD_FIELD_TYPE.FILE_URL,
    RECORD_FIELD_TYPE.FILE_URL === selectedFieldType
  );
  dropdownWidget.addItem(
    RECORD_FIELD_TYPE.FOLDER_URL,
    RECORD_FIELD_TYPE.FOLDER_URL,
    RECORD_FIELD_TYPE.FOLDER_URL === selectedFieldType
  );
}

/**
 * Private function that returns a CardService.IconImage based on field
 * type within a Record.
 *
 * @return {CardService.IconImage}
 */
function _getStartIconForRecordFieldType(fieldType) {
  let startIcon = CardService.newIconImage();
  let startIconUrl = '';

  switch (fieldType) {
    case RECORD_FIELD_TYPE.EMAIL:
      startIconUrl = EMAIL_ICON;
      break;
    case RECORD_FIELD_TYPE.PERSON_ID:
      startIconUrl = PERSON_ICON;
      break;
    case RECORD_FIELD_TYPE.ORG_ID:
      startIconUrl = ORG_ICON;
      break;
    case RECORD_FIELD_TYPE.DATE_EPOCH_MS:
    case RECORD_FIELD_TYPE.DATE_OFFSET_DAYS:
      startIconUrl = DATE_ICON;
      break;
    case RECORD_FIELD_TYPE.FILE_URL:
      startIconUrl = FILE_ICON;
      break;
    case RECORD_FIELD_TYPE.FOLDER_URL:
      startIconUrl = FOLDER_ICON;
      break;
    case RECORD_FIELD_TYPE.TEXT:
    default:
      startIconUrl = TEXT_ICON;
      break;
  }

  startIcon.setIconUrl(startIconUrl);

  return startIcon;
}

/** Function that converts a Record Type enum into a printable
 * string that can be shown to the user.
 *
 * @param {string} Record type (i.e. RECORD_TYPE.CUSTOMER)
 *
 * @return {string}
 */
function recordTypeToPrintableString(recordType) {
  let recordTypeString = '';

  switch (recordType) {
    case RECORD_TYPE.CUSTOMER:
      recordTypeString = 'Customer';
      break;
    case RECORD_TYPE.EMPLOYEE:
      recordTypeString = 'Employee';
      break;
    case RECORD_TYPE.ASSET:
      recordTypeString = 'Asset';
      break;
  }

  return recordTypeString;
}

/**
 * Action handler that inserts a file "attachment" into an email
 * when the record is being viewed in the Gmail compose context.
 * 
 * @param {Object} Handler event
 * 
 * @return {CardService.UpdateDraftActionResponse}
 */
function _attachFileToGmailMessage(event) {
  let fileUrl = event.parameters.fileUrl;
  let fileName = event.parameters.fileName;

  let textHtmlContent =
    '<span style="display:flex;align-items:center;">' +
    `<img src="${FILE_ICON}">&nbsp;` +
    `<a href="${fileUrl}">` +
    `${fileName}</a></span>`;

  let response = CardService.newUpdateDraftActionResponseBuilder()
    .setUpdateDraftBodyAction(
      CardService.newUpdateDraftBodyAction()
        .addUpdateContent(textHtmlContent, CardService.ContentType.MUTABLE_HTML)
        .setUpdateType(CardService.UpdateDraftBodyType.IN_PLACE_INSERT)
    )
    .build();

  return response;
}

