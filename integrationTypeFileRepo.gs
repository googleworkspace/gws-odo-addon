/**
 * @fileoverview Code related to information that the Odo Add-on
 * may show in various contexts if the chosen integration type is
 * INTEGRATION_TYPE.FILE_REPOSITORY
 */

/**
 * Function that returns default configuration fields and values for
 * an integration type of INTEGRATION_TYPE.FILE_REPOSITORY, to be stored
 * as the 'integrationData' field in 'config'. Called when setting
 * up default configuration.
 *
 * @return {Object}
 */
function fileRepoGetDefaultConfig() {
  let integrationData = {
    maxFiles: 5,
  };
  
  return integrationData;
}

/**
 * Creates and returns the card that gives the user options to configure
 * the File Repository integration. Called from integrationTypeAll.gs based
 * on the value of the 'buildConfigureIntegrationCard' parameter.
 * 
 * @return {CardService.Card}
 */
function buildFileRepoConfigureCard() {
  let config = getConfig();

  const fileLimits = [5, 10, 25, 50];

  let card = CardService.newCardBuilder();

  integrationData = getConfigIntegrationData(INTEGRATION_TYPE.FILE_REPOSITORY);
  let selectedMaxFiles = integrationData.maxFiles;
  /*
  if (config.saved &&
      config.integrationType === INTEGRATION_TYPE.FILE_REPOSITORY
    ) {
      selectedMaxFiles = config.integrationData.maxFiles;
    } else {
      selectedMaxFiles = fileLimits[0]; // default
  }
  */

  let section = CardService.newCardSection();

  let selectMaxFilesWidget = CardService.newSelectionInput()
    .setFieldName('maxFiles')
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Max Files Shown') 
  for (let i=0; i < fileLimits.length; i++) {
    let limit = fileLimits[i];
    selectMaxFilesWidget.addItem(limit.toString(), limit,
      (limit === Number(selectedMaxFiles)));
  }

  section.addWidget(selectMaxFilesWidget);

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
function saveFileRepoConfigureSelections(formInputs) {
  let integrationData = {};

  integrationData.maxFiles = formInputs.maxFiles.stringInputs.value[0];
  
  return integrationData;
}


/**
 * Function used to return the file repository data as a formatted Card to be
 * displayed. Called from integrationTypeAll.gs as a context specific
 * handler for this integration.
 *
 * @param {string} Calling context (i.e. CALL_CONTEXT.DRIVE)
 *
 * @return {Card}
 */
function buildFileRepoCard(context) {

  // if file selected, giver user chance to check it in
  let fileName = getMergeKeyValue('{{fileName}}');
  let mimeType = getMergeKeyValue('{{fileMimeType}}');
  let fileId = getMergeKeyValue('{{fileId}}');

  let card = CardService.newCardBuilder();
  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  if (fileName) {
    // file is selected
    let section = CardService.newCardSection();

    // check mimeType to ensure it's DOCX
    if (!_isOfficeFile(mimeType)) {
      console.log(mimeType)
      let message = 'Only MS Office documents can be checked-in.';
      section.addWidget(CardService.newTextParagraph().setText(message));
      // inform user that only office files are allowed
    } else {
      // give user chance to check in file
      let url = _getOfficeFileIconUrl(mimeType)
      let icon = CardService.newIconImage().setIconUrl(url);
      let textField = CardService.newDecoratedText()
        .setStartIcon(icon)
        .setText(fileName);

      section.addWidget(textField);

      let params = {
        fileId: fileId,
        fileName: fileName,
      };

      let buttonAction = CardService.newAction()
        .setFunctionName('_checkInFile')
        .setParameters(params);

      let button = CardService.newTextButton()
        .setText("Check-in File")
        .setOnClickAction(buttonAction);

      section.addWidget(button);

    }
    card.addSection(section);

  } else {
    // no file is selected
    let section = CardService.newCardSection();

    section.addWidget(CardService.newImage().setImageUrl(SELECT_FILE_ICON));

    section.addWidget(
      CardService.newTextParagraph().setText(
        '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' +
        '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' +
        'Select a file to check it in'
    ));

    card.addSection(section);

    // show user list of files they can "check out"
    let ciFiles = _getCheckedInFiles();

    if (ciFiles.length) {
      let section2 = CardService.newCardSection();

      let message = '<br>Or select a file listed below to check-out:';
      section2.addWidget(CardService.newTextParagraph().setText(message));

      for (let i=0; i < ciFiles.length; i++) {
        let file = ciFiles[i];

        let fileListEntry = CardService.newDecoratedText();
        fileListEntry.setText(file.getName())
          .setBottomLabel('Modified: ' + file.getLastUpdated())
        let iconImage = CardService.newIconImage()
          .setIconUrl(_getOfficeFileIconUrl(file.getMimeType()));
        fileListEntry.setStartIcon(iconImage);
        let action = CardService.newAction()
          .setFunctionName('_checkOutFile')
          .setParameters({fileId: file.getId(), fileName: file.getName()});

        let button = CardService.newImageButton()
          .setIconUrl(CHECKOUT_FILE_ICON)
          .setAltText('Checkout File')
          .setOnClickAction(action);
        fileListEntry.setButton(button);

        section2.addWidget(fileListEntry);
      }

      card.addSection(section2);
    }
  }
  
  return card.build();
}

/**
 * Private function that handles the user's request to check out
 * a selected file when the button next to it is clicked.
 * @param {Object} event
 */
function _checkOutFile(event) {
  let up = PropertiesService.getUserProperties();
  let fileId = event.parameters.fileId;
  let fileName = event.parameters.fileName;
  let destFolder;

  let destFolderId = up.getProperty(fileId);

  if (!destFolderId) {
    destFolder = DriveApp.getRootFolder();    
  } else {
    destFolder = DriveApp.getFolderById(destFolderId);
  }
  
  let f = DriveApp.getFileById(fileId);
  f.moveTo(destFolder);

  let card = CardService.newCardBuilder();
  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let section = CardService.newCardSection();

  let message = `File ${fileName}"" has been checked-out`;

  let url = f.getUrl();
  section.addWidget(CardService.newTextParagraph().setText(message));

  card.addSection(section);
  let updateCard = card.build();

  let actionResponse = CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(updateCard))
    .setStateChanged(true)
    .build();

  return actionResponse;
}

/** 
 * Private function that handles the user's request to check in
 * a selected file when the "Check-In File" button is clicked.
 * 
 * @param {Object} event
 * 
 */
function _checkInFile(event) {
  let up = PropertiesService.getUserProperties();

  let fileId = event.parameters.fileId;
  let fileName = event.parameters.fileName;

  let f = DriveApp.getFileById(fileId);
  let parentFolderIter = f.getParents();
  
  if (parentFolderIter.hasNext()) {
    let parentFolderId = parentFolderIter.next().getId();
    // record where this file lived so we can put it back there
    // upon next check-in
    up.setProperty(fileId, parentFolderId);
  }

  // "check-in" the file. Here we simulate this by just moving
  // the file to a special folder so we can hide it from view.
  let odf = createOrGetOdoDataFolder();
  let file = DriveApp.getFileById(fileId);
  file.moveTo(odf);

  let card = CardService.newCardBuilder();

  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let section = CardService.newCardSection();

  let message = `File "${fileName}" has been checked-in`;
  section.addWidget(CardService.newTextParagraph().setText(message));

  card.addSection(section);
  let updateCard = card.build();

  let actionResponse = CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(updateCard))
    .setStateChanged(true)
    .build();

  return actionResponse;
}

/**
 * Private function that returns a list of Files that have been checked in
 * (sorted by last modified date). Number returned is limited to 
 * config.integrationData.maxFiles
 * 
 * @return {Object} Array of Files
 */
function _getCheckedInFiles() {
  let config = getConfig();

  let ciFiles = [];
  let odf = createOrGetOdoDataFolder();
  let count = 0;
  let filesIter = odf.getFiles();

  while (filesIter.hasNext()) {
    let file = filesIter.next();
    ciFiles.push(file);
    count++;
    if (count === config.integrationData.maxFiles) {
      break;
    }
  }

  return ciFiles;
}

/**
 * Private function that returns true if given mimeType corresponds to
 * an MS Office file (Word, Excel, Powerpoint), false otherwise.
 * 
 * @return {Boolean} 
 */
function _isOfficeFile(mimeType) {
  if (mimeType.includes('officedocument')) {
    return true;
  }

  return false;
}

/**
 * Private function that returns a URL for an icon for Word, Excel, or 
 * Powerpoint files (based on the passed mimeType). Call _isOfficeFile
 * first to verify if mimeType corresponds to Office file.
 * 
 * @return {String} 
 */
function _getOfficeFileIconUrl(mimeType) {
  let url;

  if (mimeType.includes('word')) {
    url = 'https://img.icons8.com/color/512/ms-word.png';  
  } else if (mimeType.includes('presentation')) {
    url = 'https://img.icons8.com/color/512/ms-powerpoint.png';
  } else if (mimeType.includes('spreadsheet')) {
    url = 'https://img.icons8.com/color/512/ms-excel.png';
  } else {
    url = 'https://img.icons8.com/color/512/office-365.png';
  }

  return url;
}


