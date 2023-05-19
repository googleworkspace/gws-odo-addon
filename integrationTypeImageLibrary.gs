/**
 * @fileoverview Code related to information that the Odo Add-on
 * may show in various contexts if the chosen integration type is
 * INTEGRATION_TYPE.IMG_LIBRARY
 */


const PROP_SELECTED_IMG_LIBRARY_REPO = 'PROP_SELECTED_IMG_LIBRARY_REPO';

/**
 * Function that returns default configuration fields and values for
 * an integration type of INTEGRATION_TYPE.IMG_LIBRARY, to be stored
 * as the 'integrationData' field in 'config'. Called when setting
 * up default configuration.
 *
 * @return {Object}
 */
function imgLibraryGetDefaultConfig() {
  let integrationData = {
    repos: [
      {
        name: 'Finance',
        driveFolderUrl: 'https://drive.google.com/drive/folders/'
          + '1XmcEiWzyiYNIOQuMr0WuuqLonDXeL6mb'
      },
      {
        name: 'HR',
        driveFolderUrl: 'https://drive.google.com/drive/folders/'
          + '1g4VlalVo3W32wSeYp5u0-3Jy1WOp5dPe'
      },
      {
        name: '',
        driveFolderUrl: ''
      },    
    ]
  };
  
  return integrationData;
}

/**
 * Creates and returns the card that gives the user options to configure
 * the Image Library integration. Called from integrationTypeAll.gs based
 * on the value of the 'buildConfigureIntegrationCard' parameter.
 * 
 * @return {CardService.Card}
 */
function buildImgLibraryConfigureCard() {
  let config = getConfig();
  let integrationData;

  integrationData = getConfigIntegrationData(INTEGRATION_TYPE.IMAGE_LIBRARY);
 
  let repos = integrationData.repos;
  let repoName1 = integrationData.repos[0].name;
  let driveFolderUrl1 = repos[0].driveFolderUrl;

  let repoName2 = integrationData.repos[1].name;
  let driveFolderUrl2 = repos[1].driveFolderUrl;

  let repoName3 = integrationData.repos[2].name;
  let driveFolderUrl3 = repos[2].driveFolderUrl;

  let card = CardService.newCardBuilder();
  let section = CardService.newCardSection();
  let input;

  section.addWidget(CardService.newTextParagraph().setText('Repo #1:'));

  input = CardService.newTextInput()
    .setFieldName('repoName1')
    .setTitle('Repo Name')
    .setValue(repoName1);
  section.addWidget(input);

  input = CardService.newTextInput()
    .setFieldName('driveFolderUrl1')
    .setTitle('Drive Folder Url')
    .setValue(driveFolderUrl1);
  section.addWidget(input);

  section.addWidget(
    CardService.newTextParagraph().setText('<br><br>Repo #2:')
  );

  input = CardService.newTextInput()
    .setFieldName('repoName2')
    .setTitle('Repo Name')
    .setValue(repoName2);
  section.addWidget(input);

  input = CardService.newTextInput()
    .setFieldName('driveFolderUrl2')
    .setTitle('Drive Folder Url')
    .setValue(driveFolderUrl2);
  section.addWidget(input);

  section.addWidget(
    CardService.newTextParagraph().setText('<br><br>Repo #3:')
  );

  input = CardService.newTextInput()
    .setFieldName('repoName3')
    .setTitle('Repo Name')
    .setValue(repoName3);
  section.addWidget(input);

  input = CardService.newTextInput()
    .setFieldName('driveFolderUrl3')
    .setTitle('Drive Folder Url')
    .setValue(driveFolderUrl3);
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
function saveImgLibraryConfigureSelections(formInputs) {  
  let repoName1 = '';
  let driveFolderUrl1 = '';
  let repoName2 = '';
  let driveFolderUrl2 = ''; 
  let repoName3 = '';
  let driveFolderUrl3 = '';
  
  if (formInputs.hasOwnProperty('repoName1')
        && formInputs.hasOwnProperty('driveFolderUrl1')) {
    repoName1 = formInputs['repoName1'].stringInputs.value[0];
    driveFolderUrl1 = formInputs['driveFolderUrl1'].stringInputs.value[0];
  }

  if (formInputs.hasOwnProperty('repoName2')
        && formInputs.hasOwnProperty('driveFolderUrl2')) {
    repoName2 = formInputs['repoName2'].stringInputs.value[0];
    driveFolderUrl2 = formInputs['driveFolderUrl2'].stringInputs.value[0];
  }

  if (formInputs.hasOwnProperty('repoName3')
        && formInputs.hasOwnProperty('driveFolderUrl3')) {
    repoName3 = formInputs['repoName3'].stringInputs.value[0];
    driveFolderUrl3 = formInputs['driveFolderUrl3'].stringInputs.value[0];
  }

  console.log(repoName1);

  let integrationData = {
    repos: [
      {
        name: repoName1,
        driveFolderUrl: driveFolderUrl1
      },
      {
        name: repoName2,
        driveFolderUrl: driveFolderUrl2
      },
      {
        name: repoName3,
        driveFolderUrl: driveFolderUrl3
      },
    ]
  };

  return integrationData;
}


/**
 * Function used to return the image library data as a formatted Card to be
 * displayed. Called from integrationTypeAll.gs as a context specific
 * handler for this integration.
 *
 * @param {string} Calling context (i.e. CALL_CONTEXT.SLIDES)
 *
 * @return {Card}
 */
function buildImgLibraryCard(context) {
  let up = PropertiesService.getUserProperties();
  let config = getConfig();

  let card = CardService.newCardBuilder();
  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let repos = config.integrationData.repos;
  if (repos.length === 0) {
    let section = CardService.newCardSection();
    section.addWidget(CardService.newTextParagraph()
      .setText("No image library repositories have been configured"));
    card.addSection(section);
    return card.build;
  }

  // gather info on the currently selected (or default) image repo  
  let selectedRepoUrl = up.getProperty(PROP_SELECTED_IMG_LIBRARY_REPO);
  let resourceKey;

  if (!selectedRepoUrl) {
    selectedRepoId = _extractFolderIdFromUrl(repos[0].driveFolderUrl);
    resourceKey = _extractResourceKeyFromUrl(repos[0].driveFolderUrl);
    name = repos[0].name;
  } else {
    // get all the info on the previously selected repo
    for (let i=0; i < repos.length; i++) {
      if (repos[i].driveFolderUrl 
            && repos[i].driveFolderUrl === selectedRepoUrl) {
        selectedRepoId = _extractFolderIdFromUrl(repos[i].driveFolderUrl);
        resourceKey = _extractResourceKeyFromUrl(repos[i].driveFolderUrl);
        name = repos[i].name;
        break;
      }
    }
  }

  // show repo drop-down selector
  let repoSection = CardService.newCardSection();
  
  let onChangeAction = CardService.newAction()
    .setFunctionName('_refreshImageRepoCard');

  let selectRepoWidget = CardService.newSelectionInput()
    .setFieldName('selectedRepo')
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle('Image Repository')
    .setOnChangeAction(onChangeAction);

  for (let i=0; i < repos.length; i++) {
    if (repos[i].driveFolderUrl) {
      selectRepoWidget.addItem(repos[i].name, repos[i].driveFolderUrl,
        (repos[i].driveFolderUrl === selectedRepoUrl));
    }
  }

  repoSection.addWidget(selectRepoWidget);


  // show images in repo folder
  let imgSection = CardService.newCardSection();
  imgSection.addWidget(CardService.newTextParagraph()
    .setText("Click on an image to insert it:"));

  let repoFolder;
  if (resourceKey) {
    repoFolder = DriveApp.getFolderByIdAndResourceKey(
      selectedRepoId,
      resourceKey);
  } else {
    repoFolder = DriveApp.getFolderById(selectedRepoId);
  }

  let imgFileIter = repoFolder.getFiles();
  while (imgFileIter.hasNext()) {
    let imgFile = imgFileIter.next();

    let imgFileId = imgFile.getId();
    let imgFileResourceKey = imgFile.getResourceKey();
    if (imgFileResourceKey === null) {
      // can't pass a null parameter
      imgFileResourceKey = '';
    }
    let imgUrl = 'https://docs.google.com/uc?id=' + imgFileId;
    if (resourceKey) {
      imgUrl += '&resourcekey=' + resourceKey;
    }

    let params = {
      imgFileId: imgFileId,
      imgFileResourceKey: imgFileResourceKey,
    };

    let clickAction = CardService.newAction()
      .setParameters(params)
      .setFunctionName('_onImageClick');

    let imgWidget = CardService.newImage()
      .setImageUrl(imgUrl)
      .setOnClickAction(clickAction);

    imgSection.addWidget(imgWidget);
  }

  card.addSection(repoSection);
  card.addSection(imgSection);
  
  return card.build();
}

/**
 * Internal function to fresh the images when a new repository is selected.
 * 
 */
function _refreshImageRepoCard(event) {
  let formInputs = event.commonEventObject.formInputs;

  let up = PropertiesService.getUserProperties();

  let selectedRepo = formInputs.selectedRepo.stringInputs.value[0];

  up.setProperty(PROP_SELECTED_IMG_LIBRARY_REPO, selectedRepo);;

  return buildImgLibraryCard(CALL_CONTEXT.SLIDES);
}

/**
 * Internal function to insert the selected image into the selected/active
 * slide.
 * 
 * @param {Object} event - Includes parameters to identify the selected image.
 */
function _onImageClick(event) {
  let imgFileId = event.parameters.imgFileId;
  let imgFileResourceKey = event.parameters.imgFileResourceKey;

  let slide = SlidesApp.getActivePresentation()
    .getSelection()
    .getCurrentPage();

  let file;
  if (imgFileResourceKey) {
    file = DriveApp.getFileByIdAndResourceKey(imgFileId, imgFileResourceKey);
  } else {
    file = DriveApp.getFileById(imgFileId);
  }

  slide.insertImage(file.getBlob());
}

/**
 * Internal function to extract the Drive Folder ID from a folder's URL.
 * If not present (not all folders will have one), returns empty string.
 * 
 * @param {String} folderUrl
 *
 * @return {String}
 */
function _extractFolderIdFromUrl(folderUrl) {
  let folderId = '';

  let regex = /https:\/\/drive\.google\.com\/drive\/folders\/(.+)/;

  // get rid of any params (i.e. '?resourcekey=')
  folderUrlSplit = folderUrl.split('?')[0];
  let found = folderUrlSplit.match(regex);
  
  if (!found) {
    return '';
  }

  folderId = found[1];

  return folderId;
}

/**
 * Internal function to extract the Drive Folder Resource Key from a folder's
 * URL. If not present (not all folders will have one), returns empty string.
 * 
 * @param {String} folderUrl
 *
 * @return {String}
 */
function _extractResourceKeyFromUrl(folderUrl) {
  let resourceKey = '';

  let regex = /resourcekey=([^&]+)/;

  // get rid of any params (i.e. '?resourcekey=')
  let folderUrlSplit = folderUrl.split('?');

  if (folderUrlSplit.length < 2) {
    return '';
  }

  let splitPart = folderUrlSplit[1];

  let found = splitPart.match(regex);
  
  if (!found) {
    return '';
  }

  resourceKey = found[1];

  return resourceKey;
}