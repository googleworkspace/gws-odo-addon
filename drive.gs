/**
 * @fileoverview Functions specific to how the Add-on looks and acts when
 * in the context of a user's Drive.
 */

/**
 * Function that's called (per manifest) for Drive homepage trigger
 *
 * @param {Object} e The event object.
 *
 * @return {CardService.Card} Card to show for Drive Homepage
 */
function onDriveHomepage(e) {
  return buildIntegrationCard(CALL_CONTEXT.DRIVE);
}


/**
 * Callback for rendering the card for specific Drive items.
 * @param {Object} e - The event object.
 * 
 * @return {CardService.Card} The card to show to the user.
 */
function onDriveItemsSelected(e) {

  // Grab first file selected
  var selectedFile = e.drive.selectedItems[0];
  
  // Set up merge tags, based on selected file.
  addMergeKeyValuePair('{{fileMimeType}}', selectedFile.mimeType);
  addMergeKeyValuePair('{{fileName}}', selectedFile.title);
  addMergeKeyValuePair('{{fileId}}', selectedFile.id);

  return buildIntegrationCard(CALL_CONTEXT.DRIVE);

}

/**
 * Returns special folder for storing Odo specific data in the user's
 * Drive. Creates the folder first if not already present.
 * 
 * @return {Object} A Folder object for use with DriveApp
 */
function createOrGetOdoDataFolder() {
  let dataFolder;
  let folders = DriveApp.getFoldersByName(ODO_DATA_FOLDER_NAME);

  if (!folders.hasNext()) {
    dataFolder = DriveApp.createFolder(ODO_DATA_FOLDER_NAME);
  } else {
    dataFolder = folders.next();
  }

  return dataFolder;
}
