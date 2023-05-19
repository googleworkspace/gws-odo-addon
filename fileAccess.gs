/**
 * Code related to obtaining permission to access Editor file with
 * drive.file scope.
 */

/**
 * Constructs card to request user grant Add-on to editor file that is open.
 *
 * @return {CardService.Card}
 */
function buildFilePermissionCard() {
  // If the add-on does not have access permission, add a button that
  // allows the user to provide that permission on a per-file basis.
  var card = CardService.newCardBuilder();
  let cardSection = CardService.newCardSection();

  cardSection.addWidget(
    CardService.newTextParagraph().setText(
      "The Add-on needs permission to access this file's contents."
    )
  );
  let buttonAction = CardService.newAction().setFunctionName(
    'onRequestFileScopeButtonClicked'
  );
  let button = CardService.newTextButton()
    .setText('Grant permission')
    .setOnClickAction(buttonAction);

  cardSection.addWidget(button);
  return card.addSection(cardSection).build();
}

/**
 * Callback function for a button action. Instructs Docs to display a
 * permissions dialog to the user, requesting `drive.file` scope for the
 * current file on behalf of this add-on.
 *
 * @param {Object} e The parameters object that contains the document’s ID
 * @return {editorFileScopeActionResponse}
 */
function onRequestFileScopeButtonClicked(e) {
  return CardService.newEditorFileScopeActionResponseBuilder()
    .requestFileScopeForActiveDocument()
    .build();
}
