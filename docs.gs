/**
 * @fileoverview Functions specific to how the Add-on looks and acts when
 * in the context of the Docs editor.
 */

/**
 * Function that's called (per manifest) for Docs homepage trigger
 *
 * return {CardService.Card} Card to show for Docs Homepage
 */
function onDocsHomepage(e) {
  if (e.docs.addonHasFileScopePermission) {
    return buildDocsKeyOnSelectedTextCard();
  }

  return buildFilePermissionCard();
}

/**
 * Function that's called (per manifest) when file access is granted
 * with drive.file scope.
 *
 * return {CardService.Card} Card to show on success
 */
function onDocsFileScopeGranted(e) {
  return buildDocsKeyOnSelectedTextCard();
}

/**
 * Build card that shows related info for user selected text in document.
 *
 * @return {CardService.Card}
 */
function buildDocsKeyOnSelectedTextCard() {
  let selectedText = getSelectedText();

  if (selectedText) {
    addMergeKeyValuePair('{{selectedText}}', selectedText);
    addMergeKeyValuePair('{{token}}', selectedText);
  }

  return buildIntegrationCard(CALL_CONTEXT.DOCS);
}

/**
 * Return currently selected text (within a paragraph).
 * https://stackoverflow.com/questions/16639331/get-user-selected-text
 *
 * @return {string} First word of selected text or empty string if no selection.
 */
function getSelectedText() {
  let doc = DocumentApp.getActiveDocument();
  let selection = doc.getSelection();

  let selectedText = '';

  if (!selection) {
    return '';
  }

  let elements = selection.getSelectedElements();
  let element = elements[0].getElement(); // first element (before '\n')
  let startOffset = elements[0].getStartOffset(); // -1 if whole element
  let endOffset = elements[0].getEndOffsetInclusive(); // -1 if whole element

  selectedText = element.asText().getText(); // All text from element

  // is only part of the element selected?
  if (elements[0].isPartial()) {
    selectedText = selectedText.substring(startOffset, endOffset + 1);
  }

  // Google Doc UI "word selection" (double click) selects trailing
  // spaces - trim them
  selectedText = selectedText.trim();

  // limit to just first word. multi-word selections won't make a lot of
  // sense with respect to simulated integrations like customer records.
  selectedText = selectedText.split(' ')[0];

  return selectedText;
}
