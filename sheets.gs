/**
 * @fileoverview Functions specific to how the Add-on looks and acts when
 * in the context of the Sheets editor.
 */

/**
 * Function that's called (per manifest) for Docs homepage trigger
 *
 * return {CardService.Card} Card to show for Docs Homepage
 */
function onSheetsHomepage(e) {
  if (e.sheets.addonHasFileScopePermission) {
    return buildSheetsKeyOnSelectedCellCard();
  }

  return buildFilePermissionCard();
}

/**
 * Function that's called (per manifest) when file access is granted
 * with drive.file scope.
 *
 * return {CardService.Card} Card to show on success
 */
function onSheetsFileScopeGranted(e) {
  return onSheetsHomepage();
}

/**
 * Build card that shows related info for user selected call in sheet.
 *
 * @return {CardService.Card}
 */
function buildSheetsKeyOnSelectedCellCard() {
  let selectedText = getSelectedCellText();

  addMergeKeyValuePair('{{selectedText}}', selectedText);
  addMergeKeyValuePair('{{token}}', selectedText);

  return buildIntegrationCard(CALL_CONTEXT.SHEETS);
}

/**
 * Returns text of currently selected cell
 *
 * @return {string} First word of text in selected cell, or empty string
 * if no selection.
 */
function getSelectedCellText() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let cell = sheet.getActiveCell();

  let selectedText = '';

  if (!cell) {
    return '';
  }

  selectedText = cell.getValue().toString();

  // Google Doc UI "word selection" (double click) selects trailing
  // spaces - trim them
  selectedText = selectedText.trim();

  // limit to just first word. multi-word selections won't make a lot of
  // sense with respect to simulated integrations like customer records.
  selectedText = selectedText.split(' ')[0];

  return selectedText;
}

/** Returns true if active cell is A1, indicating that the sheet has been
 * loaded but the user hasn't actually done much (i.e. hasn't explicitly
 * selected any cell to look up a record for)
 *
 * @return {boolean}
 */
function isFreshSheetLoad() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let cell = sheet.getActiveCell();

  if (cell.getA1Notation() === 'A1') {
    return true;
  }

  return false;
}
