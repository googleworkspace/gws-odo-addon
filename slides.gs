/**
 * @fileoverview Functions specific to how the Add-on looks and acts when
 * in the context of the Slides editor.
 */

/**
 * Function that's called (per manifest) for Slides homepage trigger
 *
 * return {CardService.Card} Card to show for Slides Homepage
 */
function onSlidesHomepage(e) {
  console.log("e: " + JSON.stringify(e));

  if (e.slides.addonHasFileScopePermission) {
    return buildIntegrationCard(CALL_CONTEXT.SLIDES);

  }

  return buildFilePermissionCard();
}

/**
 * Function that's called (per manifest) when file access is granted
 * with drive.file scope.
 *
 * return {CardService.Card} Card to show on success
 */
function onSlidesFileScopeGranted(e) {
  return onSlidesHomepage();
}
