////////////////////////////////// 80 cols /////////////////////////////////////

/**
 * @fileoverview Code related to common or universal Cards, actions, and
 * processing.
 */


/**
 *  Card builder that is called when user has yet to configure Odo.
 * 
 * @return {CardService.Card} Welcome Card for new Odo user 
 */
function buildOdoWelcomeCard() {

  let card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
      .setTitle('Welcome to Odo')
      //.setSubtitle('')
      .setImageStyle(CardService.ImageStyle.SQUARE)
      .setImageUrl(ODO_ICON));

  let section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph()
      .setText('Odo is a configurable Workspace Add-on that lets you '
               + 'demo Add-on capabilities, and help your customers '
               + 'envision their own Add-on solution in Workspace!'));
  
  let logoImage = CardService.newImage()
     .setImageUrl(ODO_LOGO_URL);

  section.addWidget(logoImage);

  section.addWidget(CardService.newTextParagraph()
      .setText('Click the button below to get started!'));
  
  let action = CardService.newAction()
      .setFunctionName('onOdoConfig');

  section.addWidget(CardService.newTextButton()
      .setText("Configure Odo")
      .setOnClickAction(action));

  card.addSection(section);
 
  return card.build();
}

/**
 * Builds and returns a Card header that is branded with the customer's name,
 * logo thumbnail, and tool/integration name.
 * 
 * @return {CardService.CardHeader} Card header with customer branding based on
 *     Odo configuration.
 */
function buildCustomerBrandedHeader() {

  let config = getConfig();

  let cardHeader = CardService.newCardHeader()
    .setTitle(config.toolName)
    .setSubtitle("by " + config.customerName)
    .setImageStyle(CardService.ImageStyle.SQUARE)
    .setImageUrl(config.customerLogoUrl);

  return cardHeader;
}

/**
 * Lets the user know that a page refresh is needed for saved changes
 * to take effect.
 * 
 * @return {CardService.Card} Card to show that a refresh is now needed .
 */
function buildRefreshNeededCard() {
  let card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
      .setTitle('Refresh Needed')
      .setImageStyle(CardService.ImageStyle.SQUARE)
      .setImageUrl(ODO_ICON));

  let descriptionSection = CardService.newCardSection();
  descriptionSection.addWidget(CardService.newTextParagraph()
      .setText('Your changes have been saved.<br><br>To ensure they take full ' 
             + 'effect, please refresh this page in your browser, as well '
             + 'as any other pages where the Odo Add-on is in use.'));

  card.addSection(descriptionSection);
   
  return card.build();
}


/**
 * Simple internal function to display a card with a message. For
 * debugging and development only.
 */
function showMessageCard(message) {
  let card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
      .setTitle('Message')
      .setImageStyle(CardService.ImageStyle.SQUARE)
      .setImageUrl(ODO_ICON));

  let messageSection = CardService.newCardSection();
  messageSection.addWidget(CardService.newTextParagraph()
      .setText(message));

  card.addSection(messageSection);
   
  return card.build();
}
