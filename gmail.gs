/**
 * @fileoverview Code related to Cards and actions specific to the Add-on when
 * showing in the context of a user's Gmail.
 */

/**
 * Function that's called (per manifest) for Gmail homepage trigger
 *
 * return {CardService.Card} Card to show for Gmail Homepage
 */
function onGmailHomepage() {
  return buildIntegrationCard(CALL_CONTEXT.GMAIL_HOMEPAGE);
}

/**
 * Function that's called (per manifest) for Gmail message opened trigger.
 *
 * @return {CardService.Card} Card to show for opened message.
 */
function onGmailMessageOpened(event) {
  let message = getCurrentMessage(event);

  // form and pass values from this context to use as merge key/values
  let senderName = getMessageSenderName(message);
  let senderEmail = getMessageSenderEmail(message);

  addMergeKeyValuePair('{{senderName}}', senderName);
  addMergeKeyValuePair('{{senderEmail}}', senderEmail);
  addMergeKeyValuePair('{{token}}', senderEmail);

  return buildIntegrationCard(CALL_CONTEXT.GMAIL_VIEW);
}

/**
 * Function that's called (per manifest) for Gmail compose trigger.
 *
 * @return {CardService.Card} Card to show for opened message.
 */
function onGmailCompose(event) {
  let toRecipients = event.draftMetadata.toRecipients;
  let toRecipient = '';

  if (toRecipients.length > 0) {
    toRecipient = toRecipients[0];
  }

  return buildGmailComposeCard(toRecipient);
}

/**
 * Function that returns Gmail homepage Card
 *
 * return {CardService.Card} Card to show for Gmail Homepage
 */
function buildGmailHomepage() {
  let card = CardService.newCardBuilder();

  let brandedHeader = buildCustomerBrandedHeader();
  card.setHeader(brandedHeader);

  let section = CardService.newCardSection();

  section.addWidget(CardService.newImage().setImageUrl(OPEN_EMAIL_ICON));

  section.addWidget(
    CardService.newTextParagraph().setText(
      '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' +
        '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' +
        'Open an email to get started!'
    )
  );

  card.addSection(section);

  return card.build();
}

/**
 * Function that returns Gmail compose Card
 *
 * return {CardService.Card} Card to show in Gmail compose message
 */
function buildGmailComposeCard(toRecipient) {
  if (!toRecipient) {
    let card = CardService.newCardBuilder();

    let section = CardService.newCardSection();

    section.addWidget(
      CardService.newTextParagraph().setText(
        'No actions can be taken for an unaddressed email.'
      )
    );
    card.addSection(section);

    return card.build();
  }

  addMergeKeyValuePair('{{senderName}}', toRecipient);
  addMergeKeyValuePair('{{senderEmail}}', toRecipient);
  addMergeKeyValuePair('{{token}}', toRecipient);

  return buildIntegrationCard(CALL_CONTEXT.GMAIL_COMPOSE);
}

/**
 * Retrieves the current message given an action event object.
 * @param {Event} event Action event object
 * @return {Message}
 */
function getCurrentMessage(event) {
  let accessToken = event.messageMetadata.accessToken;
  let messageId = event.messageMetadata.messageId;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  return GmailApp.getMessageById(messageId);
}

/**
 * Determines date the email was received.
 *
 * @param {Message} message - The message currently open.
 * @returns {String}
 */
function getMessageReceivedDate(message) {
  return message.getDate().toLocaleDateString();
}

/**
 * Determines the name of whomever sent the message.
 *
 * @param {Message} message - The message currently open.
 * @returns {String}
 */
function getMessageSenderName(message) {
  let sender = message.getFrom();

  let senderName = sender.split('<')[0].trim();

  return senderName;
}

/**
 * Determines the email address of whomever sent the message.
 *
 * @param {Message} message - The message currently open.
 * @returns {String}
 */
function getMessageSenderEmail(message) {
  let sender = message.getFrom();
  let senderEmail = '';

  // look for email address in a string like 'Sender Name <sender@email.com>'
  let re = /[^< ]+(?=>)/g;

  let senderParts = sender.match(re);
  if (senderParts) {
    senderEmail = sender.match(re)[0];
  } else {
    // can be just straight up email address with no preceeding name or <> symbols
    if (sender.includes('@')) {
      senderEmail = sender;
    }
  }

  return senderEmail;
}
