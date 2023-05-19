/**
 * @fileoverview Constants and definitions used throughout the Odo Add-on.
 */

/**
 * Contexts from which integration specific handlers can be called.
 * @enum {string}
 */
const CALL_CONTEXT = {
  GMAIL_VIEW: 'GMAIL_VIEW', // viewing a message in Gmail
  GMAIL_HOMEPAGE: 'GMAIL_HOMEPAGE', // homepage in Gmail
  GMAIL_COMPOSE: 'GMAIL_COMPOSE', // composing a message in Gmail
  SHEETS: 'SHEETS', // called from a Google Sheet
  DOCS: 'DOCS', // called from a Google Doc
  SLIDES: 'SLIDES', // called from Google Slides
  DRIVE: 'DRIVE', // call from Google Drive
  DEFAULT: 'DEFAULT', // used to handle any context if none has been defined
};

// Image/icon shown in headers and Record fields
// Some also used in sidebar (ergo also defined in the manifest file)
const ODO_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/stars_black_24dp.png';
const RECORD_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/article_black_24dp.png';
const TEXT_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/notes_black_24dp.png';
const EMAIL_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/email_black_24dp.png';
const PERSON_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/person_black_24dp.png';
const ORG_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/corporate_fare_black_24dp.png';
const DATE_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/event_black_24dp.png';
const FILE_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/file_present_black_24dp.png';
const FOLDER_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/folder_open_black_24dp.png';
const CHECKOUT_FILE_ICON =
  'https://www.gstatic.com/images/icons/material/' +
  'system/1x/check_black_24dp.png';
const OPEN_EMAIL_ICON =
  'https://storage.googleapis.com/odo-workspace-add-on/' +
  'open_email_message2.png';
const SELECT_FILE_ICON = 
  'https://storage.googleapis.com/odo-workspace-add-on/' +
  'select_file2.png'; // https://freeicons.io/test/file-icon-3237#

const DEFAULT_CUSTOMER_NAME = 'Cymbal';
const DEFAULT_CUSTOMER_LOGO_URL =
  'https://storage.googleapis.com/' +
  'odo-workspace-add-on/cymbal_logo_new2.png';
const DEFAULT_CUSTOMER_TOOL_NAME = 'Clang';

const ODO_LOGO_URL =
  'https://storage.googleapis.com/' +
  'odo-workspace-add-on/odo-logo-animated6.gif';

const USER_PROPERTY_CONFIG = 'USER_PROP_CONFIG';
const ODO_DATA_FOLDER_NAME = 'ODO_DATA';

const THIRD_PARTY_SERVICE_URL =
  'https://script.google.com/macros/s/' +
  'AKfycbwao0FuYr5m7NTRMIBZcVdEiPZSN7L4cFrILcvQgr' +
  '5-GRPOe9JKkNcCh5bP9TZaZAAl/exec';

const DEFAULT_RECORD_FILE =
  'https://drive.google.com/file/d/' + '1EJSDuBYSMXzNEiCHXIekOEg8b4JqcX8f/view';

// Properties
const PROP_STRINGIFIED_INTEGRATION_CUSTOMIZATIONS =
  'PROP_STRINGIFIED_INTEGRATION_CUSTOMIZATIONS';
