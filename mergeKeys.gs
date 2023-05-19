/**
 * @fileoverview Code related to merge keys. Merge keys are used
 * to pass context specific data (i.e. sender email of opened email) to
 * implementations of the different integration types (i.e. email field of
 * a record for a Records based integration).
 */
MERGE_KV_PAIRS = {};

mergeInit();

/**
 * Initialtize merge keys. Called each time this file is loaded to ensure
 * merge key data is always properly initialized.
 */
function mergeInit() {
  let config = getConfig();

  if (config) {
    MERGE_KV_PAIRS['{{toolName}}'] = config.toolName;
    MERGE_KV_PAIRS['{{companyName}}'] = config.companyName;
  }
}

/**
 * Adds a merge key/value pair, making it later available by calling
 * getMergeKeyValue.
 * 
 * @param {string} key
 * @param {Object} value
 */
function addMergeKeyValuePair(key, value) {
  MERGE_KV_PAIRS[key] = value;
}

/**
 * Retrieves a merge value based on its key.
 * 
 * @param {string} key - The merge key associated (i.e. '{{sender_email}}')
 */
function getMergeKeyValue(key) {
  if (MERGE_KV_PAIRS.hasOwnProperty(key)) {
    return MERGE_KV_PAIRS[key];
  }

  return undefined;
}

/**
 * Returns a structure containing all merge key/value pairs
 * 
 * @return {Object}
 */
function getAllMergeKeyPairs() {
  return MERGE_KV_PAIRS;
}


/**
 * Replaces all occurences of each merge key in the target string with its
 * corresponding value from the given key/value pairs object.
 *
 * @param {Object} kvPairs Object of merge key/value pairs
 * @param {string} targetString String with merge keys that should be replaced.
 *
 * @return {string}
 */
function findAndReplaceMergeKeys(targetString, opts) {
  if (typeof targetString !== 'string') {
    return targetString;
  }

  let kvPairs = getAllMergeKeyPairs();

  // search for a {{mergetag}}
  let regex = /.*?(\{\{.*?\}\}).*?/g;

  //console.log("targetString before: " + targetString);

  // look for default-value merge keys in targetString of the form
  // {{actualMergeKey || defaultValue}}. If actualMergeKey is present in kvPairs
  // then replace the entire pattern with {{actualMergeKey}}. Else, replace it
  // with defaultValue. All of this is done prior to actually replacing any
  // merge keys in the targetString.
  let match;
  match = regex.exec(targetString);

  while (match && match.length > 1) {
    // check for a default value to use if merge string not found
    let mergeTag = match[1];
    //console.log('mergeTag=' + mergeTag)
    let mergeDefaultPair = mergeTag.split('||');
    //console.log('mergeDefaultPair =' + mergeDefaultPair)
    if (mergeDefaultPair.length > 1) {
      let actualMergeTag = mergeDefaultPair[0].trim() + '}}';
      if (
        !kvPairs.hasOwnProperty(actualMergeTag) ||
        kvPairs[actualMergeTag] === ''
      ) {
        let defaultMergeValue = mergeDefaultPair[1].trim().slice(0, -2);
        //console.log('defaultMergeValue = ' + defaultMergeValue)
        targetString = targetString.replace(mergeTag, defaultMergeValue);
      } else {
        targetString = targetString.replace(mergeTag, actualMergeTag);
      }
    }

    match = regex.exec(targetString);
  }

  // Replace any merge keys in targetString with their corresponding values
  // based on the key/value pairs in kvPairs.
  for (let [key, value] of Object.entries(kvPairs)) {
    if (!key) continue; // skip odd blank key seen sometimes

    //console.log("checking for key: " + key)

    if (opts && opts.uriEncodeValue) {
      value = encodeURIComponent(value);
    }
    if (opts && opts.boldValue) {
      value = `<b>${value}</b>`;
    }

    targetString = targetString.replace(new RegExp(key, 'g'), value);
  }

  return targetString;
}

/**
 * Given a source string with a merge tag in it, checks if the
 * merge tag is of the type {{mergeKey || defaultValue}}, and if so
 * returns defaultValue. If no merge key is detected, or a merge key
 * is detected but is has no default value, then the empty-string is 
 * returned.
 * 
 * @param {string} sourceString - The string to check
 * 
 * @return {string}
 */
function getDefaultMergeValue(sourceString) {
  let regex = /.*?(\{\{.*?\}\}).*?/g;
  let match;

  match = regex.exec(sourceString);

  if (!match || match.length < 2) {
    return '';
  }

  let mergeTag = match[1];
  let mergeDefaultPair = mergeTag.split('||');

  if (mergeDefaultPair.length < 2) {
    return '';
  }

  let defaultMergeValue = mergeDefaultPair[1].trim().slice(0, -2);

  return defaultMergeValue;
}

/** Internal test function
 * 
 */
function _testMergeKeys() {
  mergeInit();

  addMergeKeyValuePair('{{senderEmal}}', 'dabouav@gmail.com');
  addMergeKeyValuePair('{{token}}', 'Dave Abouav');

  let targetString =
    '{{senderName || Bobby McFee}} plus also {{senderEmail}}'
    + 'and {{token || participants}}';
  console.log('final result: ' + findAndReplaceMergeKeys(targetString));
}
