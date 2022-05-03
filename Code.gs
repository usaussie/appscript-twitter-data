/**
 * 
 * When the spreadsheet is open, add a custom menu so the owner can 
 * set up the sheet with everything that's needed
 * 
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var customMenuItems = [
    {name: 'Create Config & Headers (One-Time)', functionName: 'set_up_sheet'},
    {name: 'Process Tweet Data', functionName: 'processJSON'},
  ];
  spreadsheet.addMenu('Tweet Data', customMenuItems);
}

/**
 * 
 * helper function that kicks off others for initial 
 * set up of the sheet before first use
 * 
 */
function set_up_sheet() {

  create_config_tab();
  set_sheet_headers();

}
/**
 * 
 * helper function to set the sheet headers, so the column headers are there before the data is loaded
 * 
 */
function set_sheet_headers() {
  
  var tweet_data_tab_name = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config").getRange('B3').getValue();
  var results_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tweet_data_tab_name);
  results_sheet.appendRow(["id", "created_at", "retweeted", "source", "source_normalized", "favorite_count", "truncated", "retweet_count", "favorited", "full_text", "lang", "tweet_link"]);

}

/**
 * 
 * create the configuration tab, so the user can manage 
 * the settings via the sheet and not here in the apps script itself
 * 
 */
function create_config_tab() {

  // create the tab and set initial variables
  SpreadsheetApp.getActiveSpreadsheet().insertSheet("Config", 0);
  var results_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  results_sheet.appendRow(["Variable_Name", "Variable_Value"]);
  results_sheet.appendRow(["Tweet.js Google Drive ID", "{{ see comment for expected value }}"]);
  results_sheet.appendRow(["Tweet Data Tab Name", "Sheet1"]);
  results_sheet.appendRow(["Twitter Handle", "your_handle_here"]);

  // set comments to help the user know what to put into the  fields
  // add comment for the tweet.js file ID
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config").getRange('B2').setComment('Paste in only the ID of the Tweet.js file in Google Drive. IE: If the "Get Link" for the file in Google Drive is "https://drive.google.com/file/d/1vPIx2oevcxrFi-pO2-7gXRsOm3v-v/view?usp=sharing" then the ID would be "1vPIx2oevcxrFi-pO2-7gXRsOm3v-v"');

  // add comment for the config tab
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config").getRange('B3').setComment('Name of the tab in this sheet that you want to house the parsed tweet data. Typically just "Sheet1" but you can rename it or create a new tab if you want to.');
  
  // add comment for the twitter handle value
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config").getRange('B4').setComment('Your Twitter handle, without the @ symbol. This is used for linking to the actual tweet, using a concatenation of the Twitter URL and the tweet ID.');

  // send toast notification to the browser to tell people to go update the config values now
  SpreadsheetApp.getActiveSpreadsheet().toast('Please update the config values now, according to the very helpful comments');

}
/**
 * 
 * process the JSON and normalize it, then load into the google sheet
 * 
 */
function processJSON() {
  
  // set up vars for the tweet.js file, as well as the tab name that will house the results
  var tweet_data_sheet = SpreadsheetApp.getActiveSpreadsheet();
  var tweet_json_file_id = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config").getRange('B2').getValue();
  var tweet_data_tab_name = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config").getRange('B3').getValue();
  var twitter_handle = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config").getRange('B4').getValue();
  // set up vars for twitter link concatenation - https://twitter.com/{{handle}}/status/{{id}}
  var twitter_link_prefix = "https://twitter.com/";
  var twitter_link_suffix = "/status/";
  
  // make sure we can access the specified tweet.js before going any further
  try {
    var file = DriveApp.getFileById(tweet_json_file_id);
  }
  catch(err) {
      throw "ERROR: Could not open the specified tweet.js file ID. Please check the ID value and its permissions.";
  } 
  
  // grab the file from google drive, as a blob, then string, so we can parse the JSON inside it
  var file = DriveApp.getFileById(tweet_json_file_id);
  var fileBlob = file.getBlob()
  var fileBlobAsString = fileBlob.getDataAsString();

// make sure we can access the specified tweet.js before going any further
  try {
    // remove the "window.YTD.tweet.part0 = " from the first line which messes up the JSON parsing
    let substring = "window.YTD.tweet.part0 = "
    fileBlobAsString = fileBlobAsString.substring(substring.length)
    var parsedJson = JSON.parse(fileBlobAsString); //
  }
  catch(err) {
      throw "ERROR: Could not parse JSON. Check to see if the format has changed since this script was created.";
  } 
  
  // create holding array for the data, which will be added to the google sheet later
  var normalized_data = [];

  // loop through the json and grab what we need from it
  for (i in parsedJson){
    // console.log(parsedJson[i].tweet);
    // return;
    // parse source value so we can better report on it without all the rest of the HTML junk
    source_normalized_value = extractTextFromHTMLLinkTag(parsedJson[i].tweet.source);

    // create var for tweet link
    var tweet_link = twitter_link_prefix + twitter_handle + twitter_link_suffix + parsedJson[i].tweet.id;
  
    // grab the data from the tweet JSON and parse into an array to be added to the google sheet
    newRow = [
        parsedJson[i].tweet.id,
        new Date(parsedJson[i].tweet.created_at),
        parsedJson[i].tweet.retweeted,
        parsedJson[i].tweet.source,
        source_normalized_value,
        parsedJson[i].tweet.favorite_count,
        parsedJson[i].tweet.truncated,
        parsedJson[i].tweet.retweet_count,
        parsedJson[i].tweet.favorited,
        parsedJson[i].tweet.full_text,
        parsedJson[i].tweet.lang,
        tweet_link
      ];
      
      //console.log(newRow);

      normalized_data.push(newRow);


  }

  //Logger.log(newRow);
  //Logger.log(normalized_data);
  // write data to the google sheet
  writeDataToGoogleSheet(tweet_data_sheet, tweet_data_tab_name, normalized_data);
}

/**
 * 
 * Helper function for writing the data to the google sheet
 * 
 */
function writeDataToGoogleSheet(sheet_object, tab_name, data) {

  var newRow = [];
  var rowsToWrite = [];

  for (var i = 0; i < data.length; i++) {

    // define an array of all the object keys
    var headerRow = Object.keys(data[i]);

    // define an array of all the object values
    var newRow = headerRow.map(function(key){ return data[i][key]});

    // add to row array instead of append because append is SLOOOOOWWWWW
    rowsToWrite.push(newRow);

  }

  // select the range and set its values
  var ss = sheet_object.getSheetByName(tab_name);
  ss.getRange(ss.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);

}

/**
 * Extract value from an HTML tag, like an A Tag
 */
function extractTextFromHTMLLinkTag(linkText) {
    
    // expects: something like '<a href="#" class="taskName">foo bar baz</a>'
    // returns: foo bar baz
    
    return linkText.match(/<a [^>]+>([^<]+)<\/a>/)[1];
}
