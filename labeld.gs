// The Gmail label name to use
const labelTag = "labeld"

// The list of "From" email address 
const targetFrom = [
  "user@example.com",
  "another@example.com",
  "lastone@example.com"
]

// The list of subject filters for the messages
// as per targetFrom's indexes
const targetFilters = [
  "Invitation to edit",
  "Rule triggered: ",
  "Alert: "
]

// The list of task types as per targetFrom's indexes
const targetTypes = [
  "Research",
  "Investigation",
  "Troubleshooting"
]

// the list of task sources as per targetFrom's indexes
const targetSources = [
  "Team",
  "System",
  "System"
]

// regular expression reference to find a certain string in a 
// message's snippet
const targetSourceRegexp = '.+ has deleted the file .+'
const subjectIDRegexp = [
  "^Invitation to edit .* ([0-9][0-9][0-9][0-9][0-9][0-9])$",
  "^Rule triggered: .* ([0-9][0-9][0-9][0-9][0-9][0-9])$",
  "^Alert: .* ([0-9][0-9][0-9][0-9][0-9][0-9])$"
]

// shortQuery and longQuery variables will define a default amount of time 
// to look back in Gmail 
const shortQuery = " newer_than:7d"
const longQuery = " newer_than:9999d"

// NewUser boolean value will set the query to be long or short
// a new user will gather all content while an existing user will fetch the latest results
var NewUser = false

// Test function to test logic without touching 
// the Spreadsheet
function test() {

  messages = []
  var pageToken
  for (i in targetFrom) {
    do {
      page = Gmail.Users.Messages.list('me', {
        "q": "from:" + targetFrom[i] + ",subject:\"" + targetFilters[i] + "\"",
        "maxResults": 250,
        "pageToken": pageToken,
      })
      messages.push(page.messages)
      pageToken = page.nextPageToken;
    } while (pageToken)
  }
  Logger.log(messages)
}

// newUserSetup function will setup a new user in the Sheet
// and Apps Script triggers
// It will also setup Gmail labels if not existing 
// and gather the existing data from the Inbox
function newUserSetup() {
  // create Gmail labels if they don't exist
  setupGmailLabels()

  // populate the spreadsheet for the first time
  // with all possible results
  runGmailLabelQuery()

  // get active user
  activeUser = Session.getActiveUser()

  // check if triggers already exist in this project 
  // (it checks per user), and creates one if none exists
  var triggers = ScriptApp.getProjectTriggers()
  if (triggers.length <= 0 ) {
    // new time-based trigger every # minutes
    ScriptApp.newTrigger("runGmailLabelQuery")
        .timeBased()
        .everyMinutes(15)
        .create()
  }
}

// setupGmailLabels function will look if the set label 
// already exists and create it if it doesn't
// after creating the label, it will lookup the inbox 
// for all messages matching the constants defined, 
// and modify them with the created label (apply for all messages)
function setupGmailLabels() {
  // get all user labels
  var userLabels = Gmail.Users.Labels.list("me")

  // check if existing labels contain the defined tag in the constants
  if (userLabels.labels.find(x => x.name === labelTag)) {

    // set labelID with its id if exists
    var labelID = userLabels.labels.find(x => x.name === labelTag).id

  } else {

    // otherwise, create a new label with this tag
    Gmail.Users.Labels.create(
      {
        "labelListVisibility": "labelShow",
        "messageListVisibility": "show",
        "name": labelTag
      },
      'me'
    )

    // retrieve the labels once again
    userLabels = Gmail.Users.Labels.list("me")

    // get the created label's ID as labelID
    labelID = userLabels.labels.find(x => x.name === labelTag).id
  }

  // default state is that the filter doesn't exist, so set as false
  var filterExists = false

  // check for existing user-created email filters
  var userFilters = Gmail.Users.Settings.Filters.list("me")
  
  // if the response is not null (no custom filters ever created)
  if (userFilters != null) {
    
    // replay this action (...while filterExists is false)
    do {

      // iterate through each filter
      for (var i = 0 ; i < userFilters.filter.length ; i++) {

        // try to find the labelID previously retrieved in the addLabelIds action
        if ( (userFilters.filter[i].action) && (userFilters.filter[i].action.addLabelIds) && userFilters.filter[i].action.addLabelIds.length > 0) {

          // loop through all label IDs
          for (var x = 0 ; x < userFilters.filter[i].action.addLabelIds.length ; x++) {
            if (userFilters.filter[i].action.addLabelIds[x] == labelID) {
              
              // if found, filterExists is set to true, and break away from the loop
              filterExists = true
              break
            }
          }
        }
      }

      // if the expected filter doesn't really exist, create it
      // and modify all related messages to contain it
      if (filterExists != true) {
        
        // create the Gmail filter with the labelID to apply
        createGmailFilters(labelID)

        // apply the new labels to all related messages in the inbox
        applyGmailLabels(labelID)

        // finally, set filterExists as true
        filterExists = true
      }
    } while (filterExists != true)

  } else {
    // otherwise (if the user has no filters; if userFilters == null)
    // create the Gmail filter with the labelID to apply
    createGmailFilters(labelID)

    // apply the new labels to all related messages in the inbox
    applyGmailLabels(labelID)
  }
}

// createGmailFilters function will create a Gmail filter
// for each targetFrom address, setting the input labelID
function createGmailFilters(labelID) {

  // iterate through each targetFrom address
  for (var i = 0 ; i < targetFrom.length; i++) {
    // create a new Gmail filter with labels: 
    //     starred; important; primary; and labelID
    // criteria is from address matches this (or, each) targetFrom address
    Gmail.Users.Settings.Filters.create(
      {
        "action": {
          "addLabelIds": [
            "IMPORTANT",
            "CATEGORY_PERSONAL",
            labelID
          ],
        },
        "criteria": {
          "from": targetFrom[i],
          "subject": targetFilters[i]
        }
      },
      "me"
    )
    Logger.log('Created filter for ' + targetFrom[i] + ' using label ' + labelID)
  }
}

// applyGmailLabels function will go through all the 
// related messages in the inbox (according to the filter)
// and tag them with the new label, in batches of 250 messages at a time
function applyGmailLabels(labelID) {
  
  var messages = []
  var pageToken
  
  // iterate through each targetFrom address
  for (i in targetFrom) {

    // replay action (...while there is a nextPageToken value)
    do {

      // list all Gmail messages with the following query:
      //     from:{targetFrom},subject:"{targetFilters}"
      page = Gmail.Users.Messages.list(
        'me', 
        {
          "q": "from:" + targetFrom[i] + ",subject:\"" + targetFilters[i] + "\"",
          "maxResults": 250,
          "pageToken": pageToken,
        }
      )

      // add page to the messages list; this is created encapsulated lists
      messages.push(page.messages)

      // grab nextPageToken as pageToken to replay the request
      pageToken = page.nextPageToken;
    } while (pageToken)
  }
  Logger.log('Grabbed' + messages.length +  'message batches with the query')
  
  // to retrieve only the messageIDs (from an object with .id and .threadId)
  // the fastest method is to refer to C-style for-loops
  var messageIDs = []

  // iterate through the number of pages in the messages results
  for (var a = 0; a < messages.length; a++) {
    var idSet = []

    // if this page is not null
    if (messages[a]) {
      
      // iterate through each item in the page
      for (var b = 0; b < messages[a].length; b++) {

        // if there is an .id key, get its value into idSet
        if (messages[a][b].id) {
          idSet.push(messages[a][b].id)
        }
      }

      // after iterating through each page, push the array of IDs into 
      // the messageIDs array
      messageIDs.push(idSet)
    }
  }

  // with all messageIDs separated in pages of up to 250 elements
  // it's easy to batchModify these messages to contain the label;
  // iterate through each page in messageIDs
  for (var a = 0; a < messageIDs.length; a++) {

    // batchModify request to add the defined label to this page of 
    // messageIDs
    Gmail.Users.Messages.batchModify(
      {
        "addLabelIds": [
          labelID
        ],
        "ids": messageIDs[a]
      },
      'me'
    )
    Logger.log('Applied label to message batch #' + a)
  }
}

// initSheet function will initialize a Sheet in the document
// for a new user, on setup
function initSheet(sheet) {

  // Defines headers for the table if non-existent
  if (sheet.getRange("A1").getValue() != "From") {
    sheet.getRange("A1").setValue("From")  
  }
  
  if (sheet.getRange("B1").getValue() != "To") {
    sheet.getRange("B1").setValue("To")  
  }
  
  if (sheet.getRange("C1").getValue() != "Snippet") {
    sheet.getRange("C1").setValue("Snippet")  
  }
  
  if (sheet.getRange("D1").getValue() != "Task Type") {
    sheet.getRange("D1").setValue("Task Type")  
  }
  
  if (sheet.getRange("E1").getValue() != "Task Source") {
    sheet.getRange("E1").setValue("Task Source")  
  }
  
  if (sheet.getRange("F1").getValue() != "Time") {
    sheet.getRange("F1").setValue("Time")  
  }
  
  if (sheet.getRange("G1").getValue() != "Message") {
    sheet.getRange("G1").setValue("Message")  
  }
  
  if (sheet.getRange("H1").getValue() != "ID") {
    sheet.getRange("H1").setValue("ID")  
  }
  
  if (sheet.getRange("I1").getValue() != "Unix timestamp") {
    sheet.getRange("I1").setValue("Unix timestamp")  
  }
  
  if (sheet.getRange("J1").getValue() != "Reference ID") {
    sheet.getRange("J1").setValue("Reference ID")  
  }
  
  if (sheet.getRange("K1").getValue() != "Duplicate") {
    sheet.getRange("K1").setValue("Duplicate")  
  }
  Logger.log("Initialized spreadsheet's headers")

}

// getLatest function will capture the timestamp value
// for the latest registered message in the sheet, for a user
function getLatest(sheet) {

  // Getting the latest value present in the sheet 
  // by looking through all the Unix Timestamp cells
  // and storing the last value
  var range = "I2:I50000"
  var cells = sheet.getRange(range).getValues();
  
  // Loops through each cell and stores its value 
  // while the cell isn't empty, also storing the
  // empty cell number
  for (var i = 0 ; i < cells.length ; i++) {
    if (cells[i][0] === "" && !blank) {
      var blank = true
      var blankRow = (i+2)
      break
    } else {
      var blank = false
      var lastValue = cells[i][0]
    }
  }

  // In case there are no entries, all messages are fetched
  if (!lastValue) {
    var lastValue = 0
    Logger.log("No values found. Fetching all that is reachable")
  }
    
  // Returns both the blank row number 
  // and the last unix timestamp found
  return [blankRow, lastValue]
}


// getLatestBacklog function will find references in the existing Sheets backlog
// to match when its time to push the events into Sheets, marking as duplicate
// those events that appear as repeated
function getLatestBacklog(sheet, startRow, numRows) {

  // define columns to pick data from
  var sheetCols = [
    "D",
    "E",
    "J"
  ]

  var block = [];
  var object = {};
  var backlogObjs = [];
 

  // iterate through each column for the set range, to retrieve its values
  for (a = 0 ; a < sheetCols.length ; a++ ) {
    var array = [];
    var range = sheetCols[a] + (startRow-numRows) + ":" + sheetCols[a] + startRow
    var cells = sheet.getRange(range).getValues();

    // iterate through each value, and push it to a temporary array / list
    for (var b = 0 ; b < cells.length ; b++) {
      if (cells[b][0] === "" && !blank) {
        break      
      } else {
        array.push(cells[b][0])
      }
    }
    
    // push each temporary array into a block (or, array of arrays) like a map
    block.push(array)
  }

  // create an object from the set of arrays above, containg the respective information
  // about the event
  for (c = 0 ; c < block[0].length ; c++) {
    object = {
      type: block[0][c],
      source: block[1][c],
      ref: block[2][c]
    }

    // push each object to the backlogObjs list
    backlogObjs.push(object)
  }

  // return the list of objects
  return backlogObjs

}

// getCurrentBacklog function will find references in the existing messages (in-memory)
// backlog to match repeated ones when its time to push the events into Sheets, marking 
// as duplicate those events that appear as repeated
function getCurrentBacklog(entries) {

  var backlogRefs = [];
  var backlogUnix = [];
  var curBacklogObjs = [];

  // iterate through each collected message, from oldest to newest
  for ( var i = (entries.length - 1) ; i >= 0 ; i-- ) {

    // if the entry's reference ID is not present in backlogRefs, it's not a duplicate
    if (!backlogRefs.includes(entries[i].ref)) {
      object = {
        unix: entries[i].unix,
        ref: entries[i].ref,
        dup: false
      }
      // add this reference ID to backlogRefs and unix time to backlogUnix
      backlogRefs.push(entries[i].ref)
      backlogUnix.push(entries[i].unix)
    } else {
      // if the reference ID was found in backlogRefs, get its last reference unix time
      var index = backlogRefs.lastIndexOf(entries[i].ref)

      if (((backlogUnix[index] * 1) + 259200000) < entries[i].unix) {
        // if the unix time difference is lower than 3 days, the case is a duplicate
        object = {
          unix: entries[i].unix,
          ref: entries[i].ref,
          dup: true
        }
      } else {
        // if the unix time is greater than 3 days, then the case is not duplicate
        object = {
          unix: entries[i].unix,
          ref: entries[i].ref,
          dup: false
        }
        // add this reference ID to backlogRefs and unix time to backlogUnix
        backlogRefs.push(entries[i].ref)
        backlogUnix.push(entries[i].unix)
      }
    }

    // add the created object to curBacklogObjs, after the checks
    curBacklogObjs.push(object)
  }
  
  // return a list containing all created objects
  return curBacklogObjs
}


// gmailMessageQuery function will query Gmail messages
// for the set label, for the set number of days as per input
function gmailMessageQuery(newerThan) {
  var messages = [];
  var pageToken;

  Logger.log('Fetching messages with newer_than filter: %s', newerThan)
  // iterate through all pages (...while a nextPageToken exists)
  do {

    // list all the messages with the following query:
    //     label:{labelTag},newer_than:{days}d
    var response = Gmail.Users.Messages.list(
    'me', 
      {
        "q": "label:" + labelTag + newerThan,
        "pageToken": pageToken
      }
    );

  // if the response is not null, and lists out 1 or more results
  if (response.messages && response.messages.length > 0) {

    // push the ID for the message into the messages array
    response.messages.forEach(function(message) {
      messages.push(message)
    });
  }
    pageToken = response.nextPageToken;
  } while (pageToken);

  return messages
}


// handleSubjectRefs function will take in a message subject 
// and take its reference ID from RegExp matches
function handleSubjectRefs(subject) {
  var match
  var reference

  // iterate through all subject ID regular expression filters
  for ( var y = 0 ; y < subjectIDRegexp.length ; y ++) {

    // when the current subject matches one of the input regexp filters,
    // retrieve its reference ID from the matching substring and break the loop
    if (subject.match(RegExp(subjectIDRegexp[y]))) {
      match = subject.match(RegExp(subjectIDRegexp[y]))
      reference = match[1]
      break
    }
  }

  // return reference ID from subject
  return reference
}


// getLatestMessages function will cycle through the 
// user's inbox, looking up for the messages in the defined label
// returning the found entries according to the number of input days
function getLatestMessages(NewUser) {

  var entries = [];
  var usedThreadIDs = [];
  var message = {};

  // initiate a long or short query, depending if NewUser is set to true or false
  if (NewUser == true) {
    Logger.log('Fetching messages with a long query.')
    messages = gmailMessageQuery(longQuery)
  } else {
    Logger.log('Fetching messages with a short query.')
    messages = gmailMessageQuery(shortQuery)
  }
  

  // if the messages array exists and isn't null
  if (messages && messages.length > 0) {

    // loop through all message ids
    for (var i = 0 ; i < messages.length; i++) {

      // fetch each email message via its ID
      // if the response is not null
      if (response = Gmail.Users.Messages.get('me', messages[i].id)) {

        // do not reuse threadID's, to avoid repeated entries
        if (usedThreadIDs.includes(response.threadId))  {
          continue
        } else {

          // push new threadID to array
          usedThreadIDs.push(response.threadId)

          var snippet = response.snippet
          
          // iterate through its headers
          for (var x = 0 ; x < response.payload.headers.length ; x++) {

            // grab the Subject header into a variable
            if (response.payload.headers[x].name == 'Subject') {
              var subject = response.payload.headers[x].value
            }
            
            // grab the From header into a variable
            if (response.payload.headers[x].name == 'From') {
              var sender = response.payload.headers[x].value
            }

            // grab the To header into a variable
            if (response.payload.headers[x].name == 'To') {
              var to = response.payload.headers[x].value
            }
          }

          // define task type and task provider as per sender
          for (var x = 0 ; x < targetFrom.length ; x++) {
            if (sender.match(RegExp(targetFrom[x]))) {
              var taskType = targetTypes[x]
              var taskSource = targetSources[x]
              break
            }
          }

          // exceptions in case it's necessary to look into the 
          // message snippet to apply a different target source
          if (taskType == targetTypes[1] && snippet.match(RegExp(targetSourceRegexp))) {
            taskSource = targetSources[2];
          }

          // optional - retrieve the reference identifier from the subject value
          reference = handleSubjectRefs(subject)
      
          
          // compose a new message object with all the metadata
          message = {
                id: response.id,
                unix: response.internalDate,
                time: new Date(response.internalDate * 1),
                subj: subject,
                to: to,
                type: taskType,
                source: taskSource,
                sender: sender,
                snippet: snippet,
                ref: reference
              }
          
          // add the new message object to the entries array
          entries.push(message)
        }
        
      }
    }
  }

  // return the entries
  Logger.log(entries.length)
  return entries
}

// runGmailLabelQuery function will run a query
// and populate the Sheet with the retrieved data
function runGmailLabelQuery() {

  // get the active user
  activeUser = Session.getActiveUser()

  Logger.log('Setting up Sheet for: %s', activeUser)

  // Open the associated Spreadsheet and added Sheets
  var file = SpreadsheetApp.getActiveSpreadsheet();
  var namedSheets = file.getSheets();

  // setSheet is false while a user isn't found in the existing Sheets
  var setSheet = false

  // replay action (...while setSheet isn't true) 
  do {

    // iterate through all named sheets
    for (var i = 0 ; i < namedSheets.length; i++) {
      
      // if a Sheet exists with this user's name,
      if (namedSheets[i].getSheetName() == activeUser) {
        Logger.log('Matched user: %s', activeUser)

        // set the active sheet to it, and setSheet is now true
        var sheet = file.getSheets()[i]
        setSheet = true
        break
      } 
    }

    // otherwise, if this user doesn't have a Sheet yet
    if (setSheet == false) {

      // create it
      var sheet = file.insertSheet()

      // set Sheet name to the active user's
      sheet.setName(activeUser)
      Logger.log('Created new sheet for user: %s', activeUser)

      // setSheet is now true
      setSheet = true
    }
  } while (setSheet == false) 
  
  // First checks whether the Sheet is new and initialize it
  if (sheet.getRange("A1").getValue() === "") {
    Logger.log("Sheet seems blank, initializing.")
    initSheet(sheet)
  } else {
    Logger.log("Sheet check: OK")
  }

  // Fetches latest values and splits which is the next empty 
  // row as well as which is the last unix timestamp reference
  var getLatestValues = getLatest(sheet)
  var nextRow = getLatestValues[0]
  var latestID = getLatestValues[1] 

  Logger.log("nextRow: " + nextRow)
  Logger.log("latestID: " + latestID)

  // grab all entries when the Sheet is empty
  if ( latestID == 0 ) {
    NewUser = true
    Logger.log('No entries found, setting NewUser to: %s', NewUser)
  } 

  // backlog lookup will show duplicates if needed
  var backlogObjs = [];
  var curBacklogObjs = [];

  // Sheets backlog lookup will loop through the last 50 rows by default
  // 
  // this block will handle all 3 types of event where the user:
  //  - has less than 30 events listed in Sheets (so, from row 2 to nextRow-1)
  //  - has more than 30 events listed in Sheets (so, from nextRow-31 to nextRow-1)
  //  - has no events listed (skip Sheets backlog lookup)
  if ( latestID != 0 && nextRow < 30) {
    Logger.log("Performing backlog query from rows %s to %s", ((nextRow -1) - (nextRow -2)), (nextRow -1))

    backlogObjs = getLatestBacklog(sheet, (nextRow - 1), (nextRow - 2))
  } else if (latestID != 0 && nextRow > 30) {

    Logger.log("Performing backlog query from rows %s to %s", ((nextRow -1) - 30), (nextRow -1))
    backlogObjs = getLatestBacklog(sheet, (nextRow - 1), 30)
  } else {

    Logger.log("Skipping backlog query")
  }

  // get the newContent from a new Gmail query
  newContent = getLatestMessages(NewUser);


  // Message backlog lookup will loop through the retrieved messages from
  // the API query, and parses them in case any duplicates are included in the batch
  Logger.log("Performing backlog query for all fetched messages")
  curBacklogObjs = getCurrentBacklog(newContent)


  // iterate from last to first, through newContent
  for (var i = (newContent.length - 1) ; i >= 0 ; i-- ) {

    // if the current message is newer than the lastest retrieved,
    if (newContent[i].unix  > latestID) {

      var duplicate = false

      // check Sheets backlog for duplicates
      if (backlogObjs) {
        for ( var y = 0; y < backlogObjs.length ; y++ ) {
          // mark as duplicate when matching:
          //   - reference
          //   - type
          //   - source
          if (backlogObjs[y].ref == newContent[i].ref && backlogObjs[y].type == newContent[i].type && backlogObjs[y].source == newContent[i].source) {
            duplicate = true;
            break;
          }
        }
      }

      // check incoming message backlog for duplicates
      if (curBacklogObjs) {
        for ( var y = 0 ; y < curBacklogObjs.length ; y++) {
          // mark as duplicate is set via the getCurrentBacklog function
          // so duplicates are taken in as evaluated in the function call
          if (curBacklogObjs[y].unix == newContent[i].unix && curBacklogObjs[y].ref == newContent[i].ref) {
            duplicate = curBacklogObjs[y].dup
            break;
          }
        }
      }


      // add this content to the Sheet
      pushToSheets(
        sheet,
        nextRow,
        newContent[i].sender,
        newContent[i].to,
        newContent[i].snippet,
        newContent[i].type,
        newContent[i].source,
        newContent[i].time,
        newContent[i].subj,
        newContent[i].id,
        newContent[i].unix,
        newContent[i].ref,
        duplicate
      )
      
      // increment the nextRow value
      nextRow = (nextRow + 1);
    } 
  }
}

// pushToSheets function will have the boilerplate code to add 
// the input data to Sheets, with the desired formatting
function pushToSheets(sheet, newRow, sender, to, snippet, taskType, taskSource, time, subj, id, unix, ref, dup) {
  sheet.getRange(newRow, 1).setValue(sender);
  sheet.getRange(newRow, 2).setValue(to);
  sheet.getRange(newRow, 3).setValue(snippet);
  sheet.getRange(newRow, 4).setValue(taskType);
  sheet.getRange(newRow, 5).setValue(taskSource);
  sheet.getRange(newRow, 6).setValue(time);
  sheet.getRange(newRow, 6).setNumberFormat("dd/MM/yyyy HH:MM:SS");
  sheet.getRange(newRow, 7).setValue(subj);
  sheet.getRange(newRow, 8).setValue(id);
  sheet.getRange(newRow, 9).setValue(unix);
  sheet.getRange(newRow, 9).setNumberFormat("0000000000000");
  sheet.getRange(newRow, 10).setValue(ref);
  sheet.getRange(newRow, 10).setNumberFormat("00000000");
  sheet.getRange(newRow, 11).setValue(dup);
}
