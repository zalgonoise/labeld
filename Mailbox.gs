/**
 * Mailbox class will handle all actions towards Gmail 
 * (directly or indirectly, via the Labeler property)
 */
class Mailbox {

    /**
     * @param {Object} config - config is an instance of LabelDConfig
     * with the appropriate settings to retrieve the data you need
     * from your Gmail labels
     */
    constructor(config) {
    this.query
    this.messageList = [];
    this.idMatrix = [];
    this.uniqueIDs = [];
    this.nestedIDs = [];
    this.rawEntries = [];
    this.entries = [];
    this.entriesBulk = [];
    this.labeler = new Labeler(config.labelTag)
    this.cfg = config

    /**
     * QueryFrom method will build a query string based on the input lists of 
     * from addresses and subject filters, encapsulating both lists respectively 
     * as a map or matrix. This will allow to contain more than one query within 
     * one search; ergo within one filter.
     * 
     * It defines its query property with this method, which is used to then fetch
     * messages, for example.
     * 
     * @param {string[]} address address - a list of email addresses to use
     * in the "from:" part of the filter
     * 
     * @param {string[]} filter filter - a list of strings to use in the 
     * "subject:" part of the filter
     */
    this.QueryFrom = function(address, filter) {
      var query

      for (var i = 0; i < address.length; i++) {
        if (query != null) {
          query = query + ` || `
        } 
        query = (query || "") + `(from:"${address[i]}",subject:"${filter[i]}")`
      }

      this.query = query

    }

    /**
     * QueryLabel method will create a query string based on an input label string.
     * It also supports scoping the Gmail message query to a number of days by adding
     * another parameter (`newerThan`), which is added to the query property if supplied 
     * as an argument
     * 
     * @param {string} label label - the string representation of the Gmail 
     * label
     * 
     * @param {int} newerThan newerThan - an integer value to scope the number of days 
     * for which to look back to, when fetching messages. This is an optional parameter.
     */
    this.QueryLabel = function(label, newerThan) {
      if (!newerThan) {
        this.query = `label:${label}`
        return
      }
      this.query = `label:${label},newer_than:${newerThan}d`
    }

    /**
     * SetOffset method will take the index difference between the inboxes' unique IDs and 
     * the Sheet's database number of records.
     * 
     * With this information it then iterates through the uniqueIDs list in reverse (newest to oldest)
     * storing a (reduced) list based on the offset.
     * 
     * The output list is then reversed (oldest to newest, as per Sheets order) and replaces the values 
     * in uniqueIDs
     * 
     * @param {int} offset offset - the index difference between the inboxes' unique IDs and 
     * the Sheet's database number of records.
     */
    this.SetOffset = function(offset) {
      var list = []
      
      for (var i = (this.uniqueIDs.length - 1); i >= (this.uniqueIDs.length - offset) ; i--) {
        list.push(this.uniqueIDs[i])
      }

      this.uniqueIDs = list.reverse()
    }

    /**
     * GetNewestID method will return the latest message ID in either the uniqueIDs list, or the 
     * idMatrix list (within the Message ID array)
     * 
     * @returns {string} messageID
     */
    this.GetNewestID = function() {
      if (this.uniqueIDs.length > 0) {
        return this.uniqueIDs[(this.uniqueIDs.length - 1)]
      } else {
        return this.idMatrix[0][(this.idMatrix[0].length - 1)]
      }
    }

    /**
     * ListMessages method will perform a Gmail user.messages.list query based on the
     * defined query parameter for Mailbox (configured before this call). It builds its
     * messageList property from this execution.
     * 
     * @param {bool} reverse reverse - setting this option to true reverses the order of the
     * input message IDs (from newest to oldest; to oldest to newest)
     */
    this.ListMessages = function(reverse) {
      this.messageList = [];
      var page
      var pageToken

      // replay action (...while there is a nextPageToken value)
      do {
        page = Gmail.Users.Messages.list(
          'me', 
          {
            "q": this.query,
            "maxResults": 250,
            "pageToken": pageToken,
          }
        )
        if (page.messages && page.messages.length > 0) {
          if (reverse) {
            page.messages.reverse()
          }
          this.messageList.push(page.messages)

          pageToken = page.nextPageToken;
        } else {
          pageToken = null
        }
      } while (pageToken)
      Logger.log(`Grabbed ${this.messageList.length} message batches with the query`)

      // if reverse is passed (for example the boolean true), then the list is reversed
      // or, older events show up first
      if (this.messageList && reverse) {
        this.messageList = this.messageList.reverse()
      }
    }

    /**
     * MakeMatrix method will take the available messageList property and create a 
     * map or matrix with the message ID and thread ID, kept separate (by iterating through
     * each block of message IDs and separating the properties from the object)
     */
    this.MakeMatrix = function() {
      var msgIDList = [];
      var threadIDList = [];

      if (this.messageList.length > 0) {
        for (var a = 0 ; a < this.messageList.length; a++) {
          if (this.messageList[a].length > 0) {
            for (var b = 0 ; b < this.messageList[a].length ; b++) {
              if (
                  (this.messageList[a][b]) 
                  && 
                  this.messageList[a][b].id != "" 
                  && this.messageList[a][b].threadId != ""
                ) {
                  msgIDList.push(this.messageList[a][b].id)
                  threadIDList.push(this.messageList[a][b].threadId)
                }
            }
          }
        }
      }
      this.idMatrix.push(msgIDList, threadIDList)
    }

    /**
     * DerefMatrix method will simply take the message IDs (disregarding the
     * thread IDs) from the idMatrix and define uniqueIDs with it
     */
    this.DerefMatrix = function() {
      if (this.idMatrix[0].length > 0) {
        this.uniqueIDs = this.idMatrix[0]
      }
    }      
    
    /**
     * DedupeMatrix method will take the idMatrix created to build a list
     * of only the unique messages in it, by iterating through each entry
     * while checking the thread ID for duplicates.
     * 
     * If a duplicate is found, the affected message ID value is *not* added 
     * to uniqueIDs.
     */
    this.DedupeMatrix = function() {
      var idList = [];

      // check if input isn't empty
      if (this.idMatrix[0].length > 0) {

        // iterate through thread ID matrix
        for (var a = 0 ; a < this.idMatrix[1].length ; a++ ) {

          // if the current thread ID hasn't been added to the idList yet:
          if (!idList.includes(this.idMatrix[1][a])) {
            
            // this is a unique message ID, and is added to uniqueIDs
            // also threadID is added to idList for the next iteration
            this.uniqueIDs.push(this.idMatrix[0][a])
            idList.push(this.idMatrix[1][a])
          }
        }
      }
    }

    /**
     * NestIDs method will take in an integer number which will group 
     * the uniqueIDs / idMatrix[0] in blocks of a fixed size.
     * 
     * This is especially useful to issue bulk requests or to iterate through
     * big amounts of data without fearing timing out before starting to actually 
     * write-out any data.
     * 
     * @param {int} length length - the block / list size to encapsulate Message IDs in
     */
    this.NestIDs = function(length) {
      var block = [];
      var range = this.uniqueIDs.length || this.idMatrix[0].length

      // check if input isn't empty
      if (range) {

        // iterate through input message IDs
        for (var a = 0 ; a < range ; a++) {

          // if the current block has fewer than ${length} items
          if (block.length < (length + 1)) {

            // add ID to block
            if (this.uniqueIDs.length > 0) {
              block.push(this.uniqueIDs[a])
            } else {
              block.push(this.idMatrix[0][a])
            }
            
          } else {

            // otherwise push this block to the main list and initiate a new block
            this.nestedIDs.push(block)
            block = [];
            if (this.uniqueIDs.length > 0) {
              block.push(this.uniqueIDs[a])
            } else {
              block.push(this.idMatrix[0][a])
            }
          }
        }

        // if at the end of the loop there are items in the current block
        // push it to the main list
        if (block.length > 0) {
          this.nestedIDs.push(block)
        }
      }
    }

    /**
     * GetMessages method will iterate through the uniqueIDs property and fetch
     * each message by its ID.
     * 
     * Each valid response is handled by a new MessageBuilder who creates the 
     * message object after processing its contents.
     * 
     * The processed message is added to the rawEntries list
     */
    this.GetMessages = function() {
      var response;

      if (this.uniqueIDs.length > 0) {
        for (var i = 0 ; i < this.uniqueIDs.length; i++) {
          if (response = Gmail.Users.Messages.get('me', this.uniqueIDs[i])) {
              var message = new MessageBuilder(response)
              this.rawEntries.push(message)
          }
        }
      }
    }

    /**
     * ProcessAllMessages method will iterate through the rawEntries list of objects and:
     * - check for duplicates from the input message list
     * - check for duplicates in the database backlog (if within the first 30 objects tested)
     * - based on this check, apply the Duplicate status if needed, and push to Sheets
     */
    this.ProcessAllMessages = function() {

      if (this.rawEntries.length > 0) {

        var db = new Database();
        db.LatestEntry()        
        var bl = new Backlog()
        bl.Database()

        for (var i = 0 ; i < this.rawEntries.length; i++) {
          var duplicate = bl.InputLookup(this.rawEntries[i].output)
          
          if (!duplicate && i < 30) {
            duplicate = bl.DatabaseLookup(this.rawEntries[i].output)
          }

          if (duplicate) {
            this.rawEntries[i].SetDuplicate(duplicate)
          }

          db.PushEntry(this.rawEntries[i])
          db.IncrementRow()

          this.entries.push(this.rawEntries[i].output)
          
        }      
      }
    }
    
    /**
     * GetAndProcessMessages method will iterate through the available message IDs list and:
     * - fetch the message by its ID
     * - create a message object with MessageBuilder based on the response
     * - check for duplicates from the input message list
     * - check for duplicates in the database backlog (if within the first 30 objects tested)
     * - based on this check, apply the Duplicate status if needed, and push to Sheets
     */
    this.GetAndProcessMessages = function() {
      var response;
      if (this.uniqueIDs.length > 0) {
        var ids = this.uniqueIDs
      } else {
        var ids = this.idMatrix[0]
      }

      if (ids.length > 0) {
        var db = new Database();
        db.LatestEntry()        
        var bl = new Backlog()
        bl.Database()

        for (var i = 0 ; i < ids.length; i++) {
          if (response = Gmail.Users.Messages.get('me', ids[i])) {
              var message = new MessageBuilder(response)

              var duplicate = bl.InputLookup(message.output)
              
              if (!duplicate && i < 30) {
                duplicate = bl.DatabaseLookup(message.output)
              }

              if (duplicate) {
                message.SetDuplicate(duplicate)
              }

              db.PushEntry(message)
              db.IncrementRow()

              this.entries.push(message.output)
          }
        }
      }
    }

    /**
     * GetOneMessage method will (most commonly for debugging) be used to 
     * fetch one single message by supplying a message ID, while creating a
     * message object using a new MessageBuilder with the retrieved response
     * 
     * @param {string} msgID msgID - the Gmail message ID to fetch
     * @returns {Object} message - an object with the processed content of the message
     */
    this.GetOneMessage = function(msgID) {
      var response;
      if (response = Gmail.Users.Messages.get('me', msgID)) {
        var message = new MessageBuilder(response)
        return message.output
      }
    }

    /**
     * GetAndProcessMessages method will work with a nestedIDs list to fetch message
     * objects in parts, as per the Nesting block length.
     * 
     * The results are stored in entriesBulk instead, since it's a list of object lists 
     * instead ({Object[][]})
     */
    this.BulkGetMessages = function() {
      var block = [];
      var response;

      if (this.nestedIDs.length > 0) {
        for (var a = 0 ; a < this.nestedIDs.length; a++) {
          if (this.nestedIDs[a].length > 0) {
            for (var b = 0 ; b < this.nestedIDs[a].length ; b++) {
                if (response = Gmail.Users.Messages.get('me', this.nestedIDs[a][b])) {
                    var message = new MessageBuilder(response)
                    block.push(message.output)
                }  
            }
            this.entriesBulk.push(block)
            block = [];
          }
        }
        if (block.length > 0) {
          this.entriesBulk.push(block)
        }
      }
    }
    
    /**
     * BulkGetAndProcessMessages method will iterate through the available message IDs list 
     * in nestedIDs (after a mailbox NestIDs({int}) call), and:
     * - fetch the message by its ID
     * - create a message object with MessageBuilder based on the response
     * - check for duplicates from the input message list
     * - check for duplicates in the database backlog (if within the first 30 objects tested)
     * - based on this check, apply the Duplicate status if needed, and push to Sheets
     * 
     * Note that within each block the sequence is split and the check for duplicates occurs 
     * once again. This is intended as between blocks there could be duplicate entries that, otherwise,
     * would easily be missed out.
     */
    this.BulkGetAndProcessMessages = function() {
      var block = [];
      var response;

      if (this.nestedIDs.length > 0) {
        for (var a = 0 ; a < this.nestedIDs.length; a++) {
          if (this.nestedIDs[a].length > 0) {
            Logger.log(`Processing messages on block ${a + 1}`)

            var db = new Database();
            db.LatestEntry()        
            var bl = new Backlog()
            bl.Database()

            for (var b = 0 ; b < this.nestedIDs[a].length ; b++) {
              if (response = Gmail.Users.Messages.get('me', this.nestedIDs[a][b])) {
                var message = new MessageBuilder(response)
                var duplicate = bl.InputLookup(message.output)
                
                if (!duplicate && b < 30) {
                  duplicate = bl.DatabaseLookup(message.output)
                }

                if (duplicate) {
                  message.SetDuplicate(duplicate)
                }

                db.PushEntry(message)
                db.IncrementRow()

                this.entries.push(message.output)
                block.push(message.output)
              }  
            }
            if (block.length > 0) {
              this.entriesBulk.push(block)
            }
            block = [];
          }
        }
      }
    }

    /**
     * GetEntry method will return the corresponding entry item from entries, 
     * based on the supplied index value.
     * 
     * @param {int} index index - the index for `this.entries[index]`
     */
    this.GetEntry = function(index) {
      return this.entries[index]
    }

    /**
     * LookupFilters method will initiate a labeler by fetching the label ID 
     * and then fetching the filters. If when checked, the filter does not exist, 
     * a label ID is returned (set to filterToCreate). Looking if this variable is 
     * defined and populated will determine whether new filters and labels need to be 
     * applied
     */
    this.LookupFilters = function() {
      this.labeler.SetLabelID()
      this.GetFilters()
      var filterToCreate = this.labeler.CheckFilters()

      if (filterToCreate) {
        this.ApplyFilters(filterToCreate)
        this.ApplyLabel(filterToCreate)
      }
    }

    /**
     * GetFilters will define this mailboxes' labeler's filters by fetching them
     * from the Gmail API
     */
    this.GetFilters = function() {
      this.labeler.SetFilters(Gmail.Users.Settings.Filters.list("me"))
    }
    
    /**
     * ApplyFilters method will take in a label ID and a (pre-configured) query to 
     * create a Gmail filter.
     * 
     * @param {string} labelID labelID - the ID for the label to mark messages with,
     * when matched a certain criteria in a new Gmail filter
     */
    this.ApplyFilters = function(labelID) {
      this.QueryFrom(this.cfg.targets.targetFrom, this.cfg.targets.targetFilters)
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
            "query": this.query
            
          }
        },
        "me"
      )
      Logger.log(`Created filter using label ${labelID}`)
    }
   
   /**
    * ApplyLabel method will take in a label ID, and apply it to all
    * messages present in a query headed for the configured targetFrom 
    * and targetFilters.
    *
    * Message IDs are fetched from a new query, which is followed by a 
    * batchModify action (using the nested IDs)
    *
    * @param {string} labelID labelID - the unique identifier for the Label
    * to apply
    */
    this.ApplyLabel = function(labelID) {
      this.QueryFrom(this.cfg.targets.targetFrom, this.cfg.targets.targetFilters)
      this.ListMessages()
      this.MakeMatrix()
      this.DerefMatrix()
      this.NestIDs(250)

      for (var a = 0; a < this.nestedIDs.length; a++) {
        if (this.nestedIDs[a].length > 0 ) {
          Gmail.Users.Messages.batchModify(
            {
              "addLabelIds": [
                labelID
              ],
              "ids": this.nestedIDs[a]
            },
            'me'
          )
          Logger.log(`Applied label to message batch #${a}`)
        }
      }
    }
    /**
     * RemoveLabel method will fetch all messages marked with the configured label string,
     * and batchModify them (in bulk, with nestedIDs) to now remove the fetched label ID.
     */
    this.RemoveLabel = function() {
      // retrieve encapsulated list of unique message IDs in sets of 250 items by running
      //   - Gmail query using the input query string
      //   - breaking down the response into an ID Matrix
      //   - removing duplicates from ID Matrix
      //   - breakdown unique IDs list into blocks of 250 items
      this.QueryLabel(this.labeler.label)
      this.ListMessages()
      this.MakeMatrix()
      this.DerefMatrix();
      this.NestIDs(250);
      this.labeler.SetLabelID()

      for (var i = 0 ; i < this.nestedIDs.length ; i++) {
        Gmail.Users.Messages.batchModify(
          {
            ids: this.nestedIDs[i],
            removeLabelIds: this.labeler.labelID
          },
          "me"
        )
      }
    }

    /**
     * RemoveFilter method will generate a new list of configured Gmail filters,
     * and then remove the one which contains the same label being pointed to.
     */
    this.RemoveFilter = function() {
      this.labeler.SetLabelID() 
      this.GetFilters()

      Gmail.Users.Settings.Filters.remove("me", this.labeler.GetFilterID())
      Logger.log(`Removed Gmail filter with ID ${this.labeler.GetFilterID()}`)

    }  
  }
}

/**
 * MessageBuilder class will create and update message objects
 * as they are being processed for their content
 */
class MessageBuilder {
  /**
   * @param {Object} input input - the response from a Gmail.Users.Messages.get request
   */
  constructor(input) {
    
    this.input = input
    this.snippet = input.snippet
    this.subject;
    this.sender;
    this.to;
    this.taskType;
    this.taskSource;
    this.reference;
    this.duplicate = false;
    this.bodyRef;
    this.bodyRefURL;
    this.priority;
    this.level;
    this.output = new Message();
    this.cfg = new LabelDConfig()

    /**
     * SetID method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} id id - the task's message ID
     */
    this.SetID = function(id) {
      this.output.id = id
    }

    /**
     * SetTime method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {int} time time - the task's message Unix timestamp
     */
    this.SetTime = function(time) {
      this.output.unix = time
      this.output.time = new Date(time * 1);
    }


    /**
     * SetSubject method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} subject subject - the task's message subject
     */
    this.SetSubject = function(subject) {
      this.subject = subject
      this.output.subj = subject
    }

    /**
     * SetTo method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} to to - the task's message recipient value
     */
    this.SetTo = function(to) {
      this.to = to
      this.output.to = to
    }

    /**
     * SetTaskType method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} taskType taskType - the task's type value
     */
    this.SetTaskType = function(taskType) {
      this.taskType = taskType
      this.output.type = taskType
    }


    /**
     * SetTaskSource method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} taskSource taskSource - the task's source value
     */
    this.SetTaskSource = function(taskSource) {
      this.taskSource = taskSource
      this.output.source = taskSource
    }

    /**
     * SetSender method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} sender sender - the task's sender value
     */
    this.SetSender = function(sender) {
      this.sender = sender
      this.output.sender = sender
    }

    /**
     * SetSnippet method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} snippet snippet - the task's snippet from the message body
     */
    this.SetSnippet = function(snippet) {
      this.snippet = snippet
      this.output.snippet = snippet
    }

    /**
     * SetReference method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {int} ref ref - the task's reference ID
     */
    this.SetReference = function(ref) {
      this.ref = ref 
      this.output.ref = ref
    }

    /**
     * SetDuplicate method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {bool} status status - the task's duplicate status
     */
    this.SetDuplicate = function(status) {
      this.dup = status
      this.output.dup = status
    }

    /**
     * SetBodyReference method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} bodyRef bodyRef - the task's secondary reference from the message body
     */
    this.SetBodyReference = function(bodyRef) {
      this.bodyRef = bodyRef 
      this.output.bodyRef = bodyRef
    }

    /**
     * SetBodyReferenceURL method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} bodyRefURL bodyRefURL - the task's secondary reference URL from the message body
     */
    this.SetBodyReferenceURL = function(bodyRefURL) {
      this.bodyRefURL = bodyRefURL
      this.output.bodyRefURL = bodyRefURL
    }

    /**
     * SetPriority method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} priority priority - the task's priority from the message body
     */
    this.SetPriority = function(priority) {
      this.priority = priority
      this.output.priority = priority
    }

    /**
     * SetLevel method will apply the input value to both the MessageBuilder
     * and Message object (this.output) 
     * 
     * @param {string} level level - the task's level from the message body
     */
    this.SetLevel = function(level) {
      this.level = level
      this.output.level = level
    }

    /**
     * CheckHeaders method will iterate through the message headers
     * looking for matches for key properties found in them
     */
    this.CheckHeaders = function() {
      for (var x = 0 ; x < this.input.payload.headers.length ; x++) {
        if (this.input.payload.headers[x].name == 'Subject') {
          this.SetSubject(this.input.payload.headers[x].value)
        }

        if (this.input.payload.headers[x].name == 'From') {
          this.SetSender(this.input.payload.headers[x].value)
        }

        if (input.payload.headers[x].name == 'To') {
          this.SetTo(this.input.payload.headers[x].value)
        }
      }
    }
    
    /**
     * CheckTypeAndSource will look for evidence in the defined
     * target senders to define the message's task type and source.
     * 
     * Exceptions to this definition are evaluated in this method as well
     * such as if the subject matches the target source regexp 
     */
    this.CheckTypeAndSource = function() {
      // define task type and task provider as per sender
      
      for (var x = 0 ; x < this.cfg.targets.targetFrom.length ; x++) {
        
        if (this.sender.match(RegExp(this.cfg.targets.targetFrom[x]))) {
          this.SetTaskType(this.cfg.templates.targetTypes[x])
          this.SetTaskSource(this.cfg.templates.targetSources[x])
          
          break
        }
      }

      // exceptions in case it's necessary to look into the 
      // message snippet to apply a different target source
      if (
        this.taskType == this.cfg.templates.targetTypes[1] && 
          (
            this.snippet.match(RegExp(this.cfg.regex.targetSourceRegexp[0])) 
            || 
            this.snippet.match(RegExp(this.cfg.regex.targetSourceRegexp[1])) 
          )
        ) {
        this.SetTaskSource(this.cfg.templates.targetSources[2])
      }
    }

    /**
     * SubjectReferences method will look for the reference ID value
     * in the message's subject
     */
    this.SubjectReferences = function() {
      for ( var y = 0 ; y < this.cfg.regex.subjectIDRegexp.length ; y ++) {

        // when the current subject matches one of the input regexp filters,
        // retrieve its reference ID from the matching substring and break the loop
        if (this.subject.match(RegExp(this.cfg.regex.subjectIDRegexp[y]))) {
          var match = this.subject.match(RegExp(this.cfg.regex.subjectIDRegexp[y]))
          this.SetReference(match[1])
          break
        }
      }
    }

    /**
     * BodyReferences method will look for certain patterns in the message
     * body to define key properties of the message, such as the secondary
     * reference ID and its URL, the priority and the level of the task
     * 
     * Invalid or unspecified fields are defined as the hardcoded string "N/A"
     */
    this.BodyReferences = function() {
      function convertASCII(input) {
        var result = "";
        var output = [];
        for (var i = 0; i < input.length; i++) {
          result += String.fromCharCode(input[i])
        }
        output = result.split('\n')
        return output
      }

      if (this.input.payload.parts && this.input.payload.parts.length >= 2) {
        var chars = convertASCII(this.input.payload.parts[1].body.data)
        var match;

        for (var i = 0; i < chars.length; i++) {
          if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[0]))) {
            this.SetBodyReference(match[2])
            this.SetBodyReferenceURL(match[1])
            this.SetLevel("N/A")
          } else if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[4]))) {
            this.SetBodyReference("N/A")
            this.SetBodyReferenceURL(`${this.cfg.prefixes.baseURL}${match[1]}`)
          }

          if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[1]))) {
            this.SetPriority(match[1])
          } else if ((match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[2])))) {
            this.SetPriority(match[1])
          }

          if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[3]))) {
            this.SetLevel(match[1])
          }

        }
        return

      } else if (this.input.payload.parts && this.input.payload.parts.length == 1) {
        var chars = convertASCII(this.input.payload.parts[0].body.data)
        var match;

        for (var i = 0; i < chars.length; i++) {
          if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[2]))) {
            this.SetPriority(match[1])
          }
          if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[3]))) {
            this.SetLevel(match[1])
          }
          if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[4]))) {
            this.SetBodyReferenceURL(`${this.cfg.prefixes.baseURL}${match[1]}`)
          }
        }
        this.SetBodyReference("N/A")       
        return 
      } else {
        var chars = convertASCII(this.input.payload.body.data)
        var match;

        for (var i = 0; i < chars.length; i++) {
          if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[5]))) {
            this.SetPriority(match[1])
          }
          if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[6]))) {
            this.SetLevel(match[1])
          }
          if (match = chars[i].match(RegExp(this.cfg.regex.bodyIDRegexp[7]))) {
            this.SetBodyReferenceURL(`${this.cfg.prefixes.baseURL}${match[1]}`)
          }          
        }
        this.SetBodyReference("N/A")
        return
      }
                
    }

    /**
     * MessageBuilder runtime
     */      
    this.SetTime(this.input.internalDate)
    this.SetID(this.input.id)
    this.SetSnippet(input.snippet)
    this.SetDuplicate(false)
    this.CheckHeaders()
    this.CheckTypeAndSource()
    this.SubjectReferences()
    this.BodyReferences()
    return this
  }
}

/**
 * Message class is an isolated class to build a Message object
 * outside of the MessageBuilder class, not containing any of its
 * prototypes
 */
class Message {
  /**
   * @param {string} id id - the task's message ID
   * @param {int} time time - the task's message Unix timestamp
   * @param {string} subj subj - the task's message subject
   * @param {string} to to - the task's message recipient value
   * @param {string} type type - the task's type value
   * @param {string} taskSource taskSource - the task's source value
   * @param {string} sender sender - the task's sender value
   * @param {string} snippet snippet - the task's snippet from the message body
   * @param {int} ref ref - the task's reference ID
   * @param {bool} status status - the task's duplicate status
   * @param {string} bodyRef bodyRef - the task's secondary reference from the message body
   * @param {string} bodyRefURL bodyRefURL - the task's secondary reference URL from the message body
   * @param {string} priority priority - the task's priority from the message body
   * @param {string} level level - the task's level from the message body
   */
  constructor(id, time, subj, to, type, source, sender, snippet, ref, dup, bodyRef, bodyRefURL, priority, level) {
    this.id = id;
    this.unix = time;
    this.time = new Date(time * 1);
    this.subj = subj;
    this.to = to;
    this.type = type; 
    this.source = source; 
    this.sender = sender;
    this.snippet = snippet;
    this.ref = ref;
    this.dup = dup;
    this.bodyRef = bodyRef;
    this.bodyRefURL = bodyRefURL;
    this.priority = priority;
    this.level = level;
  }
}
