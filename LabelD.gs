/**
 * LabelD class will be the placeholder for labeld.
 * It manages the Gmail, Sheets and Apps Script integrations
 * to create a processed backlog for messages with a 
 * certain label
 */
class LabelD {

  /**
   * @param {Object} config - config is an instance of LabelDConfig
   * with the appropriate settings to retrieve the data you need
   * from your Gmail labels
   */
  constructor(config) {
    this.user = Session.getActiveUser();
    this.labelTag = config.labelTag
    this.mailbox = new Mailbox(config);
    this.triggers;
    this.database = new Database();
    this.backlog = new Backlog(); 

    /**
     * Labeler method will initiate the Labeler module in mailbox,
     * preparing the filter and label for Gmail if required
     */
    this.Labeler = function() {
      this.mailbox.LookupFilters()
    }

    /**
     * GetTriggers method will define the triggers item with this 
     * user's project triggers, for Apps Script
     */
    this.GetTriggers = function() {
      this.triggers = ScriptApp.getProjectTriggers();
    }


    /**
     * CheckTriggers method will check if there are triggers configured
     * for this user, otherwise create one accordingly
     */
    this.CheckTriggers = function() {
      this.GetTriggers()
      if (this.triggers.length <= 0 ) {
        // new time-based trigger every # minutes
        ScriptApp.newTrigger("runGmailLabelQuery")
            .timeBased()
            .everyMinutes(15)
            .create()
        this.triggers = ScriptApp.getProjectTriggers();
      }
    }

    /**
     * RemoveTriggers method will fetch and cycle through all configured 
     * Apps Script triggers, deleting them.
     */
    this.RemoveTriggers = function() {
      this.GetTriggers()

      for (var i = 0 ; i < this.triggers.length; i++) {
        ScriptApp.deleteTrigger(this.triggers[i])
      }
      Logger.log(`Apps Script Triggers have been removed`)
    }

    /**
     * Run method is the default runtime / opt-in function to run LabelD
     */
    this.Run = function() {

      /**
       * Initialize Labeler module in mailbox and check if 
       * triggers are configured already
       */
      this.Labeler();
      this.CheckTriggers()

      /**
       * Fetch and group all messages in inbox, as per defined label
       */
      this.mailbox.QueryLabel(this.labelTag)
      this.mailbox.ListMessages(true) // true is a boolean to reverse order, or, older first
      this.mailbox.MakeMatrix()
      this.mailbox.DedupeMatrix() // Dedupe vs Deref; Dedupe seems more accurate as per incoming tasks

      /**
       * Fetch latest entries in Sheets database for comparison
       */
      this.database.LatestEntry()

      /**
       * Offset will create an index of the available messages vs the existing 
       * entries in Sheets. Since all results must match, it is compared in here
       */
      var offset = this.mailbox.uniqueIDs.length - (this.database.blankRow - 2)

      /**
       * + mailbox entries match the database's count 
       * --> Run backlog.Comparison() method 
       */
      if (offset == 0) {
        if (this.backlog.Comparison(this.mailbox.GetNewestID())) {
          Logger.log("Sync OK!")
        }
        
      /**
       * + there are incoming messages, fewer than 250
       * --> Define offset in message IDs list
       * --> Get and process messages
       * --> Run backlog.Comparison() method 
       */
      } else if (offset > 0 && offset < 250) {
        
        this.mailbox.SetOffset(offset)
        
        this.mailbox.GetAndProcessMessages()
        if (this.backlog.Comparison(this.mailbox.GetNewestID())) {
          Logger.log("Sync OK!")
        }
        return

      /**
       * + there are incoming messages, more than 250
       * --> Define offset in message IDs list
       * --> Get and process messages in bulk
       * --> Run backlog.Comparison() method 
       */
      } else if (offset => 250) {
        this.mailbox.SetOffset(offset)
        this.mailbox.NestIDs(250)
        this.mailbox.BulkGetAndProcessMessages()
        if (this.backlog.Comparison(this.mailbox.GetNewestID())) {
          Logger.log("Sync OK!")
        }
        return
      }
      return


    }

    /**
     * OptOut method will remove labels from messages, remove the label itself,
     * the Gmail filter, and the Apps Script trigger
     */
    this.OptOut = function() {
      this.mailbox.RemoveLabel()
      this.mailbox.RemoveFilter()
      this.RemoveTriggers()
    }

    /**
     * LabelD runtime
     */
    Logger.log(`LabelD Runtime: ${this.user} on label ${this.labelTag}`)
    this.GetTriggers();

  }
}

/**
 * LabelIDConfig class is a placeholder object to isolate the access
 * to the needed config and variables used through LabelD.
 * 
 * This ends up being more organized, compact and scoped instead of using
 * constants and variables scoped to the whole project (or even worse,
 * defining them repeatedly)
 * 
 * Classes which need configs (only the Mailbox for now) have it implemented
 * as one of their properties (which they can call upon directly)
 */
class LabelDConfig {
  constructor(label, targets, templates, regex, prefixes) {
    this.labelTag = label || null
    /**
     * Query targets
     */
    this.targets = targets || {
      /**
       * "from:" address targets
       */
      targetFrom: [
        "user@example.com",
        "another@example.com",
        "lastone@example.com"
      ],
      /**
       * "subject:" string targets
       */
      targetFilters: [
        "Invitation to edit",
        "Rule triggered: ",
        "Alert: "
      ]
    }
    /**
     * Data templates
     */
    this.templates = templates || {
      /**
       * map / matrix mapped types, as per "from:" targets
       */
      targetTypes: [
        "Research",
        "Investigation",
        "Troubleshooting"
      ],
      /**
       * map / matrix mapped sources, as per "from:" targets
       */      
      targetSources: [
        "Team",
        "System",
        "System"
      ]
    }
    /**
     * Regular Expressions
     */
    this.regex = regex || {
      /**
       * Subject regular expressions for task sources
       */
      targetSourceRegexp: [
        '.+ has deleted the file .+',
        '.+ has created the file .+'
      ],
      /**
       * Subject regular expressions for task reference ID
       */
      subjectIDRegexp: [
        "^Invitation to edit .* ([0-9][0-9][0-9][0-9][0-9][0-9])$",
        "^Rule triggered: .* ([0-9][0-9][0-9][0-9][0-9][0-9])$",
        "^Alert: .* ([0-9][0-9][0-9][0-9][0-9][0-9])$"
      ],
      /**
       * Body regular expressions for task secondary ID, task URL, priority and level
       */
      bodyIDRegexp: [
        "File ID #(.*) - (https://.*)",
        "File priority - (.*)",
        'Task Priority: (.*)',
        'Task Level: (.*)',
        'Task URL: (https://.*)',
        'Priority: (.*)',
        'T. Level: (.*)',
        'href=(https://.*)'
      ]
    }
    /**
     * Internal prefixes
     */
    this.prefixes = prefixes || {
      /**
       * base url for internal links
       */
      baseURL: 'https://support.business.com/ticket/'
    }
    return this
  }
}
