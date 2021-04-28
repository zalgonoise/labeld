/**
 * Database class will handle the interactions needed with Google Sheets
 * such as setting up a Sheet for a user, checking content, and adding content
 */
class Database {
  constructor() {
    this.user = Session.getActiveUser();
    this.file;
    this.sheet;
    this.blankRow;
    this.lastValue;

    /**
     * GetSheet method will define the file property, and look if the user has a 
     * named sheet already; creating one if otherwise (while defining the sheet property)
     */
    this.GetSheet = function() {
      this.file = SpreadsheetApp.getActiveSpreadsheet();
      var namedSheets = this.file.getSheets();

      for (var i = 0 ; i < namedSheets.length; i++) {
        if (namedSheets[i].getSheetName() == this.user) {

          this.sheet = this.file.getSheets()[i]
          if (this.sheet.getRange("A1").getValue() === "") {
            this.SetupSheet()
            return
          }
          break
        } 
      }

      if (!this.sheet) {
        this.sheet = this.file.insertSheet()
        this.sheet.setName(this.user)
        this.SetupSheet()
        Logger.log(`Created new sheet for user: ${this.user}`)
      }      
    }

    /**
     * SetupSheet method will define the Sheet's headers in case a new user is setting up
     * a new instance of Database / LabelD
     */
    this.SetupSheet = function() {
      function setHeader(sheet, range, value) {
        if (sheet.getRange(range).getValue() != value) {
          sheet.getRange(range).setValue(value)
        }
      }

      setHeader(this.sheet, "A1", "From")
      setHeader(this.sheet, "B1", "To")
      setHeader(this.sheet, "C1", "Snippet")
      setHeader(this.sheet, "D1", "Task Type")
      setHeader(this.sheet, "E1", "Task Source")
      setHeader(this.sheet, "F1", "Time")
      setHeader(this.sheet, "G1", "Message")
      setHeader(this.sheet, "H1", "ID")
      setHeader(this.sheet, "I1", "Unix timestamp")
      setHeader(this.sheet, "J1", "Reference ID")
      setHeader(this.sheet, "K1", "Duplicate")
      setHeader(this.sheet, "L1", "Secondary Reference ID")
      setHeader(this.sheet, "M1", "Reference URL")        
      setHeader(this.sheet, "N1", "Priority")
      setHeader(this.sheet, "O1", "Service Level")
      Logger.log("Initialized spreadsheet's headers")      
    }

    /**
     * RepairLastEntry method will look through the latest row inspecting if
     * any interrupted batch processes left a row half-populated; removing its entire content
     * prior to pushing new entries on top of it 
     */
    this.RepairLastEntry = function() {
      var range = `A${(this.blankRow - 1)}:L${this.blankRow}`
      var cells = this.sheet.getRange(range).getValues();
      for (var a = 0 ; a < cells.length ; a++) {
        for (var b = 0 ; b < cells[a].length ; b++) {
          if (cells[a][b] === "" ) {
            this.sheet.getRange(range).setValue("")
            return true
          }
        }
      }
      return false
    }

    /**
     * RemoveEntry method will delete a single row in Sheets,
     * as per provided index value
     * 
     * @param {int} index index - the Sheets row value to remove
     */
    this.RemoveEntry = function(index) {
      var range = `A${index}:O${index}`
      this.sheet.getRange(range).setValue("");
    }

    /**
     * RemoveBelow method will delete all row in Sheets below the
     * provided index value
     * 
     * @param {int} index index - the Sheets row value to reference a 
     * removal point
     */
    this.RemoveBelow = function(index) {
      var range = `A${index}:O99999`
      this.sheet.getRange(range).setValue("")
    }    
    
    /**
     * LatestEntry method will look through this user's sheet, and retrieve both the last empty 
     * cell's index, and the latest value from the Sheet's unix time column
     */
    this.LatestEntry = function() {
      // Getting the latest value present in the sheet 
      // by looking through all the Unix Timestamp cells
      // and storing the last value
      var range = "I2:I50000"
      var cells = this.sheet.getRange(range).getValues();
      
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
          //var prevLastValue = lastValue
          var lastValue = cells[i][0]
        }
      }

      // In case there are no entries, all messages are fetched
      if (!lastValue) {
        var lastValue = 0
      } 
      
      this.blankRow = blankRow;
      this.lastValue = lastValue;

      //if (this.RepairLastEntry()) {
      //  this.blankRow = (this.blankRow - 1)
      //  this.lastValue = prevLastValue
      //}


    }

    /**
     * CheckBacklog method will find references in the existing Sheets backlog
     * to match when its time to push the events into Sheets, marking as duplicate
     * those events that appear as repeated
     * 
     * @param {int} startRow startRow - the initial point (row) where to start 
     * picking up references
     * 
     * @param {int} numRows numRows - how many rows to include in the backlog search
     */
    this.CheckBacklog = function(startRow, numRows) {
      // define columns to pick data from
      var sheetCols = [
        "I",
        "J",
        "L"
      ]

      var block = []; 

      // iterate through each column for the set range, to retrieve its values
      for (var a = 0 ; a < sheetCols.length ; a++ ) {
        var array = [];
        var range = sheetCols[a] + (startRow - (numRows - 1)) + ":" + sheetCols[a] + startRow
        var cells = this.sheet.getRange(range).getValues();

        // iterate through each value, and push it to a temporary array / list
        for (var b = 0 ; b < cells.length ; b++) {
          if (cells[b][0] === "") {
            break      
          } else {
            array.push(cells[b][0])
          }
        }
        
        // push each temporary array into a block (or, array of arrays) like a map
        block.push(array)
      }

      return block
    }

    /**
     * PushEntry method will take in a Message object and populate the Sheet's 
     * latest blankRow with it
     * 
     * @param {Object} obj obj - a Message object, generated by MessageBuilder
     */    
    this.PushEntry = function(obj) {
      this.sheet.getRange(this.blankRow, 1).setValue(obj.output.sender);
      this.sheet.getRange(this.blankRow, 2).setValue(obj.output.to);
      this.sheet.getRange(this.blankRow, 3).setValue(obj.output.snippet);
      this.sheet.getRange(this.blankRow, 4).setValue(obj.output.type);
      this.sheet.getRange(this.blankRow, 5).setValue(obj.output.source);
      this.sheet.getRange(this.blankRow, 6).setValue(obj.output.time);
      this.sheet.getRange(this.blankRow, 6).setNumberFormat("dd/MM/yyyy HH:MM:SS");
      this.sheet.getRange(this.blankRow, 7).setValue(obj.output.subj);
      this.sheet.getRange(this.blankRow, 8).setValue(obj.output.id);
      this.sheet.getRange(this.blankRow, 9).setValue(obj.output.unix);
      this.sheet.getRange(this.blankRow, 9).setNumberFormat("0000000000000");
      this.sheet.getRange(this.blankRow, 10).setValue(obj.output.ref);
      this.sheet.getRange(this.blankRow, 10).setNumberFormat("00000000");
      this.sheet.getRange(this.blankRow, 11).setValue(obj.output.dup);
      this.sheet.getRange(this.blankRow, 12).setValue(obj.output.bodyRef);
      this.sheet.getRange(this.blankRow, 13).setValue(obj.output.bodyRefURL);     
      this.sheet.getRange(this.blankRow, 14).setValue(obj.output.priority); 
      this.sheet.getRange(this.blankRow, 15).setValue(obj.output.level);        
    }

    /**
     * GetEntry method will take in an integer value representing the intended index to
     * retrieve, and return this user's Sheets content.
     * 
     * @param {int} index index - the Sheets document index (row) to retrieve data from
     */
    this.GetEntry = function(index) {
      return this.sheet.getRange(`I${index}`).getValue()
    }

    /**
     * IncrementRow method will increment 1 to the blankRow property
     */
    this.IncrementRow = function() {
      this.blankRow = (this.blankRow + 1)
    }

    /**
     * Database runtime
     */
    this.GetSheet()
  }
}
