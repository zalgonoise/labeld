/**
 * Labeler class is part of the Mailbox class, as it implements it.
 * Labeler is targetted at performing the "more-background" tasks for 
 * Gmail which don't involve so much fetching and processing messages,
 * but filters and labels.
 */
class Labeler {

  /**
   * @param {string} label label - the string representation of the Gmail 
   * label
   */
  constructor(label) {
    this.label = label || ""
    this.labelID
    this.filters

    /**
     * SetLabelID method will fetch the user's labels in Gmail,
     * and find the configured label's ID by its name.
     * 
     * If the defined label name does not exist yet, this method will
     * create one, and store its ID.
     */
    this.SetLabelID = function() {
      var userLabels = Gmail.Users.Labels.list("me")
      if (userLabels.labels.find(x => x.name === this.label)) {
        this.labelID = userLabels.labels.find(x => x.name === this.label).id
        return
      } else {
        Gmail.Users.Labels.create(
          {
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
            "name": this.label
          },
          'me'
        )
        var userLabels = Gmail.Users.Labels.list("me")
        this.labelID = userLabels.labels.find(x => x.name === this.label).id
      }
    }

    /**
     * SetFilters method will take in a Gmail user.filters.get response and 
     * store its filter list as the filters item in the Labeler object
     * 
     * @param {Object[]} obj obj - The response from a Gmail user.filters.get request
     */
    this.SetFilters = function(obj) { 
      this.filters = obj.filter
    }

    /**
     * CheckFilters method will iterate through the configured filters, looking for one
     * which adds the configured label ID. 
     * 
     * While it returns null if the filter already exists, it will return the label ID 
     * as a string if it doesn't, for it to be created
     * 
     * @returns {string} labelID
     */
    this.CheckFilters = function() {

      if (this.filters.length > 0) {
        for (var i = 0 ; i < this.filters.length ; i++) {
          if ( 
              (this.filters[i].action) && 
              (this.filters[i].action.addLabelIds) && 
              this.filters[i].action.addLabelIds.length > 0
            ) {
              for (var x = 0 ; x < this.filters[i].action.addLabelIds.length ; x++) {
                if (this.filters[i].action.addLabelIds[x] == this.labelID) {
                  return null
                }
              }
            }

        }
        return this.labelID
      }

    }

    /**
     * GetFilterID method will iterate through the configured filters, looking for one
     * which adds the configured label ID. 
     * 
     * It returns the Filter ID once there is a match, namely used to remove the actual filter
     * 
     * @returns {string} filterID
     */
    this.GetFilterID = function() {
      if (this.filters.length > 0) {
        for (var i = 0 ; i < this.filters.length ; i++) {
          if ( 
              (this.filters[i].action) && 
              (this.filters[i].action.addLabelIds) && 
              this.filters[i].action.addLabelIds.length > 0
            ) {
              for (var x = 0 ; x < this.filters[i].action.addLabelIds.length ; x++) {
                if (this.filters[i].action.addLabelIds[x] == this.labelID) {
                  return this.filters[i].id
                }
              }
            }

        }
      }
      return null
    }
  }
}
