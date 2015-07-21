function logAttributes() {
  /* https://docs.google.com/a/debortoli.com.au/document/d/1m1zIBAok3HLgSoTUtuguAumFTE97wPwE7UzBJWMgFhA/edit?usp=sharing */
  var templateDocID = "1m1zIBAok3HLgSoTUtuguAumFTE97wPwE7UzBJWMgFhA"; // get template file id - DBW Scheduling API - tables template
  var file = DriveApp.getFileById(templateDocID);
  var docID = file.getId();
  var doc = DocumentApp.openById(docID);
  var body = doc.getBody();
  
  for (var j = 0; j < body.getNumChildren(); j++) {
    var element = body.getChild(j);
    var type = element.getType(); // need to handle different types para, table etc differently
    //Logger.log("Element type is "+type);
    if (type == DocumentApp.ElementType.PARAGRAPH ) {
      // Retrieve the paragraph's attributes.
      var atts = element.getAttributes();
      // Log the paragraph attributes.
      Logger.log("##PARA");
      for (var att in atts) {
        Logger.log(att + ":" + atts[att]);
      }  
    } else if (type == DocumentApp.ElementType.TABLE ) {
      Logger.log("##TABLE");
      // Retrieve the tables attributes.
      var atts = element.getAttributes();
      // Log the table attributes.
      for (var att in atts) {
        Logger.log(att + ":" + atts[att]);
      }
      Logger.log("##ROWS");
      for (var ii = 0; ii < element.asTable().getNumRows(); ii++){
        var rw = element.asTable().getRow(ii);
        // Retrieve the table row attributes.
        var atts = rw.getAttributes();
        // Log the row attributes.
        for (var att in atts) {
          Logger.log(att + ":" + atts[att]);
        }
      }
      Logger.log("##CELLS");
      for (var ii = 0; ii < element.asTable().getNumRows(); ii++){
        var rw = element.asTable().getRow(ii);
        for (var jj = 0; jj < rw.getNumCells(); jj++ ) {
          var cl = rw.getCell(jj);
          // Retrieve the table cell attributes.
          var atts = cl.getAttributes();
          // Log the cell attributes.
          for (var att in atts) {
            Logger.log(att + ":" + atts[att]);
          }
        }
      }
    } else if( type == DocumentApp.ElementType.LIST_ITEM ) {
      // Retrieve the list's attributes.
      var atts = element.getAttributes();
      // Log the list attributes.
      Logger.log("##LIST");
      for (var att in atts) {
        Logger.log(att + ":" + atts[att]);
      }  
    } else if( type == DocumentApp.ElementType.TABLE_OF_CONTENTS ) {
      // Retrieve the list's attributes.
      var atts = element.getAttributes();
      // Log the list attributes.
      Logger.log("##TOC");
      for (var att in atts) {
        Logger.log(att + ":" + atts[att]);
      }  
    } else
      throw new Error("Unknown element type: "+type);
  }
}
