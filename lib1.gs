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



// https://docs.google.com/document/d/1qdC6s7Jgt1F8xauXyEu_khlTBYh8IExpNQaLMu_NPig/edit?usp=sharing

// https://gist.github.com/mhawksey/1170597
 

function ExcelDateToJSDate(serial) {
   var utc_days  = Math.floor(serial - 25569);
   var utc_value = utc_days * 86400;                                        
   var date_info = new Date(utc_value * 1000);

   var fractional_day = serial - Math.floor(serial) + 0.0000001;

   var total_seconds = Math.floor(86400 * fractional_day);

   var seconds = total_seconds % 60;

   total_seconds -= seconds;

   var hours = Math.floor(total_seconds / (60 * 60));
   var minutes = Math.floor(total_seconds / 60) % 60;

   return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}


  // A4 in points = 595pt x 842pt
  // 1 cm = 28.346456693 points

        /*
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
        */
  /*
  var newTable = TemplateCopyBody.appendTable(myTable);
  for(var i=1; i < myTable.length ;i++){
    if(i % 2 == 0)
      newTable.getRow(i).setMinimumHeight(22);
    else
      newTable.getRow(i).setMinimumHeight(45);
  }
  newTable.setAttributes(tableStyle).setColumnWidth(0, 300);
  */
  
  //Logger.log(JSON.stringify(data))
  
  /*
  for(n=0 ; n<data.length;++n){
    var row = data[n];
    for(var i=0; i<row.length; i++) {
      Logger.log(typeof data[n][i]); //  **** #### typeof
      if (row[i] instanceof Date) {
        data[n][i] = "'"+Utilities.formatDate(row[i], SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yy");// is a string
        // added a ' before the string to ensure the spreadsheet won't convert it back automatically
        Logger.log('Array cell '+n+','+i+' modified to '+data[n][i]);
      }          
    }
  }
  */
