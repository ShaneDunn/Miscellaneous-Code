/**
 * Add a custom menu to the active spreadsheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('API Table Menu')
      .addItem('Client Tables', 'createDoc01')
      .addItem('External Tables', 'createDoc02')
      .addToUi();
}

function createDoc01(){
  var templateDocID = "13Wsgk7jmjXtjfyNorlX_OlBJLZoy3u8PuRarozETwOo"; // get template file id - DBW Scheduling API - tables template
  var API_DATA = "Sheet1"; // name of sheet with api data
  var DOC_PREFIX = "APIs - "; // prefix for name of document to be loaded with api tables
  var START_ROW = 2; // The row on which the data in the spreadsheet starts
  var START_COL = 1; // The column on which the data in the spreadsheet starts

  // get the data for the api's
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(API_DATA);
  var data = sheet.getRange(START_ROW, START_COL, sheet.getLastRow()-(START_ROW-1), sheet.getLastColumn()).getValues();

  // create new document
  var docNbr = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd-HH:mm:ss"); // get date and time
  var doc = DocumentApp.create(DOC_PREFIX+docNbr);
  var body = doc.getBody();

  addTableInDocument(body, data, tz);
}

function createDoc02(){
  var templateDocID = "13Wsgk7jmjXtjfyNorlX_OlBJLZoy3u8PuRarozETwOo"; // get template file id - DBW Scheduling API - tables template
  var API_DATA = "Sheet1"; // name of sheet with api data
  var DOC_PREFIX = "APIs - "; // prefix for name of document to be loaded with api tables
  var START_ROW = 2; // The row on which the data in the spreadsheet starts
  var START_COL = 1; // The column on which the data in the spreadsheet starts

  // get the data for the api's
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(API_DATA);
  var data = sheet.getRange(START_ROW, START_COL, sheet.getLastRow()-(START_ROW-1), sheet.getLastColumn()).getValues();

  // create new document
  var docNbr = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd-HH:mm:ss"); // get date and time
  var doc = DocumentApp.create(DOC_PREFIX+docNbr);
  var body = doc.getBody();

  addTableInDocument(body, data, tz);
  addTableInDocument2(body, data, tz);
}

  // move file to right folder
  //var file = DocsList.getFileById(doc.getId());
  //var folder = DocsList.getFolder(FOLDER_NAME);
  //file.addToFolder(folder);
  //file.removeFromFolder(DocsList.getRootFolder());
  /*
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var file = DriveApp.getFileById(templateDocID).makeCopy(DOC_PREFIX+adviceNbr, folder);
  var docID = file.getId();
  var doc = DocumentApp.openById(docID);
  var body = doc.getBody();
  var bodyCopy = doc.getBody().copy();
  body = body.clear();
  */

  // Get the body of the template document
  //var bodyCopy = DocumentApp.openById(templateDocID).getBody();
  //body.setMarginTop(bodyCopy.getMarginTop());
  //body.setMarginBottom(bodyCopy.getMarginBottom());

  // for each water user's entry fill in the template with the data 
  /*
  for (var i in data){
    // Put in a page break between each user, but only after the first one
    if( i > 0) {
      var pgBrk = body.appendPageBreak();
    }
    // Format dates - check if a date object or a excel/calc decimal date number
    if (data[i][9] instanceof Date) {
      var temp = data[i][9];
    } else {
      var temp = ExcelDateToJSDate(data[i][10]);
    }
    var start_date = Utilities.formatDate(temp, tz, "EEEE dd/MM/yyyy hh:mm a");

    if (data[i][10] instanceof Date) {
      var temp = data[i][10];
    } else {
      var temp = ExcelDateToJSDate(data[i][11]);
    }
    var end_date = Utilities.formatDate(temp, tz, "EEEE dd/MM/yyyy hh:mm a");
    var addTable = true;
    // load template and replace tokens
    var newBody = bodyCopy.copy();
    newBody.replaceText("<<User>>", data[i][2]);
    newBody.replaceText("<<Address>>", data[i][19]);
    newBody.replaceText("<<watering_no>>", data[0][0]);
    newBody.replaceText("<<sDate>> <<sTime>> <<sPeriod>>",start_date + " [" + data[i][12]+"]");
    newBody.replaceText("<<eDate>> <<eTime>> <<ePeriod>>",end_date + " [" + data[i][13]+"]");
    newBody.replaceText("<<Hrs>>", data[i][15]);
    newBody.replaceText("<<Delivery Rate>>", data[i][14]);
    newBody.replaceText("<<UTD>>", Utilities.formatString('%11.1f', data[i][8]));
    newBody.replaceText("<<eUsage>>", Utilities.formatString('%11.1f', data[i][17]));
    newBody.replaceText("<<Remain>>", Utilities.formatString('%11.1f', data[i][9]));
    newBody.replaceText("<<eRemain>>", Utilities.formatString('%11.1f', data[i][18]));
    // append template to new document
    for (var j = 0; j < newBody.getNumChildren(); j++) {
      var element = newBody.getChild(j).copy();
      var type = element.getType(); // need to handle different types para, table etc differently
      //Logger.log("Element type is "+type);
      if (type == DocumentApp.ElementType.PARAGRAPH ) {
        if (element.asParagraph().getText() != DUMMY_PARA) {
          body.appendParagraph(element);
        }
        if (element.asParagraph().getText() == WS_TABLE ) {
          addTableInDocument(doc, data, tz);
          addTable = false;
        }
      } else if (type == DocumentApp.ElementType.TABLE ) {
        if ( addTable ) { body.appendTable(element); }
        else { addTable = true; }
      } else if( type == DocumentApp.ElementType.LIST_ITEM ) {
        body.appendListItem(element);
      } else
        throw new Error("Unknown element type: "+type);
    }
    // remove first blank line / paragraph
    if( i == 0) {
      var para = body.getChild(0).removeFromParent();
    }
  }
  doc.saveAndClose();
  ss.toast("Water Delivery Advices have been compiled");
}
*/
