function createDocFromSheet3(){
  var templateDocID = "10OlSE9c8__vLaqPydUnZQbSqNvQmxvogttx-wsWiVsE"; // get template file id - Water Statement
  var FOLDER_NAME = "GDK"; // folder name of where to put completed reports
  var FOLDER_ID = "0B6NHem9C-Di5XzlfVGRzRzVtbU0"; // folder ID of where to put completed reports
  var WATER_DATA = "Meter Readings"; // name of sheet with water meter readings
  var DOC_PREFIX = "Water Statement - "; // prefix for name of document to be loaded with water advice data
  var DUMMY_PARA = "Remove"; // Text denoting a dummy or unwanted paragraph
  var WS_TABLE = "Watering Schedule";  // Text as a place mark for the Water Scheduling table
  var START_ROW = 3; // The row on which the data in the spreadsheet starts
  var START_COL = 1; // The column on which the data in the spreadsheet starts
  
  // get the data for the statements
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(WATER_DATA);
  var data = sheet.getRange(START_ROW, START_COL, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
  var sheet = ss.getSheetByName("Data");
  var data2 = sheet.getRange(2, 1, 13, 13).getValues();
  
  // create new document
  var adviceNbr = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd"); // get watering number and date
  var doc = DocumentApp.create(DOC_PREFIX+adviceNbr);
  var body = doc.getBody();

  // move file to right folder
  //var file = DocsList.getFileById(doc.getId());
  //var folder = DocsList.getFolder(FOLDER_NAME);
  //file.addToFolder(folder);
  //file.removeFromFolder(DocsList.getRootFolder());
  
  // Get the body of the template document
  var bodyCopy = DocumentApp.openById(templateDocID).getBody();
  body.setMarginTop(bodyCopy.getMarginTop());
  body.setMarginBottom(bodyCopy.getMarginBottom());

  // for each water user fill in the template with the data 
  for (var i in data2){
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
    newBody.replaceText("<<User>>", data2[i][2]);
    newBody.replaceText("<<Address>>", data2[i][1]);
    newBody.replaceText("<<watering_no>>", "29/04/2015");
    newBody.replaceText("<<Allocation>>", Utilities.formatString('%11.1f', data2[i][10]));
    newBody.replaceText("<<UTD>>", Utilities.formatString('%11.1f', data2[i][11]));
    newBody.replaceText("<<Remain>>", Utilities.formatString('%11.1f', data2[i][12]));
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
          addTableInDocument2(doc, data, tz, data2[i][0]);
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
  ss.toast("Water Statements have been compiled");
}


// http://www.googleappsscript.org/home/create-table-in-google-document-using-apps-script

function addTableInDocument2(docBody, dataTable, tz, user_no) {
  //define header cell style
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#d9d9d9';
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  headerStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  
  //Style for the cells other than header row
  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.BOLD] = false;
  cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  cellStyle[DocumentApp.Attribute.FONT_SIZE] = 10;

  // paragraph style
  var paraStyle = {};
  paraStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
  paraStyle[DocumentApp.Attribute.LINE_SPACING] = 1;
  
  // Centre the table
  var tstyle = {};
  tstyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
     DocumentApp.HorizontalAlignment.CENTER;
  
  //Add a table in document
  var table = docBody.appendTable();
  // Put header row
  var tr = table.appendTableRow();
  var td = tr.appendTableCell('Watering No.');
  td.setAttributes(headerStyle);
  var td = tr.appendTableCell('Start Time');
  td.setAttributes(headerStyle);
  var td = tr.appendTableCell('Finish Time');
  td.setAttributes(headerStyle);
  var td = tr.appendTableCell('Meter Start');
  td.setAttributes(headerStyle);
  var td = tr.appendTableCell('Meter Finish');
  td.setAttributes(headerStyle);
  var td = tr.appendTableCell('Water Used');
  td.setAttributes(headerStyle);
  table.setBorderColor("#cccccc");
  table.setColumnWidth(0, 65); //WIDTH:111
  table.setColumnWidth(1, 140); //WIDTH:70
  table.setColumnWidth(2, 140); //WIDTH:159
  table.setColumnWidth(3, 40); //WIDTH:159
  table.setColumnWidth(4, 40); //WIDTH:159
  table.setColumnWidth(5, 40); //WIDTH:159
  table.setAttributes(tstyle);

  // Load schedule
  for (var i in dataTable){
    var wused = Number(dataTable[i][12]);
    if (isNaN(wused)) {
        var dmp = false;
    } else {
        dmp = (wused > 0) ? true : false;
    }
    if(dataTable[i][2] == user_no && dmp) {
      var dRow = dataTable[i];
      var tr = table.appendTableRow();
      var td = tr.appendTableCell(dRow[0]);
      td.setAttributes(cellStyle);
      //Apply the para style to each paragraph in cell
      var paraInCell = td.getChild(0).asParagraph();
      paraInCell.setAttributes(paraStyle);
      var td = tr.appendTableCell(Utilities.formatDate(ExcelDateToJSDate(dRow[5]), tz, "dd/MM/yyyy hh:mm a"));
      td.setAttributes(cellStyle);
      //Apply the para style to each paragraph in cell
      var paraInCell = td.getChild(0).asParagraph();
      paraInCell.setAttributes(paraStyle);
      var td = tr.appendTableCell(Utilities.formatDate(ExcelDateToJSDate(dRow[7]), tz, "dd/MM/yyyy hh:mm a"));
      td.setAttributes(cellStyle);
      //Apply the para style to each paragraph in cell
      var paraInCell = td.getChild(0).asParagraph();
      paraInCell.setAttributes(paraStyle);
      var td = tr.appendTableCell(Utilities.formatString('%11.1f', dRow[10]));
      td.setAttributes(cellStyle);
      //Apply the para style to each paragraph in cell
      var paraInCell = td.getChild(0).asParagraph();
      paraInCell.setAttributes(paraStyle);
      var td = tr.appendTableCell(Utilities.formatString('%11.1f', dRow[11]));
      td.setAttributes(cellStyle);
      //Apply the para style to each paragraph in cell
      var paraInCell = td.getChild(0).asParagraph();
      paraInCell.setAttributes(paraStyle);
      var td = tr.appendTableCell(Utilities.formatString('%11.1f', dRow[12]));
      td.setAttributes(cellStyle);
      //Apply the para style to each paragraph in cell
      var paraInCell = td.getChild(0).asParagraph();
      paraInCell.setAttributes(paraStyle);
    }
  }
}
