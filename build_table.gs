// http://www.googleappsscript.org/home/create-table-in-google-document-using-apps-script

function fillCell (cell, style, str) {

  var paraStyle1 = {};
  paraStyle1[DocumentApp.Attribute.FONT_FAMILY] = 'Consolas';
  paraStyle1[DocumentApp.Attribute.FOREGROUND_COLOR] = '#990000';
  paraStyle1[DocumentApp.Attribute.BOLD] = false;
  var paraStyle2 = {};
  paraStyle2[DocumentApp.Attribute.FONT_FAMILY] = 'Consolas';
  paraStyle2[DocumentApp.Attribute.FOREGROUND_COLOR] = '#38761d';
  paraStyle2[DocumentApp.Attribute.BOLD] = false;
  var paraStyle3 = {};
  paraStyle3[DocumentApp.Attribute.FONT_FAMILY] = 'Consolas';
  paraStyle3[DocumentApp.Attribute.FOREGROUND_COLOR] = null;
  paraStyle3[DocumentApp.Attribute.BOLD] = false;

  var searchType = DocumentApp.ElementType.PARAGRAPH;
  var searchResult = null;

  while (searchResult = cell.findElement(searchType)) {
    var par = searchResult.getElement().asParagraph();
    par.setText(str);
    var text = par.editAsText();
    if ( style == 1 ) {text.setAttributes(paraStyle1);}     // Set red mono font Attributes
    if ( style == 2 ) {text.setAttributes(paraStyle2);}     // Set green mono font Attributes
    if ( style == 3 ) {text.setAttributes(paraStyle3);}     // Set normal mono font Attributes
    if ( style == 4 ) {text.setAttributes(paraStyle3); text.setForegroundColor(2, 6, '#79aae0');}     // Set normal mono font Attributes
    return;
  }
}

function addTableInDocument(docBody, dataTable, tz) {
  // Define the styles for table elements
  
  // Table style
  var tblStyle = {};
  tblStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
     DocumentApp.HorizontalAlignment.CENTER;
  
  // Table Row style - header
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#cfe2f3';
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  //headerStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  
  // Table Row style - normal
  var rowStyle = {};
  rowStyle[DocumentApp.Attribute.BOLD] = false;
  rowStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = null;
 // rowStyle[DocumentApp.Attribute.FONT_SIZE] = 10;

  // Table Cell style
  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.BOLD] = false;
  cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = null;
  //cellStyle[DocumentApp.Attribute.FONT_SIZE] = 10;

  var para = docBody.appendParagraph(""); // Append a regular paragraph - for the description
  var tpara = para.copy();
  tpara.clear();
  //para.removeFromParent();

  // Load api's
  for (var i in dataTable){
    var dRow = dataTable[i];
    
    if (dRow[1] == "") {
      // new API
      var header = docBody.appendParagraph("API - " + dRow[2]); // Append a document header paragraph.
      header.setHeading(DocumentApp.ParagraphHeading.HEADING4);
      if (dRow[4] == "") {
        docBody.appendParagraph("Something describing what this does."); // Append a regular paragraph - for the description
        docBody.appendParagraph("Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Donec odio. Quisque volutpat mattis eros. Nullam malesuada erat ut turpis. Suspendisse urna nibh, viverra non, semper suscipit, posuere a, pede.");
      } else {
        para = docBody.appendParagraph(dRow[4]);
        para.setForegroundColor(dRow[5]);
      }
    } else {
      var section = docBody.appendParagraph("Request"); // Append a section header paragraph.
      section.setHeading(DocumentApp.ParagraphHeading.HEADING5);

      var table = docBody.appendTable(); // Append Table
      var tr = table.appendTableRow();   // Append Row
      var td = tr.appendTableCell('Method');
      td.setWidth(54);
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      td = tr.appendTableCell('URL');
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      tr = table.appendTableRow();   // Append Row
      td = tr.appendTableCell(dRow[1]);
      td.setWidth(54);
      td.setAttributes(cellStyle);     // Set Normal Row Attributes
      td = tr.appendTableCell(dRow[2]);
      td.setAttributes(cellStyle);     // Set Normal Row Attributes
      
      docBody.appendParagraph(""); // Append a blank paragraph

      table = docBody.appendTable(); // Append Table ####
      tr = table.appendTableRow();   // Append Row ##
      td = tr.appendTableCell('Type');
      td.setWidth(54);
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      td = tr.appendTableCell('Params');
      td.setWidth(156.25);
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      td = tr.appendTableCell('Values');
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      tr = table.appendTableRow();   // Append Row ##
      if (dRow[4] == "") {
        td = tr.appendTableCell('HEAD');
        td.setWidth(54);
        td .setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell('Cookie: connect.sid=<sid>');
        td.setWidth(156.25);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell('String');
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        tr = table.appendTableRow();   // Append Row ##
        td = tr.appendTableCell(dRow[1]);
        td.setWidth(54);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell('Parameter');
        td.setWidth(156.25);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell('DefaultValue|Type');
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
      } else {
        tr = table.appendTableRow();   // Append Row ##
        td = tr.appendTableCell(dRow[1]);
        td.setWidth(54);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell(dRow[4]);
        td.setWidth(156.25);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell(dRow[5]);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
      }
      
      docBody.appendParagraph(""); // Append a blank paragraph

      section = docBody.appendParagraph("Response"); // Append a section header paragraph.
      section.setHeading(DocumentApp.ParagraphHeading.HEADING5);

      table = docBody.appendTable(); // Append Table
      tr = table.appendTableRow();   // Append Row
      td = tr.appendTableCell('Status');
      td.setWidth(54);
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      td = tr.appendTableCell('Response');
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      tr = table.appendTableRow();   // Append Row
      td = tr.appendTableCell();
      td.setWidth(54);
      fillCell(td, 2, "200");
      td = tr.appendTableCell();
      if (dRow[6] == "") { 
        fillCell(td, 3, "//");
      } else {
        fillCell(td, 3, dRow[6]);
      }
      tr = table.appendTableRow();   // Append Row
      td = tr.appendTableCell();
      td.setWidth(54);
      fillCell(td, 1, "500");
      td = tr.appendTableCell();
      fillCell(td, 4, '{"error":"Something went wrong. Please try again later."}');

      docBody.appendParagraph(""); // Append a blank paragraph 

    }
  }
}


function addTableInDocument(docBody, dataTable, tz) {
  // Define the styles for table elements
  
  // Table style
  var tblStyle = {};
  tblStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
     DocumentApp.HorizontalAlignment.CENTER;
  
  // Table Row style - header
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#cfe2f3';
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  //headerStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
  
  // Table Row style - normal
  var rowStyle = {};
  rowStyle[DocumentApp.Attribute.BOLD] = false;
  rowStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = null;
 // rowStyle[DocumentApp.Attribute.FONT_SIZE] = 10;

  // Table Cell style
  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.BOLD] = false;
  cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = null;
  //cellStyle[DocumentApp.Attribute.FONT_SIZE] = 10;

  var para = docBody.appendParagraph(""); // Append a regular paragraph - for the description
  var tpara = para.copy();
  tpara.clear();
  //para.removeFromParent();

  // Load api's
  for (var i in dataTable){
    var dRow = dataTable[i];
    
    if (dRow[1] == "") {
      // new API
      var header = docBody.appendParagraph("API - " + dRow[2]); // Append a document header paragraph.
      header.setHeading(DocumentApp.ParagraphHeading.HEADING4);
      if (dRow[4] == "") {
        docBody.appendParagraph("Something describing what this does."); // Append a regular paragraph - for the description
        docBody.appendParagraph("Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Donec odio. Quisque volutpat mattis eros. Nullam malesuada erat ut turpis. Suspendisse urna nibh, viverra non, semper suscipit, posuere a, pede.");
      } else {
        para = docBody.appendParagraph(dRow[4]);
        para.setForegroundColor(dRow[5]);
      }
    } else {
      var section = docBody.appendParagraph("Request"); // Append a section header paragraph.
      section.setHeading(DocumentApp.ParagraphHeading.HEADING5);

      var table = docBody.appendTable(); // Append Table
      var tr = table.appendTableRow();   // Append Row
      var td = tr.appendTableCell('Method');
      td.setWidth(54);
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      td = tr.appendTableCell('URL');
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      tr = table.appendTableRow();   // Append Row
      td = tr.appendTableCell(dRow[1]);
      td.setWidth(54);
      td.setAttributes(cellStyle);     // Set Normal Row Attributes
      td = tr.appendTableCell(dRow[2]);
      td.setAttributes(cellStyle);     // Set Normal Row Attributes
      
      docBody.appendParagraph(""); // Append a blank paragraph

      table = docBody.appendTable(); // Append Table ####
      tr = table.appendTableRow();   // Append Row ##
      td = tr.appendTableCell('Type');
      td.setWidth(54);
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      td = tr.appendTableCell('Params');
      td.setWidth(156.25);
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      td = tr.appendTableCell('Values');
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      tr = table.appendTableRow();   // Append Row ##
      if (dRow[4] == "") {
        td = tr.appendTableCell('HEAD');
        td.setWidth(54);
        td .setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell('Cookie: connect.sid=<sid>');
        td.setWidth(156.25);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell('String');
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        tr = table.appendTableRow();   // Append Row ##
        td = tr.appendTableCell(dRow[1]);
        td.setWidth(54);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell('Parameter');
        td.setWidth(156.25);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell('DefaultValue|Type');
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
      } else {
        tr = table.appendTableRow();   // Append Row ##
        td = tr.appendTableCell(dRow[1]);
        td.setWidth(54);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell(dRow[4]);
        td.setWidth(156.25);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
        td = tr.appendTableCell(dRow[5]);
        td.setAttributes(cellStyle);     // Set Normal Row Attributes
      }
      
      docBody.appendParagraph(""); // Append a blank paragraph

      section = docBody.appendParagraph("Response"); // Append a section header paragraph.
      section.setHeading(DocumentApp.ParagraphHeading.HEADING5);

      table = docBody.appendTable(); // Append Table
      tr = table.appendTableRow();   // Append Row
      td = tr.appendTableCell('Status');
      td.setWidth(54);
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      td = tr.appendTableCell('Response');
      td.setAttributes(headerStyle);     // Set Header Row Attributes
      tr = table.appendTableRow();   // Append Row
      td = tr.appendTableCell();
      td.setWidth(54);
      fillCell(td, 2, "200");
      td = tr.appendTableCell();
      if (dRow[6] == "") { 
        fillCell(td, 3, "//");
      } else {
        fillCell(td, 3, dRow[6]);
      }
      tr = table.appendTableRow();   // Append Row
      td = tr.appendTableCell();
      td.setWidth(54);
      fillCell(td, 1, "500");
      td = tr.appendTableCell();
      fillCell(td, 4, '{"error":"Something went wrong. Please try again later."}');

      docBody.appendParagraph(""); // Append a blank paragraph 

    }
  }
}


/*
  // Put header row
  var td = tr.appendTableCell('Start');
  td.setAttributes(headerStyle);
  var td = tr.appendTableCell('Finish');
  td.setAttributes(headerStyle);
  table.setBorderColor("#cccccc");
  table.setColumnWidth(0, 65); //WIDTH:111
  table.setColumnWidth(1, 65); //WIDTH:70
  table.setColumnWidth(2, 160); //WIDTH:159
  table.setColumnWidth(3, 160); //WIDTH:159
  table.setAttributes(tstyle);

    
    var tr = table.appendTableRow();
    var td = tr.appendTableCell(dRow[2]);
    td.setAttributes(cellStyle);
    //Apply the para style to each paragraph in cell
    var paraInCell = td.getChild(0).asParagraph();
    paraInCell.setAttributes(paraStyle);
    var td = tr.appendTableCell(dRow[16] + ' / ' + dRow[14]);
    td.setAttributes(cellStyle);
    //Apply the para style to each paragraph in cell
    var paraInCell = td.getChild(0).asParagraph();
    paraInCell.setAttributes(paraStyle);
    var td = tr.appendTableCell(Utilities.formatDate(ExcelDateToJSDate(dRow[10]), tz, "EEEE dd/MM/yyyy hh:mm a"));
    td.setAttributes(cellStyle);
    //Apply the para style to each paragraph in cell
    var paraInCell = td.getChild(0).asParagraph();
    paraInCell.setAttributes(paraStyle);
    var td = tr.appendTableCell(Utilities.formatDate(ExcelDateToJSDate(dRow[11]), tz, "EEEE dd/MM/yyyy hh:mm a"));
    td.setAttributes(cellStyle);
    //Apply the para style to each paragraph in cell
    var paraInCell = td.getChild(0).asParagraph();
    paraInCell.setAttributes(paraStyle);
  }
}


##TABLE
UNDERLINE:null
LINK_URL:null
ITALIC:null
BOLD:null
BORDER_COLOR:#000000
BACKGROUND_COLOR:null
BORDER_WIDTH:1
FONT_SIZE:null
FONT_FAMILY:null
STRIKETHROUGH:null
FOREGROUND_COLOR:null
##ROWS
UNDERLINE:null
MINIMUM_HEIGHT:0
LINK_URL:null
ITALIC:null
BOLD:null
BACKGROUND_COLOR:null
FONT_SIZE:null
FONT_FAMILY:Arial
STRIKETHROUGH:null
FOREGROUND_COLOR:null
UNDERLINE:null
MINIMUM_HEIGHT:0
LINK_URL:null
ITALIC:null
BOLD:null
BACKGROUND_COLOR:null
FONT_SIZE:null
FONT_FAMILY:null
STRIKETHROUGH:null
FOREGROUND_COLOR:null
##CELLS
PADDING_TOP:5
UNDERLINE:null
VERTICAL_ALIGNMENT:Top
BOLD:true
BACKGROUND_COLOR:#cfe2f3
FONT_SIZE:null
FONT_FAMILY:Arial
STRIKETHROUGH:null
WIDTH:54.00000000000001
LINK_URL:null
ITALIC:null
PADDING_RIGHT:5
PADDING_BOTTOM:5
PADDING_LEFT:5
FOREGROUND_COLOR:#000000
PADDING_TOP:5
UNDERLINE:null
VERTICAL_ALIGNMENT:Top
BOLD:true
BACKGROUND_COLOR:#cfe2f3
FONT_SIZE:null
FONT_FAMILY:Arial
STRIKETHROUGH:null
WIDTH:null
LINK_URL:null
ITALIC:null
PADDING_RIGHT:5
PADDING_BOTTOM:5
PADDING_LEFT:5
FOREGROUND_COLOR:#000000
PADDING_TOP:5
UNDERLINE:null
VERTICAL_ALIGNMENT:Top
BOLD:null
BACKGROUND_COLOR:
FONT_SIZE:null
FONT_FAMILY:Consolas
STRIKETHROUGH:null
WIDTH:54.00000000000001
LINK_URL:null
ITALIC:null
PADDING_RIGHT:5
PADDING_BOTTOM:5
PADDING_LEFT:5
FOREGROUND_COLOR:null
PADDING_TOP:5
UNDERLINE:null
VERTICAL_ALIGNMENT:Top
BOLD:null
BACKGROUND_COLOR:
FONT_SIZE:null
FONT_FAMILY:Consolas
STRIKETHROUGH:null
WIDTH:null
LINK_URL:null
ITALIC:null
PADDING_RIGHT:5
PADDING_BOTTOM:5
PADDING_LEFT:5
FOREGROUND_COLOR:null
*/
