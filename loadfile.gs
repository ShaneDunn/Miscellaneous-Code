/** -------------------------------------------------------------------------- **/

var email = String(Session.getActiveUser().getEmail());

function doclistUI(){ // or function doGet(){  //** see text
  var folderlist = new Array();
  var folders = DocsList.getFolders()
  //Logger.log(folders);
  for(ff=0;ff<folders.length;++ff){
    //Logger.log(folders[ff]);
    folderlist.push(folders[ff].getName());
  }
  var app = UiApp.createApplication().setHeight(260).setWidth(700).setStyleAttribute('background', 'beige')
  .setStyleAttribute('margin-top', '10px').setStyleAttribute('margin-left', '10px').setTitle("Doclist UI");
  var panel = app.createVerticalPanel()
  var hpanel = app.createHorizontalPanel();
  var Flist= app.createListBox(false).setName("Flb").setId("Flb").setVisibleItemCount(4).setWidth("180");
  var Dlist= app.createListBox(false).setName("Dlb").setId("Dlb").setVisibleItemCount(8).setWidth("280");
  var hidden = app.createHidden('hidden').setId('hidden');
  var hiddenlink = app.createHidden('hiddenlink').setId('hiddenlink');
  var Flab = app.createLabel('Folder List').setWidth("80");
  var Dlab = app.createLabel('Document List').setWidth("100");
  var spacer = app.createLabel(' ').setWidth("30");

  Flist.addItem('Choose a folder').addItem('Root content');
  for(ff=0;ff<folderlist.length;++ff){
    Flist.addItem(folderlist[ff]);
    Logger.log(folderlist[ff]);
  }
  hpanel.add(Flab).add(Flist).add(spacer).add(Dlab).add(Dlist).add(hidden).add(hiddenlink);
  var docname = app.createLabel().setId('doc').setSize("360", "35");
  var link = app.createAnchor('open ', 'href').setId("link").setVisible(false);
  var load = app.createButton("Load meter data").setId("load").setVisible(false);
  var sent = app.createButton("Loading the Meter Data").setId("sent").setVisible(false);
  panel.add(hpanel).add(docname).add(link).add(load).add(sent);

  var clihandler = app.createClientHandler()
     .forTargets(sent).setVisible(true)
     .forEventSource().setVisible(false);
   load.addClickHandler(clihandler);

  var FHandler = app.createServerHandler("click");
  Flist.addChangeHandler(FHandler)
  FHandler.addCallbackElement(hpanel);

  var DHandler = app.createServerHandler("showlab");
  Dlist.addChangeHandler(DHandler);
  DHandler.addCallbackElement(hpanel);

  var keyHandler = app.createServerHandler("LoadData");
  load.addClickHandler(keyHandler)
  keyHandler.addCallbackElement(hpanel);

  app.add(panel);  
  var doc = SpreadsheetApp.getActive();//**
  doc.show(app);//**
}
/* ** if you want to publish the script as a service, rename this function as doget(e) ,  remove both lines marked with ** and replace them by the following :
  return app;
}
*/

/** -------------------------------------------------------------------------- **/

function click(e){
  var app = UiApp.getActiveApplication();
  var Dlist = app.getElementById("Dlb");  
  var hidden = app.getElementById("hidden");  
  var doclist = new Array();
  var label = app.getElementById('doc')
  var folderName = e.parameter.Flb

  if (folderName=='Choose a folder'){Dlist.clear();label.setText(" ");return app}
  if (folderName!='Root content'){
    doclist=DocsList.getFolder(folderName).getFiles(0,2000)
    var names = new Array();
    for (nn=0;nn<doclist.length;++nn){
      names.push([doclist[nn].getName(),doclist[nn].getId(),doclist[nn].getType()]);
    }
  }else{
    doclist=DocsList.getRootFolder().getFiles(0,2000)
    var names = new Array();
    for (nn=0;nn<doclist.length;++nn){
      names.push([doclist[nn].getName(),doclist[nn].getId(),doclist[nn].getType()]);
    }
  }
  names.sort();
  Dlist.clear();
  for(dd=0;dd<names.length;++dd){
    Dlist.addItem(names[dd][0]+" (doc Nr:"+dd+")");
  }
  hidden.setValue(names.toString());
  return app   ;// update UI
}
/** -------------------------------------------------------------------------- **/

function showlab(e){
  var app = UiApp.getActiveApplication();
  var label = app.getElementById('doc');
  var link = app.getElementById('link');
  var load = app.getElementById("load");
  
  var hidden = e.parameter.hidden.split(',');
  var hiddenlink = app.getElementById("hiddenlink");
  //Logger.log(hidden)
  if (e.parameter.Dlb!=""){
    var docname = e.parameter.Dlb
    var docN = docname.substr(0,docname.lastIndexOf("("));
    var docindex = docname.substring(Number(docname.lastIndexOf(":"))+1,Number(docname.lastIndexOf(")")));
    var doctype = hidden[docindex*3+2]
    //Logger.log(doctype)
    label.setText(doctype+" : "+docN).setEnabled(false).setStyleAttribute('fontSize', '15');
      if (doctype=='document'){var urlstring = "https://docs.google.com/document/d/";var poststring = "/edit"}
      if (doctype=='spreadsheet'){var urlstring = "https://docs.google.com/spreadsheet/ccc?key=";var poststring = "#gid=0"}
      if (doctype=='item'||doctype=='other'||doctype=='blob_item'||doctype=='photo'){var urlstring = "https://docs.google.com/file/d/";var poststring = "/edit"}

    var doclink = urlstring+hidden[docindex*3+1]+poststring;
    //Logger.log(doclink)
    link.setVisible(true).setText("Open "+doctype+" in browser").setHref(doclink); 
    load.setVisible(true);
    hiddenlink.setValue(hidden[docindex*3+1]+"|"+doctype);
  }
  return app ; // update UI
}

/** -------------------------------------------------------------------------- **/

/* 0B6NHem9C-Di5M3NEb1Y0TWJLcVU */

function importGDKCSV() {
  var searchTerm = "title = 'flocom-gdk-201504261711.prn'";
   
  // search for our file
  var files = DriveApp.searchFiles(searchTerm)
  var csvFile = "";
   
  // Loop through the results
  while (files.hasNext()) {
    var file = files.next();
    // assuming the first file we find is the one we want
    if (file.getName() == 'flocom-gdk-201504261711.prn') {
      // get file as a string
      csvFile = file.getBlob().getDataAsString();
      break;
    }
  }
  // parseCsv will return a [][] array we can write to a sheet
  var csvData = CSVToArray(csvFile, ",");
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var LOAD_LOG     = "Log";  // name of sheet to load logging information about each file loaded
  var lsheet = ss.getSheetByName(LOAD_LOG);
  var llastRow = lsheet.getLastRow();
  var DATA_SHEET   = "Data";  // name of sheet to load data from each file loaded
  var dsheet = ss.getSheetByName(DATA_SHEET);
  var dlastRow = dsheet.getLastRow();
  var EVENT_SHEET   = "Events";  // name of sheet to load event data from each file loaded
  var esheet = ss.getSheetByName(EVENT_SHEET);
  var elastRow = esheet.getLastRow();
  var OTHER_SHEET   = "Other";  // name of sheet to load unaccounted for data from each file loaded
  var osheet = ss.getSheetByName(OTHER_SHEET);
  var olastRow = osheet.getLastRow();

  var firstData = true;
  var data = [];
  var event = [];
  var other = [];
  var compDate = new Date();
  var lastTotal = 0.0;
  var flowRatepermin = 0.0;
  var interval = 0.0;
  var intervalflow = 0.0;
  var newtotal = 0.0;
  var newadjtotal = 0.0;
  var rowstatus = "";
  var cnt1 = 0;
  var cnt2 = 0;
  var cnt3 = 0;

  var lastTimeStamp = new Date(1970, 01, 01); // Time stamp of last event in data sheet

  if (dlastRow > 1) {
    data = dsheet.getRange(2,1,dlastRow-1,11).getValues();
  }
  if (elastRow > 1) {
    event = esheet.getRange(2,1,elastRow-1,3).getValues();
    lastTimeStamp = event[event.length-1][0];
  }

  Logger.log(llastRow);
  Logger.log(dlastRow + " : " + data.length);
  Logger.log(elastRow + " : " + event.length);
  Logger.log(lastTimeStamp);

  var logRow = new Array(20);
  logRow[0] = 'flocom-gdk-201504261711.prn';
  logRow[1] = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm:ss");

  for (var i = 0; i < csvData.length; i++) {
    var row = csvData[i];
    var checkString = row[0].slice(0,1);
    if ( checkString == "!" ) {
      logData( logRow, row ); // Comments
    } else {
      if ( row[1] == null ) {
        other.push(row) ; // unaccounted for
        Logger.log(row + " : " + i) ;
      } else {
        checkString = row[1].slice(0,1);
        compDate = stringToDate(row[0]);
        logRow[19] = row[0]; // Last date in file
        if ( firstData ) {
          logRow[18] = row[0]; // First date in file
          firstData = false;
        }
        cnt1++;
        if ( compDate > lastTimeStamp ) {
          cnt2++;
          if ( checkString == "*" ) {
            event.push(row) ; // events
            if ( row[1] == "*Flow stop" || row[1] == "*Flow start" || row[1] == "*Flow total" ) {
              lastTotal = parseFloat(row[2]);
              rowstatus = row[1].split(" ").pop();
            }
            if ( row[1] == "*Flow stop" ) {
              var lastrec = data.length - 1;
              interval = (compDate - data[lastrec][0]) / ( 1000 * 60 );
              intervalflow = data[lastrec][2] / (60 * 24) * interval;
              newadjtotal = data[lastrec][9] + intervalflow;
              data[lastrec][6] = interval;
              data[lastrec][7] = intervalflow;
              var rowData = new Array(compDate,
                                      rowstatus,
                                      0.0,
                                      adjustedDate(compDate, logRow[17]),
                                      lastTotal,
                                      0.0,
                                      interval,
                                      0.0,
                                      lastTotal,
                                      newadjtotal,
                                      "");
              data.push(rowData); // data
              rowstatus = "";
            }
          } else {
            flowRatepermin = parseFloat(row[1]) / (60 * 24);
            if (data.length > 0) {
              var lastrec = data.length - 1;
              interval = (compDate - data[lastrec][0]) / ( 1000 * 60 );
              intervalflow = data[lastrec][2] / (60 * 24) * interval;
              data[lastrec][6] = interval;
              data[lastrec][7] = intervalflow;

              if ( data[lastrec][8] == 0 ) {
                newtotal = 0;
                newadjtotal = 0;
              } else {
                newtotal = data[lastrec][8] + intervalflow;
                newadjtotal = data[lastrec][9] + intervalflow;
              }
              if ( rowstatus == "total" ) {
                newtotal = lastTotal;
                if ( data[lastrec][9] == 0 ) {
                  newadjtotal = lastTotal;
                }
              }
            } else {
              interval = logRow[15] / 60;
              intervalflow = flowRatepermin * interval;
              newtotal = 0;
              newadjtotal = 0;
            }
            rowData = new Array(compDate,
                                rowstatus,
                                parseFloat(row[1]),
                                adjustedDate(compDate, logRow[17]),
                                lastTotal,
                                flowRatepermin,
                                interval,
                                intervalflow,
                                newtotal,
                                newadjtotal,
                                "");
            data.push(rowData); // data
            rowstatus = "";
          }
        } else {
          cnt3++;
        }
      }
    }
  }
  if ( data.length > 0 ) {
    Logger.log((dlastRow+1) + " : " + data.length + " : " + data[0].length);
    dsheet.getRange(2,1,data.length,data[0].length).setValues(data);
  }
  if ( event.length > 0 ) {
    Logger.log((elastRow+1) + " : " + event.length + " : " + event[0].length);
    esheet.getRange(2,1,event.length,event[0].length).setValues(event);
  }
  if ( other.length > 0 ) {
    osheet.getRange(2,1,other.length,other[0].length).setValues(other);
  }
  var theLogRow = [];
  theLogRow[0] = logRow;
  lsheet.getRange(llastRow+1,1,1,logRow.length).setValues(theLogRow);
  Logger.log(cnt1 + " : " + cnt2 + " : " + cnt3);
}
/** -------------------------------------------------------------------------- **/

function LoadData(e) {
  var app = UiApp.getActiveApplication();
  var load = app.getElementById("load");
  var hiddenlink =  e.parameter.hiddenlink;
  var ID = hiddenlink.substring(0,Number(hiddenlink.lastIndexOf("|")));
  var sent = app.getElementById("sent");
  var fileBlob = DocsList.getFileById(ID).getBlob();
  var content = fileBlob.getDataAsString();
  var csvData = CSVToArray(content, ",");
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var LOAD_LOG     = "Log";  // name of sheet to load logging information about each file loaded
  var lsheet = ss.getSheetByName(LOAD_LOG);
  var llastRow = lsheet.getLastRow();
  var DATA_SHEET   = "Data";  // name of sheet to load data from each file loaded
  var dsheet = ss.getSheetByName(DATA_SHEET);
  var dlastRow = dsheet.getLastRow();
  var EVENT_SHEET   = "Events";  // name of sheet to load event data from each file loaded
  var esheet = ss.getSheetByName(EVENT_SHEET);
  var elastRow = esheet.getLastRow();
  var OTHER_SHEET   = "Other";  // name of sheet to load unaccounted for data from each file loaded
  var osheet = ss.getSheetByName(OTHER_SHEET);
  var olastRow = osheet.getLastRow();

  var firstData = true;
  var data = [];
  var event = [];
  var other = [];
  var compDate = new Date();
  var lastTotal = 0.0;
  var flowRatepermin = 0.0;
  var interval = 0.0;
  var intervalflow = 0.0;
  var newtotal = 0.0;
  var newadjtotal = 0.0;
  var rowstatus = "";
  var cnt1 = 0;
  var cnt2 = 0;
  var cnt3 = 0;

  var lastTimeStamp = new Date(1970, 01, 01); // Time stamp of last event in data sheet

  if (dlastRow > 1) {
    data = dsheet.getRange(2,1,dlastRow-1,11).getValues();
  }
  if (elastRow > 1) {
    event = esheet.getRange(2,1,elastRow-1,3).getValues();
    lastTimeStamp = event[event.length-1][0];
  }

  Logger.log(llastRow);
  Logger.log(dlastRow + " : " + data.length);
  Logger.log(elastRow + " : " + event.length);
  Logger.log(lastTimeStamp);

  var logRow = new Array(20);
  logRow[0] = DocsList.getFileById(ID).getName();
  logRow[1] = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm:ss");

  for (var i = 0; i < csvData.length; i++) {
    var row = csvData[i];
    var checkString = row[0].slice(0,1);
    if ( checkString == "!" ) {
      logData( logRow, row ); // Comments
    } else {
      if ( row[1] == null ) {
        other.push(row) ; // unaccounted for
        Logger.log(row + " : " + i) ;
      } else {
        checkString = row[1].slice(0,1);
        compDate = stringToDate(row[0]);
        logRow[19] = row[0]; // Last date in file
        if ( firstData ) {
          logRow[18] = row[0]; // First date in file
          firstData = false;
        }
        cnt1++;
        if ( compDate > lastTimeStamp ) {
          cnt2++;
          if ( checkString == "*" ) {
            event.push(row) ; // events
            if ( row[1] == "*Flow stop" || row[1] == "*Flow start" || row[1] == "*Flow total" ) {
              lastTotal = parseFloat(row[2]);
              rowstatus = row[1].split(" ").pop();
            }
            if ( row[1] == "*Flow stop" ) {
              var lastrec = data.length - 1;
              interval = (compDate - data[lastrec][0]) / ( 1000 * 60 );
              intervalflow = data[lastrec][2] / (60 * 24) * interval;
              newadjtotal = data[lastrec][9] + intervalflow;
              data[lastrec][6] = interval;
              data[lastrec][7] = intervalflow;
              var rowData = new Array(compDate,
                                      rowstatus,
                                      0.0,
                                      adjustedDate(compDate, logRow[17]),
                                      lastTotal,
                                      0.0,
                                      interval,
                                      0.0,
                                      lastTotal,
                                      newadjtotal,
                                      "");
              data.push(rowData); // data
              rowstatus = "";
            }
          } else {
            flowRatepermin = parseFloat(row[1]) / (60 * 24);
            if (data.length > 0) {
              var lastrec = data.length - 1;
              interval = (compDate - data[lastrec][0]) / ( 1000 * 60 );
              intervalflow = data[lastrec][2] / (60 * 24) * interval;
              data[lastrec][6] = interval;
              data[lastrec][7] = intervalflow;

              if ( data[lastrec][8] == 0 ) {
                newtotal = 0;
                newadjtotal = 0;
              } else {
                newtotal = data[lastrec][8] + intervalflow;
                newadjtotal = data[lastrec][9] + intervalflow;
              }
              if ( rowstatus == "total" ) {
                newtotal = lastTotal;
                if ( data[lastrec][9] == 0 ) {
                  newadjtotal = lastTotal;
                }
              }
            } else {
              interval = logRow[15] / 60;
              intervalflow = flowRatepermin * interval;
              newtotal = 0;
              newadjtotal = 0;
            }
            rowData = new Array(compDate,
                                rowstatus,
                                parseFloat(row[1]),
                                adjustedDate(compDate, logRow[17]),
                                lastTotal,
                                flowRatepermin,
                                interval,
                                intervalflow,
                                newtotal,
                                newadjtotal,
                                "");
            data.push(rowData); // data
            rowstatus = "";
          }
        } else {
          cnt3++;
        }
      }
    }
  }
  if ( data.length > 0 ) {
    Logger.log((dlastRow+1) + " : " + data.length + " : " + data[0].length);
    dsheet.getRange(2,1,data.length,data[0].length).setValues(data);
  }
  if ( event.length > 0 ) {
    Logger.log((elastRow+1) + " : " + event.length + " : " + event[0].length);
    esheet.getRange(2,1,event.length,event[0].length).setValues(event);
  }
  if ( other.length > 0 ) {
    osheet.getRange(2,1,other.length,other[0].length).setValues(other);
  }
  var theLogRow = [];
  theLogRow[0] = logRow;
  lsheet.getRange(llastRow+1,1,1,logRow.length).setValues(theLogRow);
  Logger.log(cnt1 + " : " + cnt2 + " : " + cnt3);
  sent.setVisible(false);
  return app ;  // update UI 
} // function LoadData
