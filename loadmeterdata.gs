function doGet(e) {  
  var app = UiApp.createApplication().setTitle("Upload Flocom Data").setHeight(100).setWidth(300);
  var formContent = app.createGrid(3,1);
  formContent.setWidget(0,0,app.createFileUpload().setName('thefile'));
  formContent.setWidget(2,0,app.createSubmitButton("Read Flocom Data"));
  var form = app.createFormPanel();
  form.add(formContent);
  app.add(form);
  SpreadsheetApp.getActive().show(app);
  return app;
}

function doPost(e) {
  var WATER_DATA = "Flocom_Data"; // name of sheet to load water meter data
  var app = UiApp.getActiveApplication();  
  var fileBlob = e.parameter.thefile;
  var content = fileBlob.getDataAsString();// no need to create an intermediate file
  var csvData = CSVToArray(content, ",");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getActiveSheet();
  var sheet = ss.getSheetByName(WATER_DATA);
  
  Logger.log(csvData.length);
  Logger.log(csvData[0].length);
  
  //sheet.getRange(1,1,csvData.length, csvData[0].length).setValues(csvData);
  //sheet.getRange(1,1,1000, 14).setValues(csvData);
  for (var i = 0; i <= csvData.length - 1; i++) {
    var row = csvData[i];
    // Logger.log(row);
    for (var j = 0; j <= row.length - 1; j++) {
      //var value = row[i];
      sheet.getRange(i+1,j+1,1,1).setValue(row[j]);
    }
  }

  return app.close();
}

function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};


function CSVToArray( strData, strDelimiter ){
  //Logger.log(strData);
  strDelimiter = (strDelimiter || ",");
  var objPattern = new RegExp(
    (
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );
  var arrData = [[]];
  var arrMatches = null;
  while (arrMatches = objPattern.exec( strData )){
    var strMatchedDelimiter = arrMatches[ 1 ];
    if (
      strMatchedDelimiter.length &&
      (strMatchedDelimiter != strDelimiter)
    ){
      arrData.push( [] );
    }
    if (arrMatches[ 2 ]){
      var strMatchedValue = arrMatches[ 2 ].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );
    } else {
      var strMatchedValue = arrMatches[ 3 ];
    }
    arrData[ arrData.length - 1 ].push( strMatchedValue );
  }
  return( arrData );
} 

function listFilesInFolder() {
  var FOLDER_NAME = "GDK"; // folder name of where to put completed reports
  var FOLDER_ID = "0B6NHem9C-Di5XzlfVGRzRzVtbU0"; // folder name of where to put completed reports
  var MAX_FILES = 20; //use a safe value, don't be greedy
  var arnames=[];
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastExecution = scriptProperties.getProperty('LAST_EXECUTION');
  if( lastExecution === null )
    lastExecution = '';

  var continuationToken = scriptProperties.getProperty('IMPORT_ALL_FILES_CONTINUATION_TOKEN');
  var iterator = continuationToken == null ?
    DriveApp.getFolderById(FOLDER_ID).getFiles() : DriveApp.continueFileIterator(continuationToken);


  try { 
    for( var i = 0; i < MAX_FILES && iterator.hasNext(); ++i ) {
      var file = iterator.next();
      // var dateCreated = formatDate(file.getDateCreated());
      var dateCreated = file.getDateCreated().getTime();
      // var currentExecution = new Date().getTime();
      // scriptProperties.setProperty('LAST_EXECUTION',currentExecution);

      // if(dateCreated > lastExecution) {
        processFile(file);
        arnames.push([file.getName(),file.getId()]);
      // }
    }
  } catch(err) {
    Logger.log(err);
  }

  if( iterator.hasNext() ) {
    scriptProperties.setProperty('IMPORT_ALL_FILES_CONTINUATION_TOKEN', iterator.getContinuationToken());
  } else { // Finished processing files so delete continuation token
    scriptProperties.deleteProperty('IMPORT_ALL_FILES_CONTINUATION_TOKEN');
    scriptProperties.setProperty('LAST_EXECUTION', formatDate(new Date()));
  }
  Logger.log(arnames);
}

function formatDate(date) { return Utilities.formatDate(date, "GMT", "yyyy-MM-dd HH:mm:ss"); }

function processFile(file) {
  var id = file.getId();
  var name = file.getName();

  //your processing...
  Logger.log(name);
  Logger.log(id);
  Logger.log(file.getMimeType());
  Logger.log(file.getDateCreated());
}

function doGet1() {
  var app = UiApp.createApplication();
  var panel = app.createVerticalPanel();
  var listBox = app.createListBox().setName('myList').setWidth('80px');
  listBox.addItem('Item 1');
  listBox.addItem('Item 2');
  listBox.addItem('Item 3');
  listBox.addItem('Item 4');
  
  //Add a handler to the ListBox when its value is changed
  var handler = app.createServerChangeHandler('showSelectedinfo');
  handler.addCallbackElement(listBox);
  listBox.addChangeHandler(handler);
  var infoLabel = app.createLabel('Select from List').setId('info');
  panel.add(listBox);
  panel.add(infoLabel);
  app.add(panel);
  return app; 
}

//This functions displays the infolabel with when ListBox value is changed
function showSelectedinfo(e){
  var app = UiApp.getActiveApplication();
  app.getElementById('info').setText('You selected :'+e.parameter.myList).setVisible(true)
    .setStyleAttribute('color','#008000');
  return app;
}
