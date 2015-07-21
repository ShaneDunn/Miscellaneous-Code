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


function loadtheData (dataRow, theSheet) {
  var THE_SHEET = theSheet; // name of sheet to load water meter data
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet  = ss.getSheetByName(THE_SHEET);
  var lastRow = sheet.getLastRow() + 1;
  if ( dataRow.length > 2 ) {
  }
  else {
    //sheet.getRange(lastRow,1,1,2).setValues(dataRow);
    for (var j = 0; j <= dataRow.length - 1; j++) {
      sheet.getRange(lastRow,j+1,1,1).setValue(dataRow[j]);
    }

  }
  
/*
"Low battery"
"Battery normal"
"Host comms. start"
"Host comms. end"
"Unit started"
"Unit stopped"
"Unit restarted"
"No solar power"
"Solar power restored"
"Battery was flat"
"Head detected"
"Head not detected"

*Note
*Flow total
*Flow start
*Flow stop
*Missed meas
*Mace
*Temperature


  */
}

function logData( plogRow, pdataRow ) {
  //Logger.log(plogRow);
  switch (pdataRow[0]) {
    case "!":
      break;

    case "!Program":
      plogRow[2] = pdataRow[1] + " - " + pdataRow[2]; //,FloCom,2.6.1.9
      break;

    case "!Download Start":
      plogRow[3] = pdataRow[1]; //,2015/01/05 20:33:58
      break;

    case "!Ident":
      plogRow[4] = pdataRow[1]; //,"1-228-1"
      break;

    case "!SerialNo":
      plogRow[5] = pdataRow[1]; //,"14553"
      break;

    case "!Version":
      plogRow[6] = pdataRow[1] + " - " + pdataRow[2]; //,313,167
      break;

    case "!Logger Time":
      plogRow[7] = pdataRow[1]; //,2015/01/05 19:58:38
      plogRow[17] = stringToDate(plogRow[3]) - stringToDate(plogRow[7]); // time difference for time correction
      break;

    case "!Battery":
      plogRow[8] = pdataRow[1] + " - " + pdataRow[2]; //,"OK",6.44V
      break;

    case "!Logger temperature":
      plogRow[9] = pdataRow[1]; //,30
      break;

    case "!Total flow units":
      plogRow[10] = pdataRow[1]; //,"Ml"
      break;

    case "!Channels":
      plogRow[11] = pdataRow[1]; //,1
      break;

    case "!Names":
      plogRow[12] = pdataRow[1]; //,"Flow rate"
      break;

    case "!Units":
      plogRow[13] = pdataRow[1]; //,"Ml/day"
      break;

    case "!Points":
      plogRow[14] = pdataRow[1]; //,22525
      break;

    case "!Interval":
      plogRow[15] = pdataRow[1]; //,600
      break;
      
    case "!Download End":
      plogRow[16] = pdataRow[1]; //,2015/01/05 20:35:17
      break;

    default:
      Logger.log(pdataRow);
      break;
  }
}

var toType = function(obj) {
  return ({}).toString.call(obj).match(/\s([a-zA-Z]+)/)[1].toLowerCase()
}

function stringToDate(dateString) {
  var dateTimeArray = dateString.split(" ");
  var dateArray = dateTimeArray[0].split("/");
  var vyear = Number(dateArray[0]);
  var vmonth = Number(dateArray[1]) - 1;
  var vday = Number(dateArray[2]);
  var timeArray = dateTimeArray[1].split(":");
  var vhour = Number(timeArray[0]);
  var vminute = Number(timeArray[1]);
  var vsecond = Number(timeArray[2]);
  if ( vyear == Number.NaN ||
       vmonth == Number.NaN ||
       vday == Number.NaN ||
       vhour == Number.NaN ||
       vminute == Number.NaN ||
       vsecond == Number.NaN ) {
    Logger.log(vyear + ' | ' + vmonth + ' | ' + vday + ' | ' + vhour + ' | ' + vminute + ' | ' + vsecond + ' | ' + dateTimeArray + ' | ' + dateArray + ' | ' + timeArray);
  }
  var vdate = new Date(vyear, vmonth, vday, vhour, vminute, vsecond);
  //if (date instanceof Date) {
  if ( toType(vdate) === "date" ) {
    return vdate;
  } else {
    Logger.log(vdate);
    Logger.log(dateString);
    Logger.log("non-compliant Date object detected!");
  }
}


function adjustedDate(originalDate, adjustement) {
  var vtime = originalDate.getTime() + adjustement;
  var vdate = new Date();
  vdate.setTime(vtime);
  return vdate;
}
