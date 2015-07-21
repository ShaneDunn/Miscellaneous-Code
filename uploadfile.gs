/* The script is deployed as a dialog and renders the form */
function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('up_load_file');
  SpreadsheetApp.getUi().showModalDialog(html, 'Load File');
}

/* This function will process the submitted form */
function uploadTheFile(theForm) {
  try {
    
    /* Name of the Drive folder where the files should be saved */
    var dropbox = "GDK";
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    
    /* Find the folder, create if the folder does not exist */
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    
    /* Get the file uploaded though the form as a blob */
    var fileBlob = theForm.theFile;
    var file = folder.createFile(fileBlob);    
    
    /* Set the file description as the name of the uploader */
    file.setDescription("Uploaded by " + theForm.myName);
        
    /* Return the download URL of the file once its on Google Drive */
    return "File uploaded successfully " + file.getUrl();
    
  } catch (error) {
    
    /* If there's an error, show the error message */
    return error.toString();
  }
  
}
