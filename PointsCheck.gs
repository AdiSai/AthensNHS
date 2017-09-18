function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Points Check')
      .addItem('Run', 'run')
      .addToUi();
}

function run() {
  var emailAddress = "athensnhspoints@gmail.com";
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange(52, 8, 4, 1);
  var data = dataRange.getValues();
  updateMaster(data);
  for (i in data) {
    var row = data[i];
    var message = "";
    var subject = "Points Reached";
    if (i == 0) {
      message = "Points for the entire year in school have been reached";
      if (row == 10.0) { MailApp.sendEmail(emailAddress, subject, message); }
    }
    if (i == 1) {
      message = "Points for the entire year in community have been reached";
      if (row == 10.0) { MailApp.sendEmail(emailAddress, subject, message); }
    }
    if (i == 2) {
      message = "Points for the entire year in tutoring have been reached";
      if (row == 10.0) { MailApp.sendEmail(emailAddress, subject, message); }
    }
    if (i == 3) {
      message = "Points for the entire year have been reached";
      if (row == 40.0) { MailApp.sendEmail(emailAddress, subject, message); }
    }
  }
}

function updateMaster(data) {
  //ID's are the random characters that are a part of the URL between /d/ and /edit
  var sheetId = ''; //Master Sheet ID - CHANGE YEARLY
  var baseFolder = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents().next(); //Person's Name
  for (i in data) {
    if (data[i] == null || data[i] == "") { data[i] = 0; }
  }
  var school = data[0];
  var community = data[1];
  var tutoring = data[2];
  var total = data[3];
  Logger.log(baseFolder);
  var masterSheet = SpreadsheetApp.open(DriveApp.getFileById(sheetId)).getActiveSheet();
  var dataRange = masterSheet.getRange(1, 1, 150, 6);
  var values = dataRange.getValues();
  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] == baseFolder) {
        masterSheet.getRange(i + 1, j + 3).setValue(school);
        masterSheet.getRange(i + 1, j + 4).setValue(community);
        masterSheet.getRange(i + 1, j + 5).setValue(tutoring);
        masterSheet.getRange(i + 1, j + 6).setValue(total);
      }
    }    
  } 
}
