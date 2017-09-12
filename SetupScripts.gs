function runAll() {
  var start = new Date().getTime();
  //ID's are the random characters that are a part of the URL between /d/ and /edit
  //START THINGS TO CHANGE EVERY YEAR
  var folderId = ''; //Year Folder
  var fileId = ''; //Template NHS Points Sheet
  var memberListSheetId = ''; //NHS Members + Email Sheet
  var masterSheetId = ''; //Year Master Sheet - Must be changed yearly in the Template NHS Points Sheet
  var subFolders = ["Community", "School", "Tutoring", "New Point Sheets"]; //Sub-Categories
  //END THINGS TO CHANGE EVERY YEAR
  
  //!!DON'T TUCH BELOW HERE!!
  
  //START CREATING MEMBER LIST
  var sheet = SpreadsheetApp.open(DriveApp.getFileById(memberListSheetId)).getActiveSheet();
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2);
  var data = dataRange.getValues();
  var baseFolder = DriveApp.getFolderById(folderId);
  var nhsMembers = [];
  for (i in data) {
    var member = data[i][0];
    var iterator = baseFolder.getFoldersByName(member.toString());
    if (!(iterator.hasNext())) {
      if (member != null || member != "" || member != 0) { nhsMembers.push(member); }
    }
  }
  //END CREATING MEMBER LIST

  //START OPERATIONS
  createFoldersForEveryPerson(folderId, nhsMembers);
  createSubFolders(folderId, subFolders);
  copyPointSheets(folderId, nhsMembers, fileId);
  shareFolders(folderId, data, masterSheetId, memberListSheetId);
  //END OPERATIONS
  
  var end = new Date().getTime();
  var time = end - start;
  Logger.log("Execution time: " + time);
}

function createFoldersForEveryPerson(folderId, subFolders) { 
  var folder = DriveApp.getFolderById(folderId);
  var subFoldersLength = subFolders.length;
  for (var i = 0; i < subFoldersLength; i++) {
    Logger.log("Create Folder " + subFolders[i] + " in the Matter folder "+ folder.getName());
    var sub = folder.createFolder(subFolders[i]);   
  }
}

function createSubFolders(folderId, subFolders) {
  var folder = DriveApp.getFolderById(folderId);
  var clientFolders = folder.getFolders();
  while (clientFolders.hasNext()) {
    var clientFolder = clientFolders.next();
    var subFoldersLength = subFolders.length;
    var subFolder = clientFolder.getFolders();
    if (subFolder.hasNext() === false ) {
      for (var i = 0; i < subFoldersLength; i++) {
        Logger.log("Create Folder "+subFolders[i]+" in the Matter folder "+ clientFolder.getName());
        var sub = clientFolder.createFolder(subFolders[i]);   
      }
     } else {
       Logger.log("Folder "+ clientFolder.getName()+" already has children, so move on");
     }   
  }
}

function copyPointSheets(folderId, subFolders, fileId) {
  var folder = DriveApp.getFolderById(folderId);
  var file = DriveApp.getFileById(fileId);
  var subFoldersLength = subFolders.length;
  for (var i = 0; i < subFoldersLength; i++) {
    var subFolder = folder.getFoldersByName(subFolders[i]).next();
    file.makeCopy(subFolder);
    Logger.log("Copy points file in the Matter folder "+ subFolder.getName());
  }
}

function shareFolders(folderId, subFolders, masterSheetId, memberListSheetId) {
  var baseFolderToShare = DriveApp.getFolderById(folderId);
  var masterSheet = SpreadsheetApp.open(DriveApp.getFileById(masterSheetId)).getActiveSheet();
  var memberSheet = SpreadsheetApp.open(DriveApp.getFileById(memberListSheetId)).getActiveSheet();
  var counter = 2;
  var counter2 = 2;
  for (i in subFolders) {
     var row = subFolders[i][0]; //name
     var col = subFolders[i][1]; //email
     if ((row != null || row != "" || row != 0 ) && ( col != null || col != "" || col != 0)) {
       var rootIterator = baseFolderToShare.getFoldersByName(row.toString());
       while (rootIterator.hasNext()) {
         var iterator = rootIterator.next().getFoldersByName('New Point Sheets');
         while (iterator.hasNext()) {
           var folder = iterator.next();
           folder.addEditor(col);
           Logger.log('Shared with ' + row);
           masterSheet.getRange(counter, 1).setValue(row);
           memberSheet.getRange(counter, 3).setValue("DONE 1");
           counter++;
         }
         var rootIterator2 = baseFolderToShare.getFoldersByName(row.toString());
         var iterator2 = rootIterator2.next().getFilesByName("Copy of NHS Point Sheet");
         while (iterator2.hasNext()) {
           var folder = iterator2.next();
           folder.addViewer(col);
           Logger.log('Shared with ' + row);
           memberSheet.getRange(counter2, 4).setValue("DONE 2");
           counter2++;
         }
       }
     }
  }
}
