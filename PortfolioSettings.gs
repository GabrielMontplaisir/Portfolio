function getSheetbyId(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId().toString() === id;}
  )[0];
}

// function getID() {return PropertiesService.getScriptProperties().getProperty("IDCol")}
function getComment() {return PropertiesService.getScriptProperties().getProperty('commentCol')}

// function updateID() {
//   var currentid = SpreadsheetApp.getCurrentCell().getValue();
//   PropertiesService.getScriptProperties().setProperty('IDCol', currentid);
//   return currentid
// }

function updateComment() {
  var currentComment = SpreadsheetApp.getCurrentCell().getValue();
  PropertiesService.getScriptProperties().setProperty('commentCol', currentComment);
  return currentComment
}

function getDocProps() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Document Properties');
  let uObj=PropertiesService.getDocumentProperties().getProperties();
  let keys = Object.keys(uObj);
  sh.clearContents();
  let a=[['Key','Value']];
  keys.forEach(k => {a.push([k,uObj[k]]);});
  sh.getRange(1,1,a.length, a[0].length).setValues(a);
  ss.toast('Document Properties generated.')
}

function delDocProps() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
}

function createPortfolioFolder() {
  var userFolders = DriveApp.getFolders();
  var portfolioFolderExists = false;
  var portfolioFolderID = PropertiesService.getUserProperties().getProperty("portfolioFolderID");
  
  // Check if folder already exists.
  while(userFolders.hasNext()){
    var folder = userFolders.next();

    //If the name exists return the id of the folder
    try {
      if(folder.getId() === portfolioFolderID){
        Logger.log("Portfolio Found")
        portfolioFolderExists = true;
        return checkSubFolders();
      };
    } catch(e) {
      continue
    };
  };

  if (!portfolioFolderExists) {
    Logger.log("No Portfolio folder found")
    var portfolioFolder = DriveApp.createFolder("Portfolio").getId();
    PropertiesService.getUserProperties().setProperty("portfolioFolderID", portfolioFolder);
    return checkSubFolders();
  }
}

function checkSubFolders() {
  var ssName = SpreadsheetApp.getActive().getName();
  var subFolderExists = false;
  var portfolioFolder = DriveApp.getFolderById(PropertiesService.getUserProperties().getProperty("portfolioFolderID"));
  var userSubFolders = portfolioFolder.getFolders();
  var subFolderID = PropertiesService.getUserProperties().getProperty("docFolderID");
  while (userSubFolders.hasNext()) {
    var subFolder = userSubFolders.next();

    try {
      if (subFolder.getId() === subFolderID) {
        Logger.log("SubFolder Found")
        subFolderExists = true;
        return subFolderID
      }
    } catch (e) {
      continue
    };

  }

  if (!subFolderExists) {
    Logger.log("Subfolder not found")
    subFolderID = portfolioFolder.createFolder(ssName).getId();
    PropertiesService.getUserProperties().setProperty("docFolderID", subFolderID);
    return subFolderID
  }
}