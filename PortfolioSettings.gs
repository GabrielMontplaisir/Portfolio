var portfolio = "Portfolio";

function PortfolioExists() {
  var portfolioSheet = SpreadsheetApp.getActive().getSheetByName(portfolio);
  var portfolioExists = false;
  if (portfolioSheet) {
    return portfolioExists = true
  }
}

function createPortfolioTab() {
  createPortfolioFolder();
  var ss = SpreadsheetApp.getActive();
  ss.insertSheet(portfolio);

  var portfolioSheet = ss.getSheetByName(portfolio);
  portfolioSheet.setFrozenRows(1);
  PropertiesService.getDocumentProperties().setProperty('IDCol', "Email Address");
  PropertiesService.getDocumentProperties().setProperty('commentCol', "Comment");
  portfolioSheet.getRange(1,1).setValue('Email Address');
  portfolioSheet.getRange(2,1,3).setValue('student@domain.com');
  portfolioSheet.getRange(1,2).setValue('Student Name');
  portfolioSheet.getRange(2,2,3).setValue('Smith, John');
  portfolioSheet.getRange(1,4).setValue('Portfolio URL');
  return "Created Portfolio tab and folder."
}

function getID() {return PropertiesService.getDocumentProperties().getProperty("IDCol")}
function getComment() {return PropertiesService.getDocumentProperties().getProperty('commentCol')}

function updateID() {
  var currentid = SpreadsheetApp.getCurrentCell().getValue();
  PropertiesService.getDocumentProperties().setProperty('IDCol', currentid);
  return currentid
}

function updateComment() {
  var currentComment = SpreadsheetApp.getCurrentCell().getValue();
  PropertiesService.getDocumentProperties().setProperty('commentCol', currentComment);
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
  var subFolderID = PropertiesService.getDocumentProperties().getProperty("docFolderID");
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
    PropertiesService.getDocumentProperties().setProperty("docFolderID", subFolderID);
    return subFolderID
  }
}