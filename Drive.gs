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
        Logger.log("Portfolio Found: "+portfolioFolderID)
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
        Logger.log("SubFolder Found: "+subFolderID)
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

function openDrive() {
  var portfolioFolder = PropertiesService.getUserProperties().getProperty("portfolioFolderID");
  return 'https://drive.google.com/drive/folders/'+portfolioFolder
}
