function createForm() {
  var newForm = FormApp.create('New Portfolio Form')
    .setCollectEmail(true)
    .setLimitOneResponsePerUser(true);
  var ss = SpreadsheetApp.getActive();
  newForm.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  DriveApp.getFileById(newForm.getId()).moveTo(DriveApp.getFolderById(PropertiesService.getUserProperties().getProperty("docFolderID")))
  return newForm.getEditUrl()
}

function importForm(formData) {
  var ss = SpreadsheetApp.getActive();
  formData.forEach(function(form) {
    var formID = FormApp.openById(form);
    formID.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())

    var formName = formID.getTitle();
    var formURL = formID.getEditUrl().replace('/edit','/viewform');
    Logger.log(formName+' - '+formURL)
    
    // Find the relevant Sheet to the Form...
    var sheets = ss.getSheets().filter(function(sh){
      Logger.log(sh.getFormUrl())
      return sh.getFormUrl() === formURL
    });
    // Rename the Tab to the Form Name.
    sheets[0].setName(formName);
  });
}