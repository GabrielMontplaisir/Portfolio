function createForm() {
  var newForm = FormApp.create('New Portfolio Form')
    .setCollectEmail(true)
    .setLimitOneResponsePerUser(true);
  var ss = SpreadsheetApp.getActive();
  newForm.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  DriveApp.getFileById(newForm.getId()).moveTo(DriveApp.getFolderById(PropertiesService.getDocumentProperties().getProperty("docFolderID")))
  return newForm.getEditUrl()
}

function importForm(formData) {
  ss = SpreadsheetApp.getActive();
  formData.forEach(function(form) {
    var formID = FormApp.openById(form);
    formID.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())
  });
}