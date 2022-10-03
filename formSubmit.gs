function formSubmit(e) {
  ss = SpreadsheetApp.getActive();
  var itemResponses = e.values;
  Logger.log(itemResponses);
  var email = itemResponses[1];
  var portfolioSheet = getSheetbyId(PropertiesService.getDocumentProperties().getProperty('PortfolioSheet'));

  if (portfolioSheet.createTextFinder(email).matchEntireCell(true).findNext()) {
    return
  }

  portfolioSheet.getRange(portfolioSheet.getLastRow()+1,1).setValue(email);
}
