// The tab name for your portfolio.
var portfolio = getSheetbyId(PropertiesService.getDocumentProperties().getProperty('PortfolioSheet'));

// ID for template slide. TODO: Make it selectable by user.
var templateSlideID = "1NvIum_IB-2wUZSCeVFlcOTj5xV4e7cysBcnOifkhZj4"

// The column name to identify students. This should be a unique identifier, such as an email.
// var email = PropertiesService.getDocumentProperties().getProperty("IDCol");

// The column name for the student grades.
var comment = PropertiesService.getDocumentProperties().getProperty('commentCol');

function exportPortfolio() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();


  // Get values from Spreadsheet
  if (sh.getSheetId() != PropertiesService.getDocumentProperties().getProperty('PortfolioSheet')){
    var data = sh.getDataRange().getValues();
    // var emailIndex = data[0].indexOf(email); // Get Email column, selected by the user.
    var commentIndex = data[0].indexOf(comment);  // Get Comment column, selected by user.

    // Try to get the Responses to a linked form. If form is not linked, grab the responses from the sheet instead, between the email and the comment column.
    try {
      var formResponses = FormApp.openByUrl(sh.getFormUrl()).getResponses();
      // Logger.log("Found Form")
    } catch (e) {
      // Logger.log("Did not find form")
      // var formResponses = sh.getRange(1,emailIndex+2, sh.getLastRow(), commentIndex-2).getValues();
      return "Could not find a linked Google Form. Please link a form first, then try again."
    }

    // Place all values in an easily retrievable Array that will be passed to sortComments(). If there's an error, it's likely because the email and comment columns were not set properly.
    try {
      var studentComments = [
        // sh.getRange(2,emailIndex+1,sh.getLastRow()-1).getValues(),
        formResponses,
        sh.getRange(2,commentIndex+1,sh.getLastRow()-1).getValues()
      ];
      // Logger.log(studentComments)
    } catch (e) {
      // return "Identifier doesn't match or can't be found in portfolio. Select identifiers through the hamburger icon in the top right, or make sure they're correct in the portfolio tab."
      return "Comment Column can't be found in portfolio. Please select the Comment Column through the hamburger icon in the top right."
    }

    // Start sorting comments
    sortComments(studentComments);

  } else {
    // If the person is trying to export from the Portfolio tab, return this error.
    return "Cannot export from Portfolio Tab."
  }
  // If everything completed successfully, return this.
  return "Exported comments to Portfolios."
}


/* Function to create a portfolio for the student.
* Grab the student ID, name, and their row in the Portfolio tab.
* In theory, the row # is dependant on whether their ID matches the ID from the form. You could switch their spot, and it wouldn't affect the script in a negative way.
* If the script can't get the URL for the document (because it's in the trash or otherwise), the script will create a new portfolio for the person.
*/
function createStudentPortfolio(student, row){

  // Grab the Portfolio URLs
  var portfolioURL = portfolio.getRange(row+1,student.indexOf(student[1])+1);

  // Get Student's name. Requires the Admin SDK API.
  var name = AdminDirectory.Users.get(student[0], {viewType:'domain_public', fields:'name'});
  var fullName = name.name.fullName;
  // Logger.log(fullName);

  // Create folders for the User to put all the portfolios in one place. Calls on the createPortfolioFolder() function.
  var classFolderID = DriveApp.getFolderById(createPortfolioFolder());

  // Create a new Portfolio with the student and gives them editor access. Move it to the user's portfolio folder. **TODO: Make it selectable by the user?**
  var newPortfolio = SlidesApp
    .create(fullName+' Portfolio')
    .addEditor(student[0]);
  DriveApp.getFileById(newPortfolio.getId()).moveTo(classFolderID);

  // Set the Portfolio URL
  portfolioURL.setValue(newPortfolio.getUrl());



  // When you create a new Slide, the first slide will be the default. These next commands are to remove the default first slide, replace it with the first slide from the template, then name the Portfolio.
  newPortfolio.getSlides()[0].remove();
  newPortfolio.appendSlide(SlidesApp.openById(templateSlideID).getSlides()[0]);
  newPortfolio.getSlides()[0].replaceAllText("{{Portfolio Name}}", fullName+' Portfolio');

  // Return this prompt if everything worked.
  return portfolioURL.getValue();
}

/* 
* The main function to the whole operation. This is where the magic happens. I will do my best to break everything down.
*/
function sortComments(studentComments) {

  // The array loops I use here are more efficient from what I can tell, but give random values which can't be directly translated to a row or column integer. They only seem to work for the loop itself. I have to therefore create a variable for the row number.

  // Grab the Portfolio sheet, then get the Portfolio URLs, the StudentIDs and the Student Names.
  var data = portfolio.getDataRange().getValues();
  // Logger.log(data);

  // For all the students in the Portfolio Tab...
  for (var l = 1; l < data.length; l++) {
    // Logger.log(data[l][0]);
    // For all the students who filled a response...
    for (var s = 0; s < studentComments[0].length; s++){
      // Logger.log(studentComments[0][s].getRespondentEmail())
      if (studentComments[0][s].getRespondentEmail() == data[l][0].toString()) {
        // Logger.log(data[l]+' - '+studentComments[0][s]+' - '+studentComments[1][s]);

        try {
          var formResponse = studentComments[0][s].getItemResponses();
          Logger.log(formResponse)
        } catch (e) {
          Logger.log("Did not find Item Responses");
          // formResponse = studentComments[1][s];
        }
      
        try {
          // Logger.log(data[l][1])
          var studentPortfolio = SlidesApp.openByUrl(data[l][1].toString());
        } catch (e) {
          Logger.log("Cannot find student Portfolio. Creating new one.");
          var studentPortfolio = SlidesApp.openByUrl(createStudentPortfolio(data[l], l));
        }

        var currentSlide = studentPortfolio.appendSlide(SlidesApp.openById(templateSlideID).getSlides()[1]);
        currentSlide.replaceAllText("{{Title}}", "Commentaire de la réponse "+SpreadsheetApp.getActiveSheet().getName());
        var responsePlaceholderArray = [];
        for (var r in formResponse) {
          responsePlaceholderArray.push("{{Response "+r+"}}\n");
        }
        currentSlide.replaceAllText("{{Response}}", responsePlaceholderArray.join(''))
        for (var r in responsePlaceholderArray) {
          try {
            currentSlide.replaceAllText("{{Response "+r+"}}", "Question: "+formResponse[r].getItem().getTitle()+"\n"+formResponse[r].getResponse());
          } catch (e) {
            Logger.log("No form Response")
            // currentSlide.replaceAllText("{{Response "+r+"}}", "Question: "+studentComments[0][0][r]+"\n"+studentComments[0][s+1][r]);
          }
        }
        currentSlide.replaceAllText("{{Comment}}", studentComments[1][s]);

          

        // Logger.log(studentComments);
        break
      };
    };
  };
}